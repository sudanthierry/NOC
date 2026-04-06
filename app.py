import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
from datetime import datetime

# --- CONFIGURACOES DA PAGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide")

# --- 1. FUNCOES DE CONEXAO E GOOGLE SHEETS ---
def conectar_google():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    else:
        try:
            creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        except Exception:
            st.error("Erro: Credenciais do Google nao encontradas.")
            st.stop()
    client = gspread.authorize(creds)
    return client.open("noc_config").worksheet("blacklist")

def carregar_blacklist_df():
    try:
        wks = conectar_google()
        data = wks.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame(columns=["Device Name", "Motivo", "NOC"])

# --- 2. LOGICA DE CALCULO DE SLA ---
@st.cache_resource
def get_holidays():
    h = holidays.BR(state='PR', years=range(2024, 2030))
    for y in range(2024, 2030):
        h.append({f"{y}-09-08": "Nossa Senhora da Luz - Curitiba"})
    return h

br_holidays = get_holidays()

def analyze_downtime(start, end):
    if pd.isnull(start) or pd.isnull(end) or start >= end: 
        return 0.0
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    for day in days:
        if day.weekday() >= 5 or day in br_holidays:
            continue 
        work_start = day.replace(hour=8, minute=0, second=0)
        work_end = day.replace(hour=18, minute=0, second=0)
        actual_start = max(start, work_start)
        actual_end = min(end, work_end)
        if actual_start < actual_end:
            total_minutes += (actual_end - actual_start).total_seconds() / 60
    return total_minutes

def format_hms(m):
    ts = int(m * 60)
    return f"{ts // 3600:02d}:{(ts % 3600) // 60:02d}:{ts % 60:02d}"

# --- 3. INTERFACE ---
st.title("NOC SLA Analyser - Business Hours")

# --- PROCESSAMENTO ---
file_main = st.file_uploader("Upload do arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        # 1. Pula as 8 primeiras linhas
        df = pd.read_excel(file_main, skiprows=8)
        
        # 2. Remove as 3 ultimas linhas
        if len(df) > 3:
            df = df.iloc[:-3]
        
        # 3. PREENCHIMENTO DOS CAMPOS EM BRANCO (ffill)
        # O ffill() propaga o nome do device para as linhas de baixo que estiverem vazias
        df['Device Name'] = df['Device Name'].ffill()
        df['Downtime Start'] = df['Downtime Start'].ffill()
        df['Downtime End'] = df['Downtime End'].ffill()

        # 4. LIMPEZA DE NOME (EXT.TEXTO + PROCURAR '(' )
        df['Device Name'] = df['Device Name'].astype(str).str.split('(').str[0].str.strip()

        # 5. CONVERSAO DE DATAS
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'].astype(str).str.strip(), dayfirst=True, errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'].astype(str).str.strip(), dayfirst=True, errors='coerce')

        # 6. FILTRO BLACKLIST
        df_bl = carregar_blacklist_df()
        if not df_bl.empty:
            ignorados = [str(x).strip().upper() for x in df_bl['Device Name'].tolist()]
            df = df[~df['Device Name'].str.upper().isin(ignorados)]

        with st.spinner('Processando dados...'):
            df['Minutos_Comerciais'] = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End']), axis=1)
            
            # Remove quedas fora do horario comercial/fds (Tempo = 0)
            df = df[df['Minutos_Comerciais'] > 0].copy()
            
            if df.empty:
                st.warning("Nenhuma queda valida encontrada.")
            else:
                df['Tempo_SLA'] = df['Minutos_Comerciais'].apply(format_hms)

                # Regras de SLA
                c_ap = df['Device Name'].str.contains('AP', case=False, na=False)
                c_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
                
                df_final = df[
                    ((c_ap) & (df['Minutos_Comerciais'] >= 240)) | 
                    ((c_wni) & (df['Minutos_Comerciais'] >= 360)) | 
                    ((~c_ap) & (~c_wni) & (df['Minutos_Comerciais'] >= 10))
                ].copy()

                st.success(f"Analise concluida: {len(df_final)} violacoes.")
                st.dataframe(df_final.drop(columns=['Minutos_Comerciais']), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("Baixar Relatorio Final", output.getvalue(), "SLA_Final.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
