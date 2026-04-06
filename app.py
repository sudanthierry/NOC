import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
from datetime import datetime

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide", page_icon="???")

# --- 1. FUNÇÕES DE CONEXÃO E GOOGLE SHEETS ---
def conectar_google():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    else:
        try:
            creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        except Exception:
            st.error("Erro: Credenciais do Google não encontradas.")
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

def adicionar_a_blacklist(nome_device, motivo_texto, noc_selecionado):
    try:
        wks = conectar_google()
        nomes_no_sheet = [str(n).strip().upper() for n in wks.col_values(1)[1:]]
        if nome_device.strip().upper() in nomes_no_sheet:
            st.error(f"?? O dispositivo '{nome_device}' já existe na Blacklist.")
            return False
        wks.append_row([nome_device.strip(), motivo_texto.strip(), noc_selecionado])
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na nuvem: {e}")
        return False

# --- 2. LÓGICA DE CÁLCULO DE SLA ---
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
    
    # Range de dias
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    
    for day in days:
        # Se for Sábado (5), Domingo (6) ou Feriado, o tempo desse dia é ZERO
        if day.weekday() >= 5 or day in br_holidays:
            continue 
        
        # Janela comercial 08:00 - 18:00
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
st.title("?? NOC SLA Analyser - Business Hours Only")

with st.sidebar:
    st.header("??? Gestão de Blacklist")
    with st.form("form_exclusao", clear_on_submit=True):
        nome_input = st.text_input("Nome do Equipamento:")
        lista_nocs = ["SME", "Leste", "Matriz", "Norte", "Oeste", "Sul"]
        noc_input = st.selectbox("Setor:", lista_nocs)
        motivo_input = st.text_area("Justificativa:")
        if st.form_submit_button("Salvar na Nuvem"):
            if nome_input and motivo_input:
                if adicionar_a_blacklist(nome_input, motivo_input, noc_input):
                    st.success("Adicionado!")
                    st.rerun()

    st.divider()
    df_bl = carregar_blacklist_df()
    if not df_bl.empty:
        st.dataframe(df_bl, use_container_width=True, hide_index=True)

# --- PROCESSAMENTO ---
file_main = st.file_uploader("Upload DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        df = pd.read_excel(file_main, skiprows=8)
        df[['Device Name', 'Downtime Start', 'Downtime End']] = df[['Device Name', 'Downtime Start', 'Downtime End']].ffill()
        
        # Filtro Blacklist (Case Insensitive)
        if not df_bl.empty:
            ignorados = [str(x).strip().upper() for x in df_bl['Device Name'].tolist()]
            df = df[~df['Device Name'].astype(str).str.strip().str.upper().isin(ignorados)]

        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'], errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'], errors='coerce')

        with st.spinner('Filtrando finais de semana e calculando minutos comerciais...'):
            # Calcula os minutos comerciais
            df['Minutos_Comerciais'] = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End']), axis=1)
            
            # --- CORREÇÃO AQUI: REMOVE QUALQUER QUEDA QUE RESULTE EM 0 MINUTOS COMERCIAIS ---
            # Isso elimina quedas de fim de semana, feriados e madrugadas antes de qualquer outra regra.
            df = df[df['Minutos_Comerciais'] > 0].copy()
            
            if df.empty:
                st.warning("Nenhuma queda registrada ocorreu dentro do horário comercial ou em dias úteis.")
            else:
                df['Tempo_SLA'] = df['Minutos_Comerciais'].apply(format_hms)

                # Regras de Corte de SLA
                c_ap = df['Device Name'].str.contains('AP', case=False, na=False)
                c_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
                
                df_final = df[
                    ((c_ap) & (df['Minutos_Comerciais'] >= 240)) | 
                    ((c_wni) & (df['Minutos_Comerciais'] >= 360)) | 
                    ((~c_ap) & (~c_wni) & (df['Minutos_Comerciais'] >= 10))
                ].copy()

                st.success(f"Análise concluída: {len(df_final)} violações válidas.")
                st.dataframe(df_final.drop(columns=['Minutos_Comerciais']), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("?? Baixar Relatório", output.getvalue(), "SLA_Final.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
