import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
import warnings
import pytz
from datetime import datetime

# Ignorar avisos de estilo do openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

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
            st.error("Erro: Credenciais do Google nao encontradas no secrets ou localmente.")
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
            st.error(f"O equipamento '{nome_device}' ja existe na Blacklist.")
            return False
        wks.append_row([nome_device.strip(), motivo_texto.strip(), noc_selecionado])
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na planilha: {e}")
        return False

# --- 2. LOGICA DE CALCULO ---
@st.cache_resource
def get_holidays():
    h = holidays.BR(state='PR', years=range(2024, 2030))
    for y in range(2024, 2030):
        h.append({f"{y}-09-08": "Nossa Senhora da Luz - Curitiba"})
    return h

def analyze_downtime_comercial(start, end, feriados):
    if pd.isnull(start) or pd.isnull(end) or start >= end: 
        return 0.0
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    for day in days:
        if day.weekday() >= 5 or day in feriados:
            continue 
        work_start = day.replace(hour=8, minute=0, second=0, microsecond=0)
        work_end = day.replace(hour=18, minute=0, second=0, microsecond=0)
        actual_start = max(start, work_start)
        actual_end = min(end, work_end)
        if actual_start < actual_end:
            total_minutes += (actual_end - actual_start).total_seconds() / 60
    return total_minutes

def format_hms(m):
    ts = int(m * 60)
    return f"{ts // 3600:02d}:{(ts % 3600) // 60:02d}:{ts % 60:02d}"

# --- 3. INTERFACE (SIDEBAR) ---
st.title("NOC SLA Analyser")

with st.sidebar:
    st.header("Gestao de Blacklist")
    
    with st.form("form_exclusao", clear_on_submit=True):
        st.subheader("Nova Excecao")
        nome_input = st.text_input("Nome do Equipamento:")
        lista_nocs = ["SME", "Leste", "Matriz", "Norte", "Oeste", "Sul"]
        noc_input = st.selectbox("Designar NOC:", lista_nocs)
        motivo_input = st.text_area("Justificativa:")
        if st.form_submit_button("Salvar na Nuvem"):
            if nome_input and motivo_input:
                if adicionar_a_blacklist(nome_input, motivo_input, noc_input):
                    st.success("Adicionado!")
                    st.rerun()
            else:
                st.warning("Preencha Nome e Motivo.")

    st.divider()
    st.subheader("Equipamentos Ignorados")
    df_bl = carregar_blacklist_df()
    if not df_bl.empty:
        st.dataframe(df_bl, width='stretch', hide_index=True)
    else:
        st.info("Nenhuma excecao na lista.")

# --- 4. PROCESSAMENTO PRINCIPAL ---
file_main = st.file_uploader("Upload do arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        # A. Limpeza de linhas (skip 8, remove last 5)
        df = pd.read_excel(file_main, skiprows=8)
        if len(df) > 5:
            df = df.iloc[:-5]
        
        # B. Preenchimento (ffill)
        cols_fill = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration']
        df[cols_fill] = df[cols_fill].ffill()

        # C. Tratamento Currently Down (Horario Brasilia)
        tz_br = pytz.timezone('America/Sao_Paulo')
        agora_br = datetime.now(tz_br).replace(tzinfo=None)
        df['Downtime End'] = df['Downtime End'].astype(str).replace('Currently Down', agora_br.strftime('%Y-%m-%d %H:%M:%S'))

        # D. Filtro Reason (apenas vazios)
        if 'Reason' in df.columns:
            df = df[df['Reason'].isna() | (df['Reason'].astype(str).str.strip() == "")].copy()

        # E. Split Nome e Conversao Datas
        df['Device Name'] = df['Device Name'].astype(str).str.split('(').str[0].str.strip()
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'].astype(str).str.strip(), dayfirst=True, errors='coerce', format='mixed')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'].astype(str).str.strip(), dayfirst=True, errors='coerce', format='mixed')

        # F. Aplicar Filtro Blacklist
        if not df_bl.empty:
            ignorados = [str(x).strip().upper() for x in df_bl['Device Name'].tolist()]
            df = df[~df['Device Name'].str.upper().isin(ignorados)]

        with st.spinner('Processando dados...'):
            feriados = get_holidays()
            is_ap = df['Device Name'].str.contains('AP', case=False, na=False)
            is_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
            
            # Calculo hibrido: Comercial para AP/WNI, 24/7 para outros
            df.loc[is_ap | is_wni, 'Minutos_SLA'] = df[is_ap | is_wni].apply(
                lambda r: analyze_downtime_comercial(r['Downtime Start'], r['Downtime End'], feriados), axis=1
            )
            df.loc[~(is_ap | is_wni), 'Minutos_SLA'] = ((df['Downtime End'] - df['Downtime Start']).dt.total_seconds() / 60).fillna(0)

            # Filtros de Corte
            c_ap = (is_ap) & (df['Minutos_SLA'] >= 240)
            c_wni = (is_wni) & (df['Minutos_SLA'] >= 360)
            c_out = (~is_ap) & (~is_wni) & (df['Minutos_SLA'] >= 10)
            
            df_final = df[c_ap | c_wni | c_out].copy()
            
            if df_final.empty:
                st.warning("Nenhuma violacao de SLA encontrada.")
            else:
                df_final['Tempo_SLA'] = df_final['Minutos_SLA'].apply(format_hms)
                
                colunas_finais = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration', 'Tempo_SLA']
                df_out = df_final[colunas_finais]

                st.success(f"Analise concluida!")
                st.dataframe(df_out, width='stretch', hide_index=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_out.to_excel(writer, index=False)
                st.download_button("Baixar Relatorio Final", output.getvalue(), "SLA_Final.xlsx")

    except Exception as e:
        st.error(f"Erro critico: {e}")
