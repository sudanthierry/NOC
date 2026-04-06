import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
from datetime import datetime

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide", page_icon="???")

# --- 1. CONEXÃO GOOGLE SHEETS (EXCEÇÕES) ---
def conectar_google():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    # Tenta ler das Secrets do Streamlit (GitHub/Cloud) ou arquivo local
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
    client = gspread.authorize(creds)
    return client.open("noc_config").worksheet("blacklist")

def carregar_blacklist():
    try:
        wks = conectar_google()
        lista = wks.col_values(1) # Pega a primeira coluna (Device Name)
        return [str(x).strip() for x in lista[1:] if x] # Remove o cabeçalho
    except:
        return []

def adicionar_excecao(nome_device):
    try:
        wks = conectar_google()
        wks.append_row([nome_device])
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

# --- 2. LÓGICA DE NEGÓCIO (SLA) ---
@st.cache_resource
def get_holidays():
    h = holidays.BR(state='PR', years=range(2024, 2027))
    for y in range(2024, 2027):
        h.append({f"{y}-09-08": "Nossa Senhora da Luz"})
    return h

br_holidays = get_holidays()

def analyze_downtime(start, end):
    if pd.isnull(start) or pd.isnull(end) or start >= end: return 0.0, "Não"
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    was_non_working = "Não"
    for day in days:
        if day.weekday() >= 5 or day in br_holidays:
            was_non_working = "Sim"
            continue 
        work_start = day.replace(hour=8, minute=0, second=0)
        work_end = day.replace(hour=18, minute=0, second=0)
        actual_start = max(start, work_start)
        actual_end = min(end, work_end)
        if actual_start < actual_end:
            total_minutes += (actual_end - actual_start).total_seconds() / 60
    return total_minutes, was_non_working

def format_hms(m):
    ts = int(m * 60)
    return f"{ts//3600:02d}:{(ts%3600)//60:02d}:{ts%60:02d}"

# --- 3. INTERFACE ---
st.title("??? NOC SLA Analyser + Cloud Exceptions")

# Aba lateral para gerenciar Blacklist
with st.sidebar:
    st.header("?? Gerenciar Exceções (Google Sheets)")
    nova_exc = st.text_input("Adicionar Device à Blacklist:")
    if st.button("Salvar na Nuvem"):
        if nova_exc:
            if adicionar_excecao(nova_exc):
                st.success(f"{nova_exc} adicionado!")
                st.rerun()
    
    st.divider()
    blacklist_atual = carregar_blacklist()
    st.write(f"**Dispositivos Ignorados:** {len(blacklist_atual)}")
    with st.expander("Ver Lista"):
        st.write(blacklist_atual)

# Upload do arquivo principal
file_main = st.file_uploader("Selecione o arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    df = pd.read_excel(file_main, skiprows=8)
    
    # Limpeza e Ffill
    cols_to_fix = ['Device Name', 'Downtime Start', 'Downtime End']
    df[cols_to_fix] = df[cols_to_fix].ffill()
    
    # Aplicar Blacklist vinda do Google Sheets
    if blacklist_atual:
        df = df[~df['Device Name'].astype(str).str.strip().isin(blacklist_atual)]
        st.info(f"Filtro aplicado: {len(blacklist_atual)} dispositivos da nuvem foram ignorados.")

    # Processamento de strings e Reason
    strings_rem = ["*****", "Personally Identifiable Data", "NOTE"]
    for s in strings_rem:
        df = df[~df.apply(lambda row: row.astype(str).str.contains(s, na=False, regex=False).any(), axis=1)]
    
    if 'Reason' in df.columns:
        df = df[(df['Reason'].isna()) | (df['Reason'].astype(str).isin(['', 'nan', 'None']))].copy()

    df['Downtime Start'] = pd.to_datetime(df['Downtime Start'], errors='coerce')
    df['Downtime End'] = pd.to_datetime(df['Downtime End'], errors='coerce')

    # Cálculos
    with st.spinner('Analisando períodos comerciais...'):
        results = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End']), axis=1)
        df[['_temp_min', 'FDS_Feriado']] = pd.DataFrame(results.tolist(), index=df.index)
        df['Tempo_Comercial'] = df['_temp_min'].apply(format_hms)

        # Regras de SLA
        c_ap = df['Device Name'].str.contains('AP', case=False, na=False)
        c_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
        df_f = df[((c_ap) & (df['_temp_min'] >= 240)) | 
                  ((c_wni) & (df['_temp_min'] >= 360)) | 
                  ((~c_ap) & (~c_wni) & (df['_temp_min'] >= 10))].copy()

    st.success(f"Análise concluída: {len(df_f)} violações encontradas.")
    st.dataframe(df_f.drop(columns=['_temp_min']), use_container_width=True)

    # Download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, index=False)
    st.download_button("?? Baixar Relatório", output.getvalue(), "SLA_NOC.xlsx")
