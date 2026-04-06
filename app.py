import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
from datetime import datetime

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide", page_icon="???")

# --- 1. FUNÇÕES DE CONEXÃO E BANCO DE DADOS (GOOGLE SHEETS) ---
def conectar_google():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    else:
        try:
            creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        except:
            st.error("Erro: Credenciais do Google não encontradas.")
            st.stop()
    client = gspread.authorize(creds)
    return client.open("noc_config").worksheet("blacklist")

def carregar_blacklist_df():
    try:
        wks = conectar_google()
        data = wks.get_all_records()
        return pd.DataFrame(data)
    except:
        return pd.DataFrame(columns=["Device Name", "Motivo"])

def adicionar_a_blacklist(nome_device, motivo_texto):
    try:
        wks = conectar_google()
        # Salva o par: Nome e Motivo
        wks.append_row([nome_device.strip(), motivo_texto.strip()])
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
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
        return 0.0, "Não"
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
    return f"{ts // 3600:02d}:{(ts % 3600) // 60:02d}:{ts % 60:02d}"

# --- 3. INTERFACE STREAMLIT ---
st.title("?? NOC SLA Analyser + Cloud Exceptions")

# --- SIDEBAR: FORMULÁRIO COM CAMPO DE MOTIVO ---
with st.sidebar:
    st.header("??? Gestão de Blacklist")
    
    # O formulário agora contém explicitamente os dois campos
    with st.form("form_exclusao", clear_on_submit=True):
        st.subheader("Cadastrar Nova Exceção")
        
        # CAMPO 1: NOME
        nome_input = st.text_input("Nome do Equipamento (Exato):")
        
        # CAMPO 2: MOTIVO (O que estava faltando)
        motivo_input = st.text_area("Motivo/Justificativa da inclusão:")
        
        btn_enviar = st.form_submit_button("Adicionar à Blacklist")
        
        if btn_enviar:
            if nome_input and motivo_input:
                if adicionar_a_blacklist(nome_input, motivo_input):
                    st.success(f"Dispositivo {nome_input} bloqueado!")
                    st.rerun()
            else:
                st.warning("Atenção: Nome e Motivo são obrigatórios.")

    st.divider()
    st.subheader("?? Lista de Dispositivos Ignorados")
    df_bl = carregar_blacklist_df()
    if not df_bl.empty:
        st.dataframe(df_bl, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma exceção cadastrada.")

# --- ÁREA PRINCIPAL: PROCESSAMENTO ---
file_main = st.file_uploader("Selecione o arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        df = pd.read_excel(file_main, skiprows=8)
        cols_to_fix = ['Device Name', 'Downtime Start', 'Downtime End']
        df[cols_to_fix] = df[cols_to_fix].ffill()
        
        # Filtragem pela Blacklist do Google Sheets
        if not df_bl.empty:
            ignorados = df_bl['Device Name'].astype(str).str.strip().tolist()
            df = df[~df['Device Name'].astype(str).str.strip().isin(ignorados)]
            st.info(f"Filtro Ativo: {len(ignorados)} dispositivos da nuvem ignorados.")

        # Limpeza de ruído e Reason
        strings_rem = ["*****", "Personally Identifiable Data", "NOTE"]
        for s in strings_rem:
            df = df[~df.apply(lambda row: row.astype(str).str.contains(s, na=False, regex=False).any(), axis=1)]
        
        if 'Reason' in df.columns:
            df = df[(df['Reason'].isna()) | (df['Reason'].astype(str).isin(['', 'nan', 'None']))].copy()

        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'], errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'], errors='coerce')

        with st.spinner('Analisando períodos...'):
            res = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End']), axis=1)
            df[['_temp_min', 'FDS_Feriado']] = pd.DataFrame(res.tolist(), index=df.index)
            df['Tempo_SLA'] = df['_temp_min'].apply(format_hms)

            # Regras SLA
            c_ap = df['Device Name'].str.contains('AP', case=False, na=False)
            c_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
            df_final = df[((c_ap) & (df['_temp_min'] >= 240)) | 
                          ((c_wni) & (df['_temp_min'] >= 360)) | 
                          ((~c_ap) & (~c_wni) & (df['_temp_min'] >= 10))].copy()

        st.success(f"Análise concluída: {len(df_final)} violações encontradas.")
        st.dataframe(df_final.drop(columns=['_temp_min']), use_container_width=True)

        # Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Violações_SLA')
        st.download_button("?? Baixar Relatório Final", output.getvalue(), "Relatorio_SLA.xlsx")

    except Exception as e:
        st.error(f"Erro: {e}")
