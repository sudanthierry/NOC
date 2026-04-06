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
        st.error(f"Erro ao salvar: {e}")
        return False

# --- 2. LOGICA DE CALCULO DE SLA ---
@st.cache_resource
def get_holidays():
    h = holidays.BR(state='PR', years=range(2024, 2030))
    for y in range(2024, 2030):
        h.append({f"{y}-09-08": "Nossa Senhora da Luz - Curitiba"})
    return h

def analyze_downtime(start, end, feriados):
    if pd.isnull(start) or pd.isnull(end) or start >= end: 
        return 0.0
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    for day in days:
        if day.weekday() >= 5 or day in feriados:
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

# --- 3. INTERFACE (SIDEBAR) ---
st.title("NOC SLA Analyser - Horario Comercial")

with st.sidebar:
    st.header("Gestao de Blacklist")
    with st.form("form_exclusao", clear_on_submit=True):
        st.subheader("Cadastrar Nova Excecao")
        nome_input = st.text_input("Nome do Equipamento:")
        lista_nocs = ["SME", "Leste", "Matriz", "Norte", "Oeste", "Sul"]
        noc_input = st.selectbox("Designar NOC:", lista_nocs)
        motivo_input = st.text_area("Justificativa:")
        if st.form_submit_button("Salvar na Nuvem"):
            if nome_input and motivo_input:
                if adicionar_a_blacklist(nome_input, motivo_input, noc_input):
                    st.success("Adicionado com sucesso!")
                    st.rerun()
            else:
                st.warning("Preencha Nome e Motivo.")
    st.divider()
    df_bl = carregar_blacklist_df()
    if not df_bl.empty:
        st.dataframe(df_bl, use_container_width=True, hide_index=True)

# --- 4. PROCESSAMENTO PRINCIPAL ---
file_main = st.file_uploader("Selecione o arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        # 1. Pula 8 linhas iniciais
        df = pd.read_excel(file_main, skiprows=8)
        
        # 2. Remove 5 linhas finais
        if len(df) > 5:
            df = df.iloc[:-5]
        
        # 3. Preenchimento (ffill)
        cols_preencher = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration']
        df[cols_preencher] = df[cols_preencher].ffill()

        # 4. Filtro Reason (apenas vazios)
        if 'Reason' in df.columns:
            df = df[df['Reason'].isna() | (df['Reason'].astype(str).str.strip() == "")].copy()

        # 5. Split do Nome (Remover parênteses)
        df['Device Name'] = df['Device Name'].astype(str).str.split('(').str[0].str.strip()

        # 6. Conversao de Datas
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'].astype(str).str.strip(), dayfirst=True, errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'].astype(str).str.strip(), dayfirst=True, errors='coerce')

        # 7. Filtro Blacklist
        if not df_bl.empty:
            ignorados = [str(x).strip().upper() for x in df_bl['Device Name'].tolist()]
            df = df[~df['Device Name'].str.upper().isin(ignorados)]

        with st.spinner('Processando SLA...'):
            feriados = get_holidays()
            df['Minutos_Comerciais'] = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End'], feriados), axis=1)
            
            df_filtrado = df[df['Minutos_Comerciais'] > 0].copy()
            
            if df_filtrado.empty:
                st.warning("Nenhum registro util encontrado.")
            else:
                df_filtrado['Tempo_SLA'] = df_filtrado['Minutos_Comerciais'].apply(format_hms)

                # Regras de Corte de SLA
                c_ap = df_filtrado['Device Name'].str.contains('AP', case=False, na=False)
                c_wni = df_filtrado['Device Name'].str.contains('WNI', case=False, na=False)
                
                df_violacoes = df_filtrado[
                    ((c_ap) & (df_filtrado['Minutos_Comerciais'] >= 240)) | 
                    ((c_wni) & (df_filtrado['Minutos_Comerciais'] >= 360)) | 
                    ((~c_ap) & (~c_wni) & (df_filtrado['Minutos_Comerciais'] >= 10))
                ].copy()

                # --- FILTRO DE COLUNAS SOLICITADO ---
                colunas_finais = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration', 'Tempo_SLA']
                df_final = df_violacoes[colunas_finais]

                st.success(f"Analise concluida: {len(df_final)} violacoes.")
                st.dataframe(df_final, use_container_width=True, hide_index=True)

                # Exportacao
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("Baixar Relatorio Final", output.getvalue(), "SLA_Final.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
