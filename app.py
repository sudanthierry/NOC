import streamlit as st
import pandas as pd
import holidays
import gspread
from google.oauth2.service_account import Credentials
import io
from datetime import datetime

# --- CONFIGURACOES DA PAGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide")

# --- 1. FUNCOES DE SUPORTE ---
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

# --- 2. INTERFACE E PROCESSAMENTO ---
st.title("NOC SLA Analyser - Business Hours")

file_main = st.file_uploader("Selecione o arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        # A. EXCLUSAO DAS 8 PRIMEIRAS LINHAS (skiprows)
        df = pd.read_excel(file_main, skiprows=8)
        
        # B. EXCLUSAO DAS 3 ULTIMAS LINHAS
        if len(df) > 3:
            df = df.iloc[:-3]
        
        # C. COMPLETAR ESPAÇOS EM BRANCO (DADOS DA LINHA ACIMA)
        # O ffill (forward fill) preenche os NaNs com o valor anterior
        cols_para_preencher = ['Device Name', 'Downtime Start', 'Downtime End']
        df[cols_para_preencher] = df[cols_para_preencher].ffill()

        # D. REALIZAR O SPLIT (Remover parênteses do Device Name)
        # Equivalente a: EXT.TEXTO(A2;1;PROCURAR("(";A2;1)-1)
        df['Device Name'] = df['Device Name'].astype(str).str.split('(').str[0].str.strip()

        # E. CONVERSAO DE DATAS (Tratamento para DD-MM-YY HH:MM:SS)
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'].astype(str).str.strip(), dayfirst=True, errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'].astype(str).str.strip(), dayfirst=True, errors='coerce')

        # F. CALCULO DE SLA
        br_feriados = get_holidays()
        
        with st.spinner('Calculando periodos comerciais...'):
            df['Minutos_Comerciais'] = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End'], br_feriados), axis=1)
            
            # Remove quedas que nao impactam o horario comercial (Tempo = 0)
            df_filtrado = df[df['Minutos_Comerciais'] > 0].copy()
            
            if df_filtrado.empty:
                st.warning("Nenhuma queda valida em horario comercial encontrada.")
            else:
                df_filtrado['Tempo_SLA'] = df_filtrado['Minutos_Comerciais'].apply(format_hms)

                # Regras de Corte de SLA
                c_ap = df_filtrado['Device Name'].str.contains('AP', case=False, na=False)
                c_wni = df_filtrado['Device Name'].str.contains('WNI', case=False, na=False)
                
                df_final = df_filtrado[
                    ((c_ap) & (df_filtrado['Minutos_Comerciais'] >= 240)) | 
                    ((c_wni) & (df_filtrado['Minutos_Comerciais'] >= 360)) | 
                    ((~c_ap) & (~c_wni) & (df_filtrado['Minutos_Comerciais'] >= 10))
                ].copy()

                st.success(f"Analise concluida: {len(df_final)} violacoes registradas.")
                st.dataframe(df_final.drop(columns=['Minutos_Comerciais']), use_container_width=True)

                # DOWNLOAD
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("Baixar Relatorio Final", output.getvalue(), "SLA_Consolidado.xlsx")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
