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

st.set_page_config(page_title="NOC SLA Analyser", layout="wide")

# --- FUNÇÕES DE SUPORTE ---
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

# --- INTERFACE ---
st.title("NOC SLA Analyser")

file_main = st.file_uploader("Selecione o arquivo DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        # 1. Limpeza de linhas
        df = pd.read_excel(file_main, skiprows=8)
        if len(df) > 5:
            df = df.iloc[:-5]
        
        # 2. Preenchimento de lacunas
        cols_fill = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration']
        df[cols_fill] = df[cols_fill].ffill()

        # 3. Tratamento "Currently Down" com Horário de Brasília
        timezone_br = pytz.timezone('America/Sao_Paulo')
        agora_br = datetime.now(timezone_br).replace(tzinfo=None)
        df['Downtime End'] = df['Downtime End'].astype(str).replace('Currently Down', agora_br.strftime('%Y-%m-%d %H:%M:%S'))

        # 4. Filtro Reason
        if 'Reason' in df.columns:
            df = df[df['Reason'].isna() | (df['Reason'].astype(str).str.strip() == "")].copy()

        # 5. Split e Conversão de Datas
        df['Device Name'] = df['Device Name'].astype(str).str.split('(').str[0].str.strip()
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'].astype(str).str.strip(), dayfirst=True, errors='coerce', format='mixed')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'].astype(str).str.strip(), dayfirst=True, errors='coerce', format='mixed')

        with st.spinner('Calculando SLA...'):
            feriados = get_holidays()
            is_ap = df['Device Name'].str.contains('AP', case=False, na=False)
            is_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
            
            # AP/WNI: Horário Comercial | Outros: 24/7
            df.loc[is_ap | is_wni, 'Minutos_SLA'] = df[is_ap | is_wni].apply(
                lambda r: analyze_downtime_comercial(r['Downtime Start'], r['Downtime End'], feriados), axis=1
            )
            df.loc[~(is_ap | is_wni), 'Minutos_SLA'] = ((df['Downtime End'] - df['Downtime Start']).dt.total_seconds() / 60).fillna(0)

            # Cortes de SLA
            cond_ap = (
