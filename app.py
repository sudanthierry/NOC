import streamlit as st
import pandas as pd
import holidays
import io
from datetime import datetime

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="NOC SLA Analyser", layout="wide", page_icon="??")

# Custom CSS para melhorar o visual
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; shadow: 2px 2px 5px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

st.title("?? Analisador de SLA - NOC In Loco")
st.subheader("Cálculo de Downtime considerando Horário Comercial e Feriados (Curitiba)")

# 1?? Configuração de Feriados (Curitiba)
@st.cache_resource
def get_holidays():
    # Abrangendo um range maior para evitar erros de data
    br_holidays = holidays.BR(state='PR', years=range(2023, 2030))
    for year in range(2023, 2030):
        # Nossa Senhora da Luz - Curitiba
        br_holidays.append({f"{year}-09-08": "Nossa Senhora da Luz - Curitiba"})
    return br_holidays

br_holidays = get_holidays()

def analyze_downtime(start, end):
    if pd.isnull(start) or pd.isnull(end) or start >= end: 
        return 0.0, "Não"
    
    # Criar range de dias entre o início e o fim
    days = pd.date_range(start.date(), end.date(), freq='D')
    total_minutes = 0.0
    was_non_working = "Não"
    
    for day in days:
        # Verifica se é final de semana ou feriado
        if day.weekday() >= 5 or day in br_holidays:
            was_non_working = "Sim"
            continue 
        
        # Define horário comercial: 08:00 às 18:00
        work_start = day.replace(hour=8, minute=0, second=0)
        work_end = day.replace(hour=18, minute=0, second=0)
        
        # Interseção entre o downtime e o horário comercial
        actual_start = max(start, work_start)
        actual_end = min(end, work_end)
        
        if actual_start < actual_end:
            total_minutes += (actual_end - actual_start).total_seconds() / 60
            
    return total_minutes, was_non_working

def format_to_hms(minutes):
    if minutes <= 0: return "00:00:00"
    total_seconds = int(minutes * 60)
    hours = total_seconds // 3600
    mins = (total_seconds % 3600) // 60
    secs = total_seconds % 60
    return f"{hours:02g}:{mins:02d}:{secs:02d}"

# --- INTERFACE DE UPLOAD ---
with st.sidebar:
    st.header("Configurações de Importação")
    file_main = st.file_uploader("Arquivo DownTime.xlsx", type=['xlsx'])
    file_excl = st.file_uploader("Lista de Exclusões (Opcional)", type=['xlsx'])
    st.info("O arquivo principal deve ter os dados a partir da linha 9 (skiprows=8).")

if file_main:
    try:
        # Carregamento dos dados
        df = pd.read_excel(file_main, skiprows=8)
        
        # Limpeza inicial de strings indesejadas
        strings_para_remover = ["*****", "Personally Identifiable Data", "NOTE"]
        for s in strings_para_remover:
            df = df[~df.apply(lambda row: row.astype(str).str.contains(s, na=False, regex=False).any(), axis=1)]

        # Ajuste de colunas e preenchimento de vazios (ffill)
        cols_to_fix = ['Device Name', 'Downtime Start', 'Downtime End']
        df[cols_to_fix] = df[cols_to_fix].ffill()
        
        # Filtro de Blacklist
        if file_excl:
            df_excl = pd.read_excel(file_excl)
            if 'Device Name' in df_excl.columns:
                lista_negra = df_excl['Device Name'].astype(str).str.strip().unique().tolist()
                df = df[~df['Device Name'].astype(str).str.strip().isin(lista_negra)]
                st.sidebar.success(f"{len(lista_negra)} dispositivos excluídos.")

        # Filtro de 'Reason' (Motivo)
        if 'Reason' in df.columns:
            df = df[(df['Reason'].isna()) | (df['Reason'].astype(str).isin(['', 'nan', 'None']))].copy()

        # Conversão de Datas
        df['Downtime Start'] = pd.to_datetime(df['Downtime Start'], errors='coerce')
        df['Downtime End'] = pd.to_datetime(df['Downtime End'], errors='coerce')

        with st.spinner('Calculando SLA e filtrando regras comerciais...'):
            # Aplicação da análise
            res = df.apply(lambda r: analyze_downtime(r['Downtime Start'], r['Downtime End']), axis=1)
            df[['_temp_min', 'Ocorreu_FDS_ou_Feriado']] = pd.DataFrame(res.tolist(), index=df.index)
            df['Downtime Duration (Comercial)'] = df['_temp_min'].apply(format_to_hms)

            # Filtros de Regra de SLA (Minutos)
            # AP >= 4h (240min) | WNI >= 6h (360min) | Outros >= 10min
            cond_ap = df['Device Name'].str.contains('AP', case=False, na=False)
            cond_wni = df['Device Name'].str.contains('WNI', case=False, na=False)
            
            df_filtered = df[
                ((cond_ap) & (df['_temp_min'] >= 240)) | 
                ((cond_wni) & (df['_temp_min'] >= 360)) | 
                ((~cond_ap) & (~cond_wni) & (df['_temp_min'] >= 10))
            ].copy()

        # Exibição de Resultados
        st.success(f"Análise concluída: {len(df_filtered)} registros violam o SLA.")
        
        # Métricas Rápidas
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Analisado", len(df))
        m2.metric("Violações de SLA", len(df_filtered))
        m3.metric("Feriados/FDS Detectados", len(df[df['Ocorreu_FDS_ou_Feriado'] == "Sim"]))

        st.dataframe(df_filtered, use_container_width=True)

        # Botão de Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name='SLA_Violations')
        
        st.download_button(
            label="?? Baixar Relatório de Violações",
            data=output.getvalue(),
            file_name=f"Relatorio_SLA_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar arquivo: {e}")
else:
    st.warning("Aguardando upload do arquivo DownTime.xlsx para iniciar.")
