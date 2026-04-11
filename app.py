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

def remover_da_blacklist(nome_device):
    try:
        wks = conectar_google()
        data = wks.get_all_values()
        df_temp = pd.DataFrame(data[1:], columns=data[0])
        filtro = df_temp['Device Name'].str.strip().str.upper() == nome_device.strip().upper()
        idx = df_temp[filtro].index
        if not idx.empty:
            wks.delete_rows(int(idx[0]) + 2)
            st.cache_data.clear()
            return True
        else:
            st.error("Equipamento nao encontrado na lista.")
            return False
    except Exception as e:
        st.error(f"Erro ao excluir: {e}")
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
    with st.form("form_cadastro", clear_on_submit=True):
        st.subheader("Nova Excecao")
        nome_input = st.text_input("Nome do Equipamento:")
        opcoes_noc = ["", "SME", "Leste", "Matriz", "Norte", "Oeste", "Sul"]
        noc_input = st.selectbox("Designar NOC (Obrigatorio):", options=opcoes_noc, index=0)
        motivo_input = st.text_area("Justificativa:")
        if st.form_submit_button("Salvar na Nuvem"):
            if not nome_input or not motivo_input or noc_input == "":
                st.warning("Aviso: Preencha todos os campos.")
            else:
                if adicionar_a_blacklist(nome_input, motivo_input, noc_input):
                    st.success("Adicionado!")
                    st.rerun()

    with st.form("form_exclusao", clear_on_submit=True):
        st.subheader("Remover da Blacklist")
        nome_remover = st.text_input("Nome para excluir:")
        if st.form_submit_button("Excluir Item"):
            if nome_remover and remover_da_blacklist(nome_remover):
                st.success("Removido!")
                st.rerun()

    st.divider()
    df_bl = carregar_blacklist_df()
    if not df_bl.empty:
        st.dataframe(df_bl, width='stretch', hide_index=True)
        output_bl = io.BytesIO()
        with pd.ExcelWriter(output_bl, engine='xlsxwriter') as writer:
            for noc in sorted(df_bl['NOC'].unique()):
                df_bl[df_bl['NOC'] == noc].to_excel(writer, sheet_name=str(noc)[:31], index=False)
        st.download_button("Baixar Blacklist", output_bl.getvalue(), "Blacklist.xlsx")

# --- 4. PROCESSAMENTO PRINCIPAL ---
file_main = st.file_uploader("Upload DownTime.xlsx", type=['xlsx'])

if file_main:
    try:
        df_raw = pd.read_excel(file_main, skiprows=8)
        if len(df_raw) > 5: df_raw = df_raw.iloc[:-5]
        
        cols_fill = ['Device Name', 'Downtime Start', 'Downtime End', 'Duration']
        df_raw[cols_fill] = df_raw[cols_fill].ffill()

        tz_br = pytz.timezone('America/Sao_Paulo')
        agora_br = datetime.now(tz_br).replace(tzinfo=None)
        df_raw['Downtime End'] = df_raw['Downtime End'].astype(str).replace('Currently Down', agora_br.strftime('%Y-%m-%d %H:%M:%S'))

        if 'Reason' in df_raw.columns:
            df_raw = df_raw[df_raw['Reason'].isna() | (df_raw['Reason'].astype(str).str.strip() == "")].copy()

        df_raw['Device Name'] = df_raw['Device Name'].astype(str).str.split('(').str[0].str.strip()

        def smart_date_parser(val):
            val = str(val).strip()
            if not val or val == "nan": return pd.NaT
            try:
                if len(val) >= 4 and val[:4].isdigit():
                    return pd.to_datetime(val, yearfirst=True, dayfirst=False)
                else:
                    return pd.to_datetime(val, dayfirst=True, yearfirst=False)
            except:
                return pd.to_datetime(val, errors='coerce')

        df_raw['Downtime Start'] = df_raw['Downtime Start'].apply(smart_date_parser).dt.floor('s')
        df_raw['Downtime End'] = df_raw['Downtime End'].apply(smart_date_parser).dt.floor('s')

        # Blacklist
        if not df_bl.empty:
            lista_bl = [str(x).strip().upper() for x in df_bl['Device Name'].tolist()]
            df_desc_bl = df_raw[df_raw['Device Name'].str.upper().isin(lista_bl)].copy()
            df_desc_bl['Motivo_Descarte'] = "Equipamento em Blacklist"
            df_working = df_raw[~df_raw['Device Name'].str.upper().isin(lista_bl)].copy()
        else:
            df_desc_bl = pd.DataFrame()
            df_working = df_raw.copy()

        with st.spinner('Analisando...'):
            feriados = get_holidays()
            
            # DEFINICAO DAS REGRAS: AP e WNI sao Comercial. SWAP e outros sao 24/7.
            is_ap = df_working['Device Name'].str.contains('AP', case=False, na=False)
            is_wni = df_working['Device Name'].str.contains('WNI', case=False, na=False)
            is_comercial = (is_ap | is_wni) & (~df_working['Device Name'].str.contains('SWAP', case=False, na=False))
            
            # Calculo SLA
            df_working.loc[is_comercial, 'Minutos_SLA'] = df_working[is_comercial].apply(
                lambda r: analyze_downtime_comercial(r['Downtime Start'], r['Downtime End'], feriados), axis=1
            )
            df_working.loc[~is_comercial, 'Minutos_SLA'] = ((df_working['Downtime End'] - df_working['Downtime Start']).dt.total_seconds() / 60).fillna(0)

            # Regras de Corte
            # AP >= 4h comercial | WNI >= 6h comercial | Outros (incluindo SWAP) >= 10min total
            cond_ap = (is_ap) & (~df_working['Device Name'].str.contains('SWAP', case=False, na=False)) & (df_working['Minutos_SLA'] >= 240)
            cond_wni = (is_wni) & (~df_working['Device Name'].str.contains('SWAP', case=False, na=False)) & (df_working['Minutos_SLA'] >= 360)
            cond_outros = (~is_comercial) & (df_working['Minutos_SLA'] >= 10)
            
            df_final = df_working[cond_ap | cond_wni | cond_outros].copy()
            df_desc_sla = df_working[~(cond_ap | cond_wni | cond_outros)].copy()
            df_desc_sla['Motivo_Descarte'] = "Tempo de SLA insuficiente ou fora do horario comercial"

            df_total_desc = pd.concat([df_desc_bl, df_desc_sla], ignore_index=True)

            # --- EXIBICAO FINAL ---
            st.subheader("Violacoes de SLA (Relatorio Final)")
            if not df_final.empty:
                df_final['Tempo_SLA'] = df_final['Minutos_SLA'].apply(format_hms)
                tab1, tab2 = st.tabs(["AP e WNI (Comercial)", "Demais Equipamentos e SWAP (24/7)"])
                
                with tab1:
                    # Filtra AP/WNI que NAO sao SWAP
                    mask_comercial = df_final['Device Name'].str.contains('AP|WNI', case=False, na=False) & ~df_final['Device Name'].str.contains('SWAP', case=False, na=False)
                    st.dataframe(df_final[mask_comercial][['Device Name', 'Downtime Start', 'Downtime End', 'Duration', 'Tempo_SLA']], width='stretch', hide_index=True)
                
                with tab2:
                    # Filtra tudo que nao caiu na regra anterior
                    mask_247 = ~mask_comercial
                    st.dataframe(df_final[mask_247][['Device Name', 'Downtime Start', 'Downtime End', 'Duration', 'Tempo_SLA']], width='stretch', hide_index=True)

                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df_final[['Device Name', 'Downtime Start', 'Downtime End', 'Duration', 'Tempo_SLA']].to_excel(writer, sheet_name='Violacoes', index=False)
                    if not df_total_desc.empty:
                        df_total_desc[['Device Name', 'Downtime Start', 'Downtime End', 'Motivo_Descarte']].to_excel(writer, sheet_name='Desconsiderados', index=False)
                st.download_button("Baixar Relatorio Completo", out.getvalue(), "Relatorio_SLA.xlsx")
            else:
                st.info("Nenhuma violacao encontrada.")

            st.divider()
            st.subheader("Itens Desconsiderados")
            if not df_total_desc.empty:
                tab3, tab4 = st.tabs(["Desconsiderados AP/WNI", "Desconsiderados Demais e SWAP"])
                
                with tab3:
                    mask_desc_com = df_total_desc['Device Name'].str.contains('AP|WNI', case=False, na=False) & ~df_total_desc['Device Name'].str.contains('SWAP', case=False, na=False)
                    st.dataframe(df_total_desc[mask_desc_com][['Device Name', 'Downtime Start', 'Downtime End', 'Motivo_Descarte']], width='stretch', hide_index=True)
                
                with tab4:
                    mask_desc_247 = ~mask_desc_com
                    st.dataframe(df_total_desc[mask_desc_247][['Device Name', 'Downtime Start', 'Downtime End', 'Motivo_Descarte']], width='stretch', hide_index=True)

    except Exception as e:
        st.error(f"Erro: {e}")
