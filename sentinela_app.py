import streamlit as st
import os, io, pandas as pd
import requests
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela - Auditoria Fiscal", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Sentinela (Laranja Restaurado e Blindado)
st.markdown("""
<style>
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    #MainMenu {visibility: hidden !important;}
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:has(img) {
        display: flex !important; justify-content: center !important; width: 100% !important;
    }
    .stButton { display: flex !important; justify-content: center !important; width: 100% !important; margin-top: 20px !important; }
    .stButton > button {
        background-color: #FF6F00 !important; color: white !important; border-radius: 25px !important;
        font-weight: bold !important; width: 300px !important; height: 50px !important; border: none !important;
    }
    .passo-container {
        background-color: #FFFFFF; padding: 10px 15px; border-radius: 10px; border-left: 5px solid #FF6F00;
        margin: 10px auto 15px auto; max-width: 600px; text-align: center;
    }
    .passinho { color: #808080 !important; font-size: 1.2rem; margin-right: 10px; }
    .passo-texto { color: #FF6F00; font-size: 1.1rem; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

def listar_empresas_no_repositorio():
    token = st.secrets.get("GITHUB_TOKEN")
    repo = st.secrets.get("GITHUB_REPO")
    if not token or not repo: return []
    url = f"https://api.github.com/repos/{repo}/contents/Bases_Tributárias"
    headers = {"Authorization": f"token {token}"}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 200:
            return sorted(list(set([f['name'].split('-')[0] for f in res.json() if f['name'].endswith('.xlsx')])))
    except: pass
    return []

with st.sidebar:
    if os.path.exists(".streamlit/Sentinela.png"):
        st.image(".streamlit/Sentinela.png", use_container_width=True)
    st.markdown("---")
    def criar_gabarito():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            f_header = wb.add_format({'bg_color': '#7F7F7F', 'font_color': '#CCECFF', 'bold': True, 'border': 1})
            # ICMS
            pd.DataFrame(columns=["NCM", "CST (INTERNA)", "ALIQ (INTERNA)", "CST (ESTADUAL)"]).to_excel(writer, sheet_name='ICMS', index=False)
            for c, v in enumerate(["NCM", "CST (INTERNA)", "ALIQ (INTERNA)", "CST (ESTADUAL)"]): writer.sheets['ICMS'].write(0, c, v, f_header)
            # PIS_COFINS
            pd.DataFrame(columns=["NCM", "CST Entrada", "CST Saída"]).to_excel(writer, sheet_name='PIS_COFINS', index=False)
            for c, v in enumerate(["NCM", "CST Entrada", "CST Saída"]): writer.sheets['PIS_COFINS'].write(0, c, v, f_header)
            # IPI
            pd.DataFrame(columns=["NCM_TIPI", "EX", "DESCRIÇÃO", "ALÍQUOTA (%)"]).to_excel(writer, sheet_name='IPI', index=False)
            for c, v in enumerate(["NCM_TIPI", "EX", "DESCRIÇÃO", "ALÍQUOTA (%)"]): writer.sheets['IPI'].write(0, c, v, f_header)
        return output.getvalue()
    st.download_button("📥 Baixar Gabarito", criar_gabarito(), "gabarito_sentinela.xlsx", use_container_width=True)

st.markdown("<div class='passo-container'><span class='passo-texto'>👣 PASSO 1: Selecione a empresa cadastrada</span></div>", unsafe_allow_html=True)
cod_cliente = st.selectbox("Empresa:", [""] + listar_empresas_no_repositorio(), label_visibility="collapsed")

if cod_cliente:
    st.markdown("<div class='passo-container'><span class='passinho'>👣</span><span class='passo-texto'>PASSO 2: Carregar Documentos</span></div>", unsafe_allow_html=True)
    c_e, c_s = st.columns(2, gap="large")
    with c_e:
        st.subheader("📥 ENTRADAS")
        xe = st.file_uploader("XMLs Entrada", type='xml', accept_multiple_files=True, key="xe_v103")
        ge = st.file_uploader("Gerencial Entrada", type=['csv'], key="ge_v103")
        ae = st.file_uploader("Autenticidade Entrada", type=['xlsx', 'csv'], key="ae_v103")
    with c_s:
        st.subheader("📤 SAÍDAS")
        xs = st.file_uploader("XMLs Saída", type='xml', accept_multiple_files=True, key="xs_v103")
        gs = st.file_uploader("Gerencial Saída", type=['csv'], key="gs_v103")
        as_f = st.file_uploader("Autenticidade Saída", type=['xlsx', 'csv'], key="as_v103")

    if st.button("🚀 GERAR RELATÓRIO"):
        with st.spinner("🧡 Sentinela está processando..."):
            try:
                df_xe = extrair_dados_xml(xe); df_xs = extrair_dados_xml(xs)
                relat = gerar_excel_final(df_xe, df_xs, ae, as_f, ge, gs, cod_cliente)
                st.success("Auditoria Concluída! 🧡")
                st.download_button("💾 BAIXAR AGORA", relat, f"Sentinela_{cod_cliente}.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Erro Crítico: {e}")
