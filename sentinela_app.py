import streamlit as st
import os, io, pandas as pd
import requests
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela - Auditoria Fiscal", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Sentinela (Maximalista)
st.markdown("""
<style>
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton > button {
        background-color: #FF6F00 !important; color: white !important; border-radius: 25px !important;
        font-weight: bold !important; width: 300px !important; height: 50px !important; border: none !important;
    }
    .passo-container {
        background-color: #FFFFFF; padding: 15px; border-radius: 10px; border-left: 5px solid #FF6F00;
        margin-bottom: 20px; text-align: center; box-shadow: 0px 2px 5px rgba(0,0,0,0.1);
    }
    .passo-texto { color: #FF6F00; font-size: 1.2rem; font-weight: 800; }
    .stFileUploader section { background-color: #FFFFFF; border: 2px dashed #FF6F00 !important; }
</style>
""", unsafe_allow_html=True)

def listar_empresas_no_github():
    token = st.secrets.get("GITHUB_TOKEN")
    repo = st.secrets.get("GITHUB_REPO")
    if not token or not repo: return []
    url = f"https://api.github.com/repos/{repo}/contents/Bases_Tributárias"
    headers = {"Authorization": f"token {token}"}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            arquivos = response.json()
            return sorted(list(set([f['name'].split('-')[0] for f in arquivos if f['name'].endswith('.xlsx')])))
    except: pass
    return []

with st.sidebar:
    try: st.image(".streamlit/Sentinela.png", use_container_width=True)
    except: st.title("SENTINELA 🧡")
    st.markdown("---")
    def criar_gabarito():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            f_ncm = wb.add_format({'bg_color': '#444444', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar = wb.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            for s, c_l in [('ICMS', ["NCM", "CST (INTERNA)", "ALIQ (INTERNA)"]), 
                           ('PIS_COFINS', ["NCM", "CST Entrada", "CST Saída"]),
                           ('IPI', ["NCM", "CST_IPI", "ALQ_IPI"])]:
                pd.DataFrame(columns=c_l).to_excel(writer, sheet_name=s, index=False)
                for c, v in enumerate(c_l): writer.sheets[s].write(0, c, v, f_ncm if c == 0 else f_lar)
        return output.getvalue()
    st.download_button("📥 Gabarito para GitHub", criar_gabarito(), "gabarito_base_sentinela.xlsx", use_container_width=True)

st.markdown("<div class='passo-container'><span class='passo-texto'>👣 PASSO 1: Selecionar Empresa Cadastrada</span></div>", unsafe_allow_html=True)
cod_cliente = st.selectbox("Selecione a empresa cadastrada no GitHub:", [""] + listar_empresas_no_github(), label_visibility="collapsed")

if cod_cliente:
    st.markdown("<div class='passo-container'><span class='passo-texto'>👣 PASSO 2: Carregar XMLs e Autenticidade</span></div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📤 SAÍDAS (Auditoria)")
        xs = st.file_uploader("XMLs Saída", type='xml', accept_multiple_files=True, key="xs_v74")
        gs = st.file_uploader("Gerencial Saída (CSV)", type=['csv'], key="gs_v74")
        as_f = st.file_uploader("Autenticidade Saída (Excel)", type=['xlsx'], key="as_v74")
    with c2:
        st.subheader("📥 ENTRADAS (Cruzamento ST)")
        xe = st.file_uploader("XMLs Entrada", type='xml', accept_multiple_files=True, key="xe_v74")
        ge = st.file_uploader("Gerencial Entrada (CSV)", type=['csv'], key="ge_v74")
        ae = st.file_uploader("Autenticidade Entrada (Excel)", type=['xlsx'], key="ae_v74")

    if st.button("🚀 EXECUTAR DIAGNÓSTICO MAXIMALISTA"):
        if not xs: st.warning("Por favor, carregue os XMLs de Saída para iniciar.")
        else:
            with st.spinner("🧡 Sentinela processando motor maximalista..."):
                try:
                    df_xe = extrair_dados_xml(xe); df_xs = extrair_dados_xml(xs)
                    relat = gerar_excel_final(df_xe, df_xs, ae, as_f, ge, gs, cod_cliente)
                    st.success("Diagnóstico Gerado com Sucesso! 🧡")
                    st.download_button("💾 BAIXAR RELATÓRIO COMPLETO", relat, f"Sentinela_Diagnostico_{cod_cliente}.xlsx", use_container_width=True)
                except Exception as e: st.error(f"Erro no Processamento: {e}")
