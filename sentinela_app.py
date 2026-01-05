import streamlit as st
import os, io, pandas as pd
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela Nascel", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Nascel (Ajustado para reduzir área branca da logo)
st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton>button { background-color: #FF6F00; color: white; border-radius: 20px; font-weight: bold; width: 100%; height: 50px; border: none; }
    .stFileUploader { border: 1px dashed #FF6F00; border-radius: 10px; }
    
    /* Reduz o espaçamento do bloco da imagem e o topo da página */
    .block-container { padding-top: 1rem !important; padding-bottom: 0rem !important; }
    [data-testid="stVerticalBlock"] > div:first-child { margin-top: -50px; }
    [data-testid="stImage"] { margin-bottom: -30px; margin-top: -20px; }
</style>
""", unsafe_allow_html=True)

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", use_container_width=True)
    
    st.markdown("---")
    st.subheader("🔄 Bases de Referência")
    u_icms = st.file_uploader("Subir Base ICMS (XLSX)", type=['xlsx'], key='base_icms_v3')
    u_ipi = st.file_uploader("Subir Base IPI (XLSX)", type=['xlsx'], key='base_ipi_v3')
    u_pc = st.file_uploader("Subir Base PIS/COFINS (XLSX)", type=['xlsx'], key='base_pc_v3')
    
    st.markdown("---")
    st.subheader("📥 Gabaritos")
    
    def criar_gabarito(colunas):
        buf = io.BytesIO()
        pd.DataFrame(columns=colunas).to_excel(buf, index=False)
        return buf.getvalue()

    st.download_button("📥 Gabarito PIS/COFINS", criar_gabarito(["NCM", "ALIQUOTA_PIS", "ALIQUOTA_COFINS", "CST"]), "gabarito_pis_cofins.xlsx", use_container_width=True)
    st.download_button("📥 Gabarito ICMS", criar_gabarito(["NCM", "ALIQUOTA_ICMS", "CST_ICMS", "REDUCAO_BC"]), "gabarito_icms.xlsx", use_container_width=True)
    st.download_button("📥 Gabarito IPI", criar_gabarito(["NCM", "ALIQUOTA_IPI", "CST_IPI", "ENQUADRAMENTO"]), "gabarito_ipi.xlsx", use_container_width=True)

# --- 4. TELA PRINCIPAL ---
# Mantive a proporção que você gostou, mas o CSS acima remove o "vazio"
c1, c2, c3 = st.columns([1.5, 1, 1.5]) 
with c2:
    if os.path.exists(".streamlit/Sentinela.png"):
        st.image(".streamlit/Sentinela.png", use_container_width=True)
    else:
        st.title("🚀 SENTINELA NASCEL")

st.markdown("---")

col_e, col_s = st.columns(2, gap="large")

with col_e:
    st.subheader("📥 FLUXO ENTRADAS")
    xe = st.file_uploader("📂 XMLs de Entrada", type='xml', accept_multiple_files=True, key="xe_v3")
    ge = st.file_uploader("📊 Gerencial Entrada (CSV)", type=['csv'], key="ge_v3")
    ae = st.file_uploader("🔍 Autenticidade Entrada (XLSX)", type=['xlsx'], key="ae_v3")

with col_s:
    st.subheader("📤 FLUXO SAÍDAS")
    xs = st.file_uploader("📂 XMLs de Saída", type='xml', accept_multiple_files=True, key="xs_v3")
    gs = st.file_uploader("📊 Gerencial Saída (CSV)", type=['csv'], key="gs_v3")
    as_f = st.file_uploader("🔍 Autenticidade Saída (XLSX)", type=['xlsx'], key="as_v3")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    if not xe and not xs:
        st.warning("Por favor, suba ao menos um arquivo XML.")
    else:
        with st.spinner("🧡 O Sentinela está cruzando os dados..."):
            try:
                df_xe = extrair_dados_xml(xe)
                df_xs = extrair_dados_xml(xs)
                relat = gerar_excel_final(df_xe, df_xs, u_icms, u_pc, ae, as_f, ge, gs, u_ipi)
                st.success("Auditoria concluída com sucesso! 🧡")
                st.download_button("💾 BAIXAR RELATÓRIO FINAL", relat, "Auditoria_Sentinela.xlsx", use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")
