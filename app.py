import streamlit as st
import os, io, pandas as pd
from motor_fiscal import extrair_dados_xml, gerar_excel_final

st.set_page_config(page_title="Sentinela Nascel 🧡", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton>button { background-color: #FF6F00; color: white; border-radius: 20px; font-weight: bold; width: 100%; height: 50px; border: none; }
    .stFileUploader { border: 1px dashed #FF6F00; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", use_container_width=True)
    st.markdown("---")
    st.subheader("🔄 Bases de Dados")
    st.file_uploader("Base ICMS", type=['xlsx'], key='base_icms_side')
    st.file_uploader("Base PIS/COFINS", type=['xlsx'], key='base_pc_side')
    st.markdown("---")
    st.subheader("📥 Modelos")
    m_buf = io.BytesIO()
    pd.DataFrame().to_excel(m_buf)
    st.download_button("Gabarito PIS/COFINS", m_buf.getvalue(), "modelo_pc.xlsx", use_container_width=True)

c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    if os.path.exists(".streamlit/Sentinela.png"):
        st.image(".streamlit/Sentinela.png", use_container_width=True)
    else:
        st.title("🚀 SENTINELA NASCEL 🧡")

st.markdown("---")
col_e, col_s = st.columns(2, gap="large")
with col_e:
    st.subheader("📥 ENTRADAS 🧡")
    xe = st.file_uploader("📂 XMLs Entrada", type='xml', accept_multiple_files=True, key="xe")
    ge = st.file_uploader("📊 Gerencial Entrada (CSV)", type=['csv'], key="ge")
    ae = st.file_uploader("🔍 Autenticidade Entrada", type=['xlsx'], key="ae")

with col_s:
    st.subheader("📤 SAÍDAS 🧡")
    xs = st.file_uploader("📂 XMLs Saída", type='xml', accept_multiple_files=True, key="xs")
    gs = st.file_uploader("📊 Gerencial Saída (CSV)", type=['csv'], key="gs")
    as_f = st.file_uploader("🔍 Autenticidade Saída", type=['xlsx'], key="as")

if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    with st.spinner("🧡 Auditoria em curso..."):
        try:
            df_xe = extrair_dados_xml(xe)
            df_xs = extrair_dados_xml(xs)
            relat = gerar_excel_final(df_xe, df_xs, ge, gs, ae, as_f)
            st.success("Concluído! 🧡")
            st.download_button("💾 BAIXAR RELATÓRIO", relat, "Auditoria_Sentinela.xlsx", use_container_width=True)
        except Exception as e:
            st.error(f"Erro: {e}")
