import streamlit as st
import os, io, pandas as pd
# Aqui conectamos com o seu arquivo de lógica
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página e Título na aba do navegador
st.set_page_config(page_title="Sentinela Nascel 🧡", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS para deixar tudo com as cores da Nascel
st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton>button { background-color: #FF6F00; color: white; border-radius: 20px; font-weight: bold; width: 100%; height: 50px; border: none; }
    .stFileUploader { border: 1px dashed #FF6F00; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# --- 3. SIDEBAR (BARRA LATERAL) ---
with st.sidebar:
    # Tenta carregar a logo da Nascel se ela estiver na pasta .streamlit
    logo_path = ".streamlit/nascel sem fundo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    
    st.markdown("---")
    st.subheader("🔄 Bases de Referência")
    st.info("Suba aqui suas planilhas de consulta (ICMS/PIS/COFINS) para o Sentinela cruzar com os XMLs.")
    
    # Campos para você subir as bases do seu computador NA HORA DO USO
    u_icms = st.file_uploader("Subir Base ICMS (XLSX)", type=['xlsx'], key='s_icms')
    u_pc = st.file_uploader("Subir Base PIS/COFINS (XLSX)", type=['xlsx'], key='s_pc')
    
    st.markdown("---")
    st.subheader("📥 Gabaritos")
    m_buf = io.BytesIO()
    pd.DataFrame(columns=["NCM", "ALIQUOTA", "CST"]).to_excel(m_buf, index=False)
    st.download_button("Baixar Modelo de Base", m_buf.getvalue(), "modelo_gabarito.xlsx", use_container_width=True)

# --- 4. TELA PRINCIPAL ---
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    soldado = ".streamlit/Sentinela.png"
    if os.path.exists(soldado):
        st.image(soldado, use_container_width=True)
    else:
        st.title("🚀 SENTINELA NASCEL 🧡")

st.markdown("---")

# Divisão em duas colunas para Entradas e Saídas
col_e, col_s = st.columns(2, gap="large")

with col_e:
    st.subheader("📥 FLUXO ENTRADAS 🧡")
    xe = st.file_uploader("📂 XMLs de Entrada", type='xml', accept_multiple_files=True, key="xe_main")
    ae = st.file_uploader("🔍 Autenticidade Entrada (XLSX)", type=['xlsx'], key="ae_main")

with col_s:
    st.subheader("📤 FLUXO SAÍDAS 🧡")
    xs = st.file_uploader("📂 XMLs de Saída", type='xml', accept_multiple_files=True, key="xs_main")
    as_f = st.file_uploader("🔍 Autenticidade Saída (XLSX)", type=['xlsx'], key="as_main")

st.markdown("<br>", unsafe_allow_html=True)

# Botão que dispara o Motor
if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    if not xe and not xs:
        st.warning("Por favor, suba ao menos um arquivo XML para analisar.")
    else:
        with st.spinner("🧡 O Sentinela está auditando seus dados..."):
            try:
                # Chama as funções que estão no sentinela_core.py
                df_xe = extrair_dados_xml(xe)
                df_xs = extrair_dados_xml(xs)
                
                # Passa as bases da sidebar (u_icms, u_pc) para o relatório final
                relat = gerar_excel_final(df_xe, df_xs, u_icms, u_pc, ae, as_f)
                
                st.success("Auditoria concluída com sucesso! 🧡")
                st.download_button("💾 BAIXAR RELATÓRIO FINAL", relat, "Auditoria_Sentinela.xlsx", use_container_width=True)
            except Exception as e:
                st.error(f"Erro no processamento: {e}")
