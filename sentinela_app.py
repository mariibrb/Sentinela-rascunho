import streamlit as st
import os, io, pandas as pd
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela Nascel", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Nascel (Compacto e Laranja)
st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton>button { background-color: #FF6F00; color: white; border-radius: 20px; font-weight: bold; width: 100%; height: 50px; border: none; }
    .stFileUploader { border: 1px dashed #FF6F00; border-radius: 10px; }
    .block-container { padding-top: 0.5rem !important; padding-bottom: 0rem !important; }
    [data-testid="stVerticalBlock"] > div:first-child { margin-top: -20px; }
    [data-testid="stImage"] { text-align: center; margin-bottom: -20px; }
</style>
""", unsafe_allow_html=True)

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", use_container_width=True)
    
    st.markdown("---")
    st.subheader("🏢 Identificação")
    cod_cliente = st.text_input("Código do Cliente (ex: 394)", key="cod_cli")

    st.subheader("🔄 Bases de Referência")
    u_base_unica = st.file_uploader("Subir Base de Auditoria (XLSX)", type=['xlsx'], key='base_unica_v5')
    
    st.markdown("---")
    st.subheader("📥 Gabaritos")
    
    def criar_gabarito_nascel():
        output = io.BytesIO()
        colunas = [
            "NCM", "BASE REDUZIDA", "CST", "ALÍQUOTA ICMS", ".", 
            "BASE REDUZIDA2", "CST3", ",", "ALÍQUOTA ICMS5",    
            "NCM_TIPI", "EX", "DESCRIÇÃO", "ALÍQUOTA (%)",     
            "NCM_PC", "Entrada", "Saída", "CFOP-CST", "Status" 
        ]
        df = pd.DataFrame(columns=colunas)
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Base_Auditoria', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Base_Auditoria']
            f_lar_e = workbook.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar_c = workbook.add_format({'bg_color': '#FFB74D', 'bold': True, 'border': 1})
            f_cin_e = workbook.add_format({'bg_color': '#757575', 'font_color': 'white', 'bold': True, 'border': 1})
            f_cin_c = workbook.add_format({'bg_color': '#E0E0E0', 'bold': True, 'border': 1})

            for c, v in enumerate(colunas):
                if c <= 4: worksheet.write(0, c, v, f_lar_e)
                elif c <= 8: worksheet.write(0, c, v, f_lar_c)
                elif c <= 12: worksheet.write(0, c, v, f_cin_e)
                else: worksheet.write(0, c, v, f_cin_c)
        return output.getvalue()

    st.download_button("📥 Gabarito Base Nascel", criar_gabarito_nascel(), "base_auditoria_nascel.xlsx", use_container_width=True)

# --- 4. TELA PRINCIPAL ---
c1, c2, c3 = st.columns([1.2, 1, 1.2]) 
with c2:
    if os.path.exists(".streamlit/Sentinela.png"):
        st.image(".streamlit/Sentinela.png", use_container_width=True)
    else:
        st.title("🚀 SENTINELA")

st.markdown("---")
col_e, col_s = st.columns(2, gap="large")

with col_e:
    st.subheader("📥 FLUXO ENTRADAS")
    xe = st.file_uploader("📂 XMLs de Entrada", type='xml', accept_multiple_files=True, key="xe_v5")
    ge = st.file_uploader("📊 Gerencial Entrada (CSV)", type=['csv'], key="ge_v5")
    ae = st.file_uploader("🔍 Autenticidade Entrada (XLSX)", type=['xlsx'], key="ae_v5")

with col_s:
    st.subheader("📤 FLUXO SAÍDAS")
    xs = st.file_uploader("📂 XMLs de Saída", type='xml', accept_multiple_files=True, key="xs_v5")
    gs = st.file_uploader("📊 Gerencial Saída (CSV)", type=['csv'], key="gs_v5")
    as_f = st.file_uploader("🔍 Autenticidade Saída (XLSX)", type=['xlsx'], key="as_v5")

if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    if not xe and not xs:
        st.warning("Por favor, suba ao menos um arquivo XML.")
    else:
        with st.spinner("🧡 O Sentinela está auditando os dados..."):
            try:
                df_xe = extrair_dados_xml(xe)
                df_xs = extrair_dados_xml(xs)
                relat = gerar_excel_final(df_xe, df_xs, u_base_unica, ae, as_f, ge, gs, cod_cliente)
                st.success("Auditoria concluída com sucesso! 🧡")
                st.download_button("💾 BAIXAR RELATÓRIO FINAL", relat, "Auditoria_Sentinela.xlsx", use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")
