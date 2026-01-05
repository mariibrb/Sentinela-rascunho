import streamlit as st
import os, io, pandas as pd
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela Nascel", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Nascel (Botões e Uploads Bonitinhos)
st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    
    /* Títulos */
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    
    /* Botões de Ação e Download */
    .stButton>button, .stDownloadButton>button {
        background-color: #FF6F00;
        color: white !important;
        border-radius: 25px !important;
        font-weight: bold;
        width: 100%;
        height: 45px;
        border: none;
        transition: 0.3s;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #E65100;
        transform: scale(1.02);
    }

    /* Customização dos Campos de Upload */
    .stFileUploader section {
        background-color: #FFFFFF;
        border: 2px dashed #FF6F00 !important;
        border-radius: 15px !important;
        padding: 10px !important;
    }
    .stFileUploader label {
        color: #555 !important;
        font-weight: 600 !important;
    }
    
    /* Ajuste de Margens */
    .block-container { padding-top: 1rem !important; }
</style>
""", unsafe_allow_html=True)

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", use_container_width=True)
    
    st.markdown("---")
    st.subheader("🏢 Identificação")
    cod_cliente = st.text_input("Código do Cliente (ex: 394)", key="cod_cli")

    st.markdown("---")
    st.subheader("🔄 Bases de Referência")
    u_base_unica = st.file_uploader("Upload da Base de Auditoria", type=['xlsx'], key='base_unica_v7')
    
    st.markdown("---")
    st.subheader("📥 Gabaritos")
    
    def criar_gabarito_multi_abas():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            f_laranja = workbook.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            f_cinza_e = workbook.add_format({'bg_color': '#757575', 'font_color': 'white', 'bold': True, 'border': 1})
            f_cinza_c = workbook.add_format({'bg_color': '#E0E0E0', 'bold': True, 'border': 1})

            # Aba ICMS
            cols_icms = ["NCM", "BASE REDUZIDA", "CST", "ALÍQUOTA ICMS", "PERC_REDUÇÃO", "BASE REDUZIDA_EST", "CST_EST", "ALÍQUOTA_EST"]
            pd.DataFrame(columns=cols_icms).to_excel(writer, sheet_name='ICMS', index=False)
            ws_icms = writer.sheets['ICMS']
            ws_icms.set_tab_color('#FF6F00')
            for c, v in enumerate(cols_icms): ws_icms.write(0, c, v, f_laranja)

            # Aba IPI
            cols_ipi = ["NCM_TIPI", "EX", "DESCRIÇÃO", "ALÍQUOTA (%)"]
            pd.DataFrame(columns=cols_ipi).to_excel(writer, sheet_name='IPI', index=False)
            ws_ipi = writer.sheets['IPI']
            ws_ipi.set_tab_color('#757575')
            for c, v in enumerate(cols_ipi): ws_ipi.write(0, c, v, f_cinza_e)

            # Aba PIS_COFINS
            cols_pc = ["NCM_PC", "Entrada", "Saída", "CFOP-CST", "Status"]
            pd.DataFrame(columns=cols_pc).to_excel(writer, sheet_name='PIS_COFINS', index=False)
            ws_pc = writer.sheets['PIS_COFINS']
            ws_pc.set_tab_color('#E0E0E0')
            for c, v in enumerate(cols_pc): ws_pc.write(0, c, v, f_cinza_c)

        return output.getvalue()

    st.download_button("📥 Baixar Gabarito Nascel", criar_gabarito_multi_abas(), "gabarito_nascel_v7.xlsx", use_container_width=True)

# --- 4. TELA PRINCIPAL ---
c1, c2, c3 = st.columns([1.2, 1, 1.2]) 
with c2:
    if os.path.exists(".streamlit/Sentinela.png"): st.image(".streamlit/Sentinela.png", use_container_width=True)

st.markdown("---")
col_e, col_s = st.columns(2, gap="large")
with col_e:
    st.subheader("📥 FLUXO ENTRADAS")
    xe = st.file_uploader("📂 XMLs de Entrada", type='xml', accept_multiple_files=True, key="xe_v7")
    ge = st.file_uploader("📊 Gerencial Entrada (CSV)", type=['csv'], key="ge_v7")
    ae = st.file_uploader("🔍 Autenticidade Entrada (XLSX)", type=['xlsx'], key="ae_v7")

with col_s:
    st.subheader("📤 FLUXO SAÍDAS")
    xs = st.file_uploader("📂 XMLs de Saída", type='xml', accept_multiple_files=True, key="xs_v7")
    gs = st.file_uploader("📊 Gerencial Saída (CSV)", type=['csv'], key="gs_v7")
    as_f = st.file_uploader("🔍 Autenticidade Saída (XLSX)", type=['xlsx'], key="as_v7")

if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    if not xe and not xs: st.warning("Por favor, suba ao menos um arquivo XML.")
    else:
        with st.spinner("🧡 O Sentinela está auditando os dados..."):
            try:
                df_xe = extrair_dados_xml(xe); df_xs = extrair_dados_xml(xs)
                relat = gerar_excel_final(df_xe, df_xs, u_base_unica, ae, as_f, ge, gs, cod_cliente)
                st.success("Auditoria concluída com sucesso! 🧡")
                st.download_button("💾 BAIXAR RELATÓRIO FINAL", relat, "Auditoria_Sentinela.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Erro: {e}")
