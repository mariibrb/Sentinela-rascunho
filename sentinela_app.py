import streamlit as st
import os, io, pandas as pd
import requests
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página
st.set_page_config(page_title="Sentinela Nascel", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Nascel
st.markdown("""
<style>
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    .stButton>button, .stDownloadButton>button {
        background-color: #FF6F00; color: white !important;
        border-radius: 25px !important; font-weight: bold; width: 100%; height: 45px; border: none;
    }
    .stFileUploader section { background-color: #FFFFFF; border: 2px dashed #FF6F00 !important; border-radius: 15px !important; }
</style>
""", unsafe_allow_html=True)

def listar_empresas_no_github():
    token = st.secrets.get("GITHUB_TOKEN")
    repo = st.secrets.get("GITHUB_REPO")
    if not token or not repo: return []
    url = f"https://api.github.com/repos/{repo}/contents/Bases_Tributárias"
    headers = {"Authorization": f"token {token}"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            arquivos = response.json()
            empresas = sorted(list(set([f['name'].split('-')[0] for f in arquivos if f['name'].endswith('.xlsx')])))
            return empresas
    except: pass
    return []

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", use_container_width=True)
    
    st.markdown("---")
    st.subheader("🏢 Seleção de Cliente")
    opcoes_empresas = listar_empresas_no_github()
    if opcoes_empresas:
        cod_cliente = st.selectbox("Selecione a Empresa", [""] + opcoes_empresas)
    else:
        st.warning("⚠️ Verifique a conexão com o GitHub.")
        cod_cliente = ""

    st.markdown("---")
    st.subheader("📥 Gabarito")
    
    def criar_gabarito_nascel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            f_ncm = workbook.add_format({'bg_color': '#444444', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar_e = workbook.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar_c = workbook.add_format({'bg_color': '#FFB74D', 'bold': True, 'border': 1})
            f_cin_e = workbook.add_format({'bg_color': '#757575', 'font_color': 'white', 'bold': True, 'border': 1})
            f_cin_c = workbook.add_format({'bg_color': '#E0E0E0', 'bold': True, 'border': 1})

            cols_icms = ["NCM", "CST (INTERNA)", "ALIQ (INTERNA)", "CST (ESTADUAL)"]
            pd.DataFrame(columns=cols_icms).to_excel(writer, sheet_name='ICMS', index=False)
            ws_i = writer.sheets['ICMS']
            ws_i.set_tab_color('#FF6F00')
            for c, v in enumerate(cols_icms):
                ws_i.write(0, c, v, f_ncm if c == 0 else (f_lar_e if c <= 2 else f_lar_c))

            cols_ipi = ["NCM_TIPI", "EX", "DESCRIÇÃO", "ALÍQUOTA (%)"]
            pd.DataFrame(columns=cols_ipi).to_excel(writer, sheet_name='IPI', index=False)
            writer.sheets['IPI'].set_tab_color('#757575')
            for c, v in enumerate(cols_ipi): writer.sheets['IPI'].write(0, c, v, f_ncm if c == 0 else f_cin_e)

            cols_pc = ["NCM", "CST Entrada", "CST Saída"]
            pd.DataFrame(columns=cols_pc).to_excel(writer, sheet_name='PIS_COFINS', index=False)
            ws_pc = writer.sheets['PIS_COFINS']
            ws_pc.set_tab_color('#E0E0E0')
            for c, v in enumerate(cols_pc): ws_pc.write(0, c, v, f_ncm if c == 0 else f_cin_c)
        return output.getvalue()

    st.download_button("📥 Baixar Gabarito Nascel", criar_gabarito_nascel(), "gabarito_nascel.xlsx", use_container_width=True)

    # CAMPO EM CONSTRUÇÃO (Abaixo do Gabarito)
    st.markdown("---")
    st.subheader("🔄 Base de Referência")
    base_bloqueada = st.file_uploader("Upload da Base (Opcional)", type=['xlsx'], key='base_construcao')
    if base_bloqueada:
        st.error("🚧 CAMPO EM CONSTRUÇÃO: O arquivo subido não será utilizado nesta versão.")

# --- 4. TELA PRINCIPAL ---
st.markdown("---")
col_e, col_s = st.columns(2, gap="large")
with col_e:
    st.subheader("📥 FLUXO ENTRADAS")
    xe = st.file_uploader("📂 XMLs de Entrada", type='xml', accept_multiple_files=True, key="xe_v14")
    ge = st.file_uploader("📊 Gerencial Entrada (CSV)", type=['csv'], key="ge_v14")
    ae = st.file_uploader("🔍 Autenticidade Entrada (XLSX)", type=['xlsx'], key="ae_v14")
with col_s:
    st.subheader("📤 FLUXO SAÍDAS")
    xs = st.file_uploader("📂 XMLs de Saída", type='xml', accept_multiple_files=True, key="xs_v14")
    gs = st.file_uploader("📊 Gerencial Saída (CSV)", type=['csv'], key="gs_v14")
    as_f = st.file_uploader("🔍 Autenticidade Saída (XLSX)", type=['xlsx'], key="as_v14")

if st.button("🚀 EXECUTAR AUDITORIA COMPLETA", type="primary"):
    if not xe and not xs: st.warning("Suba ao menos um XML.")
    elif not cod_cliente: st.warning("Por favor, selecione uma empresa na lista.")
    else:
        with st.spinner("🧡 Auditando com base no GitHub..."):
            try:
                df_xe = extrair_dados_xml(xe); df_xs = extrair_dados_xml(xs)
                # Passamos None para b_unica para garantir que use apenas a do GitHub
                relat = gerar_excel_final(df_xe, df_xs, None, ae, as_f, ge, gs, cod_cliente)
                st.success(f"Auditoria {cod_cliente} finalizada! 🧡")
                st.download_button("💾 BAIXAR RELATÓRIO", relat, f"Relatorio_{cod_cliente}.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Erro: {e}")
