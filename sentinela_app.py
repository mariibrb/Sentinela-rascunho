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
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; margin-bottom: 5px; }
    
    /* Centralização de Imagens na Sidebar */
    [data-testid="stSidebar"] [data-testid="stImage"] {
        display: block;
        margin-left: auto;
        margin-right: auto;
        text-align: center;
    }

    /* Estilo do Botão e Centralização */
    .stButton {
        display: flex;
        justify-content: center;
    }
    .stButton>button {
        background-color: #FF6F00; color: white !important;
        border-radius: 25px !important; font-weight: bold; 
        width: 300px; height: 50px; border: none;
    }
    .stButton>button:hover { background-color: #E65100; transform: scale(1.02); }
    
    /* Passos Delicados */
    .passo-container {
        background-color: #FFFFFF;
        padding: 8px 15px;
        border-radius: 10px;
        border-left: 4px solid #FF6F00;
        margin: 10px auto 15px auto;
        max-width: 650px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        text-align: center;
    }
    .passinho { color: #808080; font-size: 1.1rem; margin-right: 8px; }
    .passo-texto { color: #FF6F00; font-size: 1rem; font-weight: 700; }

    .stFileUploader section { background-color: #FFFFFF; border: 1px dashed #FF6F00 !important; border-radius: 12px !important; }
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
            return sorted(list(set([f['name'].split('-')[0] for f in arquivos if f['name'].endswith('.xlsx')])))
    except: pass
    return []

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists(".streamlit/Sentinela.png"):
        st.image(".streamlit/Sentinela.png", use_container_width=True)
    
    if os.path.exists(".streamlit/nascel sem fundo.png"):
        st.image(".streamlit/nascel sem fundo.png", width=140)
    
    st.markdown("---")
    st.subheader("📥 Gabaritos")
    
    def criar_gabarito_nascel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            f_ncm = workbook.add_format({'bg_color': '#444444', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar_e = workbook.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar_c = workbook.add_format({'bg_color': '#FFB74D', 'bold': True, 'border': 1})
            f_cin_c = workbook.add_format({'bg_color': '#E0E0E0', 'bold': True, 'border': 1})

            cols_icms = ["NCM", "CST (INTERNA)", "ALIQ (INTERNA)", "CST (ESTADUAL)"]
            pd.DataFrame(columns=cols_icms).to_excel(writer, sheet_name='ICMS', index=False)
            for c, v in enumerate(cols_icms): writer.sheets['ICMS'].write(0, c, v, f_ncm if c == 0 else (f_lar_e if c <= 2 else f_lar_c))

            cols_pc = ["NCM", "CST Entrada", "CST Saída"]
            pd.DataFrame(columns=cols_pc).to_excel(writer, sheet_name='PIS_COFINS', index=False)
            for c, v in enumerate(cols_pc): writer.sheets['PIS_COFINS'].write(0, c, v, f_ncm if c == 0 else f_cin_c)
        return output.getvalue()

    st.download_button("📥 Baixar Gabarito", criar_gabarito_nascel(), "gabarito_nascel.xlsx", use_container_width=True)
    st.markdown("---")
    st.subheader("🔄 Base de Referência")
    if st.file_uploader("Upload da Base", type=['xlsx'], key='base_construcao'): 
        st.error("🚧 CAMPO EM CONSTRUÇÃO")

# --- 4. TELA PRINCIPAL ---

# PASSO 1
st.markdown("<div class='passo-container'><span class='passinho'>👣</span><span class='passo-texto'>PASSO 1: Selecione o cliente</span></div>", unsafe_allow_html=True)
col_c = st.columns([1, 1.5, 1])
with col_c[1]:
    cod_cliente = st.selectbox("Lista de empresas:", [""] + listar_empresas_no_github(), label_visibility="collapsed")

if cod_cliente:
    # PASSO 2
    st.markdown("<div class='passo-container'><span class='passinho'>👣</span><span class='passo-texto'>PASSO 2: Inclua os arquivos disponíveis</span></div>", unsafe_allow_html=True)
    c_e, c_s = st.columns(2, gap="medium")
    with c_e:
        st.subheader("📥 ENTRADAS")
        xe = st.file_uploader("XMLs", type='xml', accept_multiple_files=True, key="xe_v21")
        ge = st.file_uploader("Gerencial", type=['csv'], key="ge_v21")
        ae = st.file_uploader("Autenticidade", type=['xlsx'], key="ae_v21")
    with c_s:
        st.subheader("📤 SAÍDAS")
        xs = st.file_uploader("XMLs", type='xml', accept_multiple_files=True, key="xs_v21")
        gs = st.file_uploader("Gerencial", type=['csv'], key="gs_v21")
        as_f = st.file_uploader("Autenticidade", type=['xlsx'], key="as_v21")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Botão de Ação Centralizado e Renomeado
    if st.button("🚀 GERAR RELATÓRIO", type="primary"):
        with st.spinner("🧡 Processando..."):
            try:
                df_xe = extrair_dados_xml(xe); df_xs = extrair_dados_xml(xs)
                relat = gerar_excel_final(df_xe, df_xs, None, ae, as_f, ge, gs, cod_cliente)
                st.success("Relatório gerado com sucesso! 🧡")
                st.download_button("💾 BAIXAR AGORA", relat, f"Relatorio_{cod_cliente}.xlsx", use_container_width=True)
            except Exception as e: st.error(f"Erro: {e}")
