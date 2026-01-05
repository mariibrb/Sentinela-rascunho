import streamlit as st
import os, io, pandas as pd
import requests
from sentinela_core import extrair_dados_xml, gerar_excel_final

# 1. Configuração da Página (DEVE ser o primeiro comando)
st.set_page_config(page_title="Sentinela - Auditoria Fiscal", page_icon="🧡", layout="wide", initial_sidebar_state="expanded")

# 2. Estilo CSS Sentinela (Revisado para não quebrar a visualização)
st.markdown("""
<style>
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .stApp { background-color: #F7F7F7; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 2px solid #FF6F00; }
    h1, h2, h3 { color: #FF6F00 !important; font-weight: 700; text-align: center; }
    
    .stButton > button {
        background-color: #FF6F00 !important; color: white !important; border-radius: 25px !important;
        font-weight: bold !important; width: 100% !important; height: 50px !important; border: none !important;
    }
    
    .passo-container {
        background-color: #FFFFFF; padding: 15px; border-radius: 10px; border-left: 5px solid #FF6F00;
        margin-bottom: 20px; text-align: center; box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
    }
    .passo-texto { color: #FF6F00; font-size: 1.1rem; font-weight: 700; }
    .stFileUploader section { background-color: #FFFFFF; border: 1px dashed #FF6F00 !important; }
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

# --- SIDEBAR ---
with st.sidebar:
    # Tenta carregar imagem, mas não trava se não achar
    try: st.image(".streamlit/Sentinela.png", use_container_width=True)
    except: st.title("SENTINELA 🧡")
    
    st.markdown("---")
    def criar_gabarito():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            f_ncm = wb.add_format({'bg_color': '#444444', 'font_color': 'white', 'bold': True, 'border': 1})
            f_lar = wb.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
            for s, c_l in [('ICMS', ["NCM", "CST (INTERNA)", "ALIQ (INTERNA)"]), ('PIS_COFINS', ["NCM", "CST Entrada", "CST Saída"])]:
                pd.DataFrame(columns=c_l).to_excel(writer, sheet_name=s, index=False)
                for c, v in enumerate(c_l): writer.sheets[s].write(0, c, v, f_ncm if c == 0 else f_lar)
        return output.getvalue()
    st.download_button("📥 Baixar Gabarito", criar_gabarito(), "gabarito_base.xlsx", use_container_width=True)

# --- TELA PRINCIPAL ---
st.markdown("<div class='passo-container'><span class='passo-texto'>👣 PASSO 1: Selecionar Empresa</span></div>", unsafe_allow_html=True)
empresas = listar_empresas_no_github()
cod_cliente = st.selectbox("Selecione a empresa cadastrada no GitHub:", [""] + empresas)

if cod_cliente:
    st.markdown("<div class='passo-container'><span class='passo-texto'>👣 PASSO 2: Carregar Documentos</span></div>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📤 XMLs SAÍDA")
        xs = st.file_uploader("Arraste os XMLs aqui", type='xml', accept_multiple_files=True, key="xs_v68")
        as_f = st.file_uploader("Relatório Autenticidade Saída", type=['xlsx'], key="as_v68")
    
    with c2:
        st.subheader("📥 XMLs ENTRADA")
        xe = st.file_uploader("Arraste os XMLs aqui", type='xml', accept_multiple_files=True, key="xe_v68")
        ae = st.file_uploader("Relatório Autenticidade Entrada", type=['xlsx'], key="ae_v68")

    if st.button("🚀 EXECUTAR AUDITORIA"):
        if not xs:
            st.warning("Carregue ao menos os XMLs de Saída.")
        else:
            with st.spinner("🧡 Sentinela auditando..."):
                try:
                    df_xe = extrair_dados_xml(xe)
                    df_xs = extrair_dados_xml(xs)
                    relat = gerar_excel_final(df_xe, df_xs, ae, as_f, cod_cliente)
                    st.success("Auditoria Concluída!")
                    st.download_button("💾 BAIXAR AGORA", relat, f"Sentinela_{cod_cliente}.xlsx", use_container_width=True)
                except Exception as e:
                    st.error(f"Erro: {str(e)}")
