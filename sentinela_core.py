import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def buscar_base_github(cod_cliente):
    token = st.secrets.get("GITHUB_TOKEN")
    repo = st.secrets.get("GITHUB_REPO")
    if not token or not cod_cliente: return None
    url = f"https://api.github.com/repos/{repo}/contents/Bases_Tributárias"
    headers = {"Authorization": f"token {token}"}
    try:
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            for item in res.json():
                if item['name'].startswith(str(cod_cliente)):
                    f_res = requests.get(item['download_url'], headers=headers)
                    return io.BytesIO(f_res.content)
    except: pass
    return None

def extrair_dados_xml(files):
    dados_lista = []
    if not files: return pd.DataFrame()
    for f in files:
        try:
            f.seek(0)
            root = ET.fromstring(re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace')))
            chave = root.find('.//infNFe').attrib.get('Id', '')[3:] if root.find('.//infNFe') is not None else ""
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": root.find('.//nNF').text if root.find('.//nNF') is not None else "",
                    "NCM_XML": re.sub(r'\D', '', prod.find('NCM').text).zfill(8) if prod.find('NCM') is not None else "",
                    "CST_ICMS_XML": "", "ALIQ_ICMS_XML": 0.0, "UF_DESTINO": root.find('.//dest/enderDest/UF').text if root.find('.//dest/enderDest/UF') is not None else ""
                }
                if imp is not None:
                    icms = imp.find('.//ICMS')
                    if icms is not None:
                        for n in icms:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST_ICMS_XML"] = cst.text.zfill(2)
                            if n.find('pICMS') is not None: linha["ALIQ_ICMS_XML"] = float(n.find('pICMS').text)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_unica, ae, as_f, ge, gs, cod_cliente=""):
    if b_unica is None and cod_cliente:
        b_unica = buscar_base_github(cod_cliente)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if b_unica:
            try:
                df_icms_b = pd.read_excel(b_unica, sheet_name='ICMS')
                def analisar(df_xml, aba):
                    if df_xml.empty: return
                    df_res = pd.merge(df_xml, df_icms_b, left_on='NCM_XML', right_on='NCM', how='left')
                    # Auditoria Interestadual focada apenas no CST agora
                    df_res['AUDIT_CST_ICMS'] = np.where(df_res['CST_ICMS_XML'] == df_res['CST (INTERNA)'].astype(str).str.zfill(2), "✅", "❌")
                    df_res.to_excel(writer, sheet_name=aba, index=False)
                analisar(df_xe, 'AUDITORIA_ENTRADA')
                analisar(df_xs, 'AUDITORIA_SAIDA')
            except: pass
        else:
            if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
        # Outras abas (Gerencial/Autenticidade)
        if ge: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENT', index=False)
        if gs: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAI', index=False)
    return output.getvalue()
