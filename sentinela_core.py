import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def buscar_base_no_github(cod_cliente):
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
    if not files: return pd.DataFrame() 
    dados_lista = []
    for f in files:
        try:
            f.seek(0)
            root = ET.fromstring(re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace')))
            inf_nfe = root.find('.//infNFe')
            chave = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, 
                    "NCM_XML": re.sub(r'\D', '', prod.find('NCM').text).zfill(8) if prod.find('NCM') is not None else ""
                }
                if imp is not None:
                    icms = imp.find('.//ICMS')
                    if icms is not None:
                        for n in icms:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST_ICMS_XML"] = cst.text.zfill(2)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_unica, ae, as_f, ge, gs, cod_cliente=""):
    base_final = buscar_base_no_github(cod_cliente)
    output = io.BytesIO()
    avisos = [] 

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # ABA MANUAL DE INSTRUÇÕES (Primeira Aba)
        df_manual = pd.DataFrame({
            "Aba": ["MANUAL", "RESUMO", "AUDITORIA_ENTRADA/SAIDA", "GERENCIAL", "AUTENTICIDADE"],
            "Objetivo": [
                "Guia de uso do relatório.",
                "Status do processamento e avisos importantes.",
                "Cruzamento fiscal: validação de CST e NCM.",
                "Espelhamento dos dados gerenciais importados.",
                "Dados de autenticidade/protocolos."
            ]
        })
        df_manual.to_excel(writer, sheet_name='MANUAL', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['MANUAL']
        header_format = workbook.add_format({'bg_color': '#FF6F00', 'font_color': 'white', 'bold': True, 'border': 1})
        for col_num, value in enumerate(df_manual.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.set_column('A:B', 30)

        if base_final and (not df_xe.empty or not df_xs.empty):
            try:
                df_icms_b = pd.read_excel(base_final, sheet_name='ICMS')
                def analisar(df_xml, aba):
                    if df_xml.empty: return
                    df_res = pd.merge(df_xml, df_icms_b, left_on='NCM_XML', right_on='NCM', how='left')
                    df_res['CHECK_CST'] = np.where(df_res['CST_ICMS_XML'] == df_res['CST (INTERNA)'].astype(str).str.zfill(2), "✅", "❌")
                    df_res.to_excel(writer, sheet_name=aba, index=False)
                analisar(df_xe, 'AUDITORIA_ENTRADA')
                analisar(df_xs, 'AUDITORIA_SAIDA')
            except: avisos.append("Erro ao ler base tributária.")

        if ge: 
            try: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENT', index=False)
            except: pass
        if gs:
            try: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAI', index=False)
            except: pass
        if ae: pd.read_excel(ae).to_excel(writer, sheet_name='AUTENTICIDADE_ENT', index=False)
        if as_f: pd.read_excel(as_f).to_excel(writer, sheet_name='AUTENTICIDADE_SAI', index=False)

        pd.DataFrame({"STATUS": avisos if avisos else ["Sucesso"]}).to_excel(writer, sheet_name='RESUMO', index=False)
            
    return output.getvalue()
