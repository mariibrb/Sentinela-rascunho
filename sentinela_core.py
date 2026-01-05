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
    if not files: return pd.DataFrame() # Não impede o fluxo se não houver XML
    dados_lista = []
    for f in files:
        try:
            f.seek(0)
            root = ET.fromstring(re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace')))
            chave = root.find('.//infNFe').attrib.get('Id', '')[3:] if root.find('.//infNFe') is not None else ""
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NCM_XML": re.sub(r'\D', '', prod.find('NCM').text).zfill(8) if prod.find('NCM') is not None else "",
                    "CST_ICMS_XML": "", "ALIQ_ICMS_XML": 0.0
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
    avisos = [] # Lista para armazenar o que faltou

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Tenta Auditoria se houver Base e XML
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
            except: avisos.append("Erro ao ler abas da Base Tributária.")
        else:
            if not base_final: avisos.append(f"Base tributária do cliente {cod_cliente} não localizada no GitHub.")
            if df_xe.empty and df_xs.empty: avisos.append("Nenhum arquivo XML foi enviado para análise.")

        # Processa Gerenciais e Autenticidade (se existirem)
        if ge: 
            try: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENT', index=False)
            except: avisos.append("Não foi possível processar o arquivo Gerencial Entrada.")
        if gs:
            try: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAI', index=False)
            except: avisos.append("Não foi possível processar o arquivo Gerencial Saída.")
        if ae: pd.read_excel(ae).to_excel(writer, sheet_name='AUTENTICIDADE_ENT', index=False)
        if as_f: pd.read_excel(as_f).to_excel(writer, sheet_name='AUTENTICIDADE_SAI', index=False)

        # Aba de Status/Avisos para o usuário
        if not avisos: avisos = ["Todos os arquivos enviados foram processados com sucesso!"]
        pd.DataFrame({"Mensagens do Sentinela": avisos}).to_excel(writer, sheet_name='STATUS_PROCESSAMENTO', index=False)
            
    return output.getvalue()
