import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re
import io

def extrair_dados_xml(files):
    dados_lista = []
    if not files: return pd.DataFrame()
    for f in files:
        try:
            f.seek(0)
            conteudo = f.read().decode('utf-8', errors='replace')
            root = ET.fromstring(re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', conteudo))
            
            def buscar(caminho, raiz=root):
                alvo = raiz.find(f'.//{caminho}')
                return alvo.text if alvo is not None and alvo.text is not None else ""

            inf_nfe = root.find('.//infNFe')
            chave = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
            emit = root.find('.//emit')
            
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'),
                    "DATA_EMISSAO": buscar('dhEmi')[:10] if buscar('dhEmi') else "",
                    "CNPJ_EMIT": buscar('CNPJ', emit), "ITEM": det.attrib.get('nItem', '0'),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "COD_PROD": buscar('cProd', prod), "DESCR": buscar('xProd', prod),
                    "VPROD": float(buscar('vProd', prod) or 0),
                    "CST-ICMS": "", "BC-ICMS": 0.0, "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0
                }
                if imp is not None:
                    ic = imp.find('.//ICMS')
                    if ic is not None:
                        for n in ic:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST-ICMS"] = cst.text.zfill(2)
                            if n.find('vBC') is not None: linha["BC-ICMS"] = float(n.find('vBC').text)
                            if n.find('vICMS') is not None: linha["VLR-ICMS"] = float(n.find('vICMS').text)
                            if n.find('pICMS') is not None: linha["ALQ-ICMS"] = float(n.find('pICMS').text)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, ge=None, gs=None, ae=None, as_f=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
        if not df_xs.empty: df_xs.to_excel(writer, sheet_name='XML_SAIDAS', index=False)
        if ae:
            try: pd.read_excel(ae).to_excel(writer, sheet_name='AUTENTICIDADE_ENT', index=False)
            except: pass
        if as_f:
            try: pd.read_excel(as_f).to_excel(writer, sheet_name='AUTENTICIDADE_SAI', index=False)
            except: pass
        if not df_xs.empty:
            resumo = df_xs.groupby('CFOP').agg({'VPROD': 'sum', 'VLR-ICMS': 'sum'}).reset_index()
            resumo.to_excel(writer, sheet_name='ANALISE_RESUMO_FISCAL', index=False)
    return output.getvalue()
