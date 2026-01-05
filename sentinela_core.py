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
            
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'),
                    "DATA_EMISSAO": buscar('dhEmi')[:10] if buscar('dhEmi') else "",
                    "ITEM": det.attrib.get('nItem', '0'), "CFOP": buscar('CFOP', prod),
                    "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "COD_PROD": buscar('cProd', prod), "DESCR": buscar('xProd', prod),
                    "VPROD": float(buscar('vProd', prod) or 0),
                    "CST-ICMS": "", "BC-ICMS": 0.0, "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0,
                    "VLR-PIS": 0.0, "VLR-COFINS": 0.0
                }
                if imp is not None:
                    ic = imp.find('.//ICMS')
                    if ic is not None:
                        for n in ic:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST-ICMS"] = cst.text.zfill(2)
                            if n.find('vBC') is not None: linha["BC-ICMS"] = float(n.find('vBC').text)
                            if n.find('vICMS') is not None: linha["VLR-ICMS"] = float(n.find('vICMS').text)
                    
                    pis = imp.find('.//PIS')
                    if pis is not None:
                        vP = pis.find('.//vPIS')
                        if vP is not None: linha["VLR-PIS"] = float(vP.text)
                    
                    cofins = imp.find('.//COFINS')
                    if cofins is not None:
                        vC = cofins.find('.//vCOFINS')
                        if vC is not None: linha["VLR-COFINS"] = float(vC.text)
                            
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_icms=None, b_pc=None, ae=None, as_f=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. Bases Brutas
        if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
        if not df_xs.empty: df_xs.to_excel(writer, sheet_name='XML_SAIDAS', index=False)
        
        # 2. Processar Bases da Sidebar (ICMS, PIS, COFINS)
        if b_icms is not None:
            try:
                df_base_icms = pd.read_excel(b_icms)
                df_base_icms.to_excel(writer, sheet_name='BASE_CONFERENCIA_ICMS', index=False)
                # Cruzamento: XML Saída vs Base ICMS
                if not df_xs.empty:
                    df_audit_icms = pd.merge(df_xs, df_base_icms, left_on='NCM', right_on=df_base_icms.columns[0], how='left')
                    df_audit_icms.to_excel(writer, sheet_name='DIVERGENCIAS_ICMS', index=False)
            except: pass

        if b_pc is not None:
            try:
                df_base_pc = pd.read_excel(b_pc)
                df_base_pc.to_excel(writer, sheet_name='BASE_PIS_COFINS', index=False)
            except: pass
            
        # 3. Autenticidade
        if ae:
            try: pd.read_excel(ae).to_excel(writer, sheet_name='VALIDACAO_AUTENTIC_ENT', index=False)
            except: pass
            
    return output.getvalue()
