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
                    "COD_PROD": buscar('cProd', prod), "DESCRICAO": buscar('xProd', prod),
                    "VALOR_PRODUTO": float(buscar('vProd', prod) or 0),
                    "CST_ICMS": "", "BC_ICMS": 0.0, "VLR_ICMS": 0.0, "VLR_IPI": 0.0, "VLR_PIS": 0.0, "VLR_COFINS": 0.0,
                    "CENQ_IPI": ""
                }
                if imp is not None:
                    # ICMS
                    icms_tag = imp.find('.//ICMS')
                    if icms_tag is not None:
                        for n in icms_tag:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST_ICMS"] = cst.text.zfill(2)
                            if n.find('vBC') is not None: linha["BC_ICMS"] = float(n.find('vBC').text)
                            if n.find('vICMS') is not None: linha["VLR_ICMS"] = float(n.find('vICMS').text)
                    
                    # IPI (Extração do Valor e Código de Enquadramento)
                    ipi_tag = imp.find('.//IPI')
                    if ipi_tag is not None:
                        vIPI = ipi_tag.find('.//vIPI')
                        if vIPI is not None: linha["VLR_IPI"] = float(vIPI.text)
                        cenq = ipi_tag.find('.//cEnq')
                        if cenq is not None: linha["CENQ_IPI"] = cenq.text

                    # PIS/COFINS
                    p = imp.find('.//PIS'); c = imp.find('.//COFINS')
                    if p is not None and p.find('.//vPIS') is not None: linha["VLR_PIS"] = float(p.find('.//vPIS').text)
                    if c is not None and c.find('.//vCOFINS') is not None: linha["VLR_COFINS"] = float(c.find('.//vCOFINS').text)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_icms, b_pc, ae, as_f, ge, gs, b_ipi, cod_cliente=""):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
        if not df_xs.empty: df_xs.to_excel(writer, sheet_name='XML_SAIDAS', index=False)
        
        if ge:
            try: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENTRADA', index=False)
            except: pass
        if gs:
            try: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAIDA', index=False)
            except: pass

        if b_icms is not None:
            try: pd.read_excel(b_icms).to_excel(writer, sheet_name='BASE_REF_ICMS', index=False)
            except: pass
        if b_ipi is not None:
            try: pd.read_excel(b_ipi).to_excel(writer, sheet_name='BASE_REF_IPI_TIPI', index=False)
            except: pass
        if b_pc is not None:
            try: pd.read_excel(b_pc).to_excel(writer, sheet_name='BASE_REF_PIS_COFINS', index=False)
            except: pass
            
    return output.getvalue()
