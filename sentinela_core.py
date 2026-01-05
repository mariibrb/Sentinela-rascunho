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
                    "ITEM": det.attrib.get('nItem', '0'),
                    "NCM_XML": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "DESCRICAO_XML": buscar('xProd', prod),
                    "VALOR_PRODUTO": float(buscar('vProd', prod) or 0),
                    "CST_ICMS_XML": "", "ALIQ_ICMS_XML": 0.0, "VLR_ICMS_XML": 0.0,
                    "VLR_IPI_XML": 0.0, "VLR_PIS_XML": 0.0, "VLR_COFINS_XML": 0.0
                }
                if imp is not None:
                    icms = imp.find('.//ICMS')
                    if icms is not None:
                        for n in icms:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST_ICMS_XML"] = cst.text.zfill(2)
                            if n.find('pICMS') is not None: linha["ALIQ_ICMS_XML"] = float(n.find('pICMS').text)
                            if n.find('vICMS') is not None: linha["VLR_ICMS_XML"] = float(n.find('vICMS').text)
                    
                    ipi = imp.find('.//IPI/vIPI')
                    if ipi is not None: linha["VLR_IPI_XML"] = float(ipi.text)
                    pis = imp.find('.//PIS//vPIS')
                    if pis is not None: linha["VLR_PIS_XML"] = float(pis.text)
                    cofins = imp.find('.//COFINS//vCOFINS')
                    if cofins is not None: linha["VLR_COFINS_XML"] = float(cofins.text)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_unica, ge, gs, cod_cliente=""):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Lógica de Auditoria (Cruzamento)
        if b_unica is not None:
            df_base = pd.read_excel(b_unica)
            df_base['NCM_LINK'] = df_base['NCM'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(8)
            
            def processar_analise(df_xml, aba_nome):
                if df_xml.empty: return
                # Procura o NCM da Nota na sua Base de Auditoria
                df_cruzado = pd.merge(df_xml, df_base, left_on='NCM_XML', right_on='NCM_LINK', how='left')
                
                # Colunas de Status de Auditoria
                df_cruzado['AUDIT_ICMS_CST'] = np.where(df_cruzado['CST_ICMS_XML'] == df_cruzado['CST'].astype(str).str.zfill(2), "✅", "❌")
                df_cruzado['AUDIT_ICMS_ALIQ'] = np.where(df_cruzado['ALIQ_ICMS_XML'] == df_cruzado['ALÍQUOTA ICMS'], "✅", "❌")
                df_cruzado['AUDIT_PIS_COFINS'] = "CONFERIR" # Base para sua análise N-P
                
                df_cruzado.to_excel(writer, sheet_name=aba_nome, index=False)

            processar_analise(df_xe, 'AUDITORIA_ENTRADA')
            processar_analise(df_xs, 'AUDITORIA_SAIDA')
        else:
            if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
            if not df_xs.empty: df_xs.to_excel(writer, sheet_name='XML_SAIDAS', index=False)

        # Abas Gerenciais
        if ge: 
            try: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENT', index=False)
            except: pass
        if gs: 
            try: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAI', index=False)
            except: pass
            
    return output.getvalue()
