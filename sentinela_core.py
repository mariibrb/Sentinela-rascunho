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
                    "CST_ICMS_XML": "", "ALIQ_ICMS_XML": 0.0,
                    "CST_PIS_XML": "", "CST_COFINS_XML": ""
                }
                if imp is not None:
                    icms = imp.find('.//ICMS')
                    if icms is not None:
                        for n in icms:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST_ICMS_XML"] = cst.text.zfill(2)
                            if n.find('pICMS') is not None: linha["ALIQ_ICMS_XML"] = float(n.find('pICMS').text)
                    
                    pis = imp.find('.//PIS')
                    if pis is not None:
                        for p in pis:
                            cst_p = p.find('CST')
                            if cst_p is not None: linha["CST_PIS_XML"] = cst_p.text.zfill(2)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, b_unica, ae, as_f, ge, gs, cod_cliente=""):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if b_unica is not None:
            try:
                # Carregando as abas
                df_icms_b = pd.read_excel(b_unica, sheet_name='ICMS')
                df_pc_b = pd.read_excel(b_unica, sheet_name='PIS_COFINS')
                
                def analisar(df_xml, aba_nome):
                    if df_xml.empty: return
                    # Cruzamento com ICMS (A-G) e PIS/COFINS (3 colunas)
                    df_res = pd.merge(df_xml, df_icms_b, left_on='NCM_XML', right_on='NCM', how='left')
                    df_res = pd.merge(df_res, df_pc_b, left_on='NCM_XML', right_on='NCM', how='left', suffixes=('', '_BASE'))
                    
                    # Exemplo de auditoria CST ICMS Interno
                    df_res['CHECK_CST_ICMS_INT'] = np.where(df_res['CST_ICMS_XML'] == df_res['CST (INTERNA)'].astype(str).str.zfill(2), "✅", "❌")
                    df_res.to_excel(writer, sheet_name=aba_nome, index=False)

                analisar(df_xe, 'AUDITORIA_ENTRADA')
                analisar(df_xs, 'AUDITORIA_SAIDA')
            except:
                if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)
        else:
            if not df_xe.empty: df_xe.to_excel(writer, sheet_name='XML_ENTRADAS', index=False)

        if ge: pd.read_csv(ge, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_ENT', index=False)
        if gs: pd.read_csv(gs, sep=None, engine='python').to_excel(writer, sheet_name='GERENCIAL_SAI', index=False)
        if ae: pd.read_excel(ae).to_excel(writer, sheet_name='AUTENTICIDADE_ENT', index=False)
        if as_f: pd.read_excel(as_f).to_excel(writer, sheet_name='AUTENTICIDADE_SAI', index=False)
            
    return output.getvalue()
