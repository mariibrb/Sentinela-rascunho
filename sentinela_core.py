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
    dados_lista = []
    if not files: return pd.DataFrame()
    for f in files:
        try:
            f.seek(0)
            root = ET.fromstring(re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace')))
            def buscar(caminho, raiz=root):
                alvo = raiz.find(f'.//{caminho}')
                return alvo.text if alvo is not None and alvo.text is not None else ""
            inf_nfe = root.find('.//infNFe')
            chave_acesso = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                ncm_limpo = re.sub(r'\D', '', buscar('NCM', prod)).zfill(8)
                linha = {
                    "CHAVE_ACESSO": chave_acesso, "NUM_NF": buscar('nNF'),
                    "DATA_EMISSAO": pd.to_datetime(buscar('dhEmi')).replace(tzinfo=None) if buscar('dhEmi') else None,
                    "UF_EMIT": buscar('UF', root.find('.//emit')), "UF_DEST": buscar('UF', root.find('.//dest')),
                    "AC": int(det.attrib.get('nItem', '0')), "CFOP": buscar('CFOP', prod), "NCM": ncm_limpo,
                    "COD_PROD": buscar('cProd', prod), "DESCR": buscar('xProd', prod),
                    "VPROD": float(buscar('vProd', prod)) if buscar('vProd', prod) else 0.0,
                    "CST-ICMS": "", "BC-ICMS": 0.0, "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0, "ICMS-ST": 0.0,
                    "CST-PIS": "", "CST-COF": "", "VAL-PIS": 0.0, "VAL-COF": 0.0, "BC-FED": 0.0,
                    "CST-IPI": "", "VAL-IPI": 0.0, "BC-IPI": 0.0, "ALQ-IPI": 0.0, "VAL-DIFAL": 0.0,
                    "VAL-FCP": 0.0, "VAL-FCPST": 0.0 
                }
                if imp is not None:
                    icms_n = imp.find('.//ICMS')
                    if icms_n is not None:
                        for n in icms_n:
                            cst = n.find('CST') or n.find('CSOSN')
                            if cst is not None: linha["CST-ICMS"] = cst.text.zfill(2)
                            if n.find('vBC') is not None: linha["BC-ICMS"] = float(n.find('vBC').text)
                            if n.find('vICMS') is not None: linha["VLR-ICMS"] = float(n.find('vICMS').text)
                            if n.find('pICMS') is not None: linha["ALQ-ICMS"] = float(n.find('pICMS').text)
                            if n.find('vICMSST') is not None: linha["ICMS-ST"] = float(n.find('vICMSST').text)
                            if n.find('vFCP') is not None: linha["VAL-FCP"] = float(n.find('vFCP').text)
                            if n.find('vFCPST') is not None: linha["VAL-FCPST"] = float(n.find('vFCPST').text)
                    pis = imp.find('.//PIS')
                    if pis is not None:
                        for p in pis:
                            if p.find('CST') is not None: linha["CST-PIS"] = p.find('CST').text.zfill(2)
                            if p.find('vBC') is not None: linha["BC-FED"] = float(p.find('vBC').text)
                            if p.find('vPIS') is not None: linha["VAL-PIS"] = float(p.find('vPIS').text)
                    cof = imp.find('.//COFINS')
                    if cof is not None:
                        for c in cof:
                            if c.find('CST') is not None: linha["CST-COF"] = c.find('CST').text.zfill(2)
                            if c.find('vCOFINS') is not None: linha["VAL-COF"] = float(c.find('vCOFINS').text)
                    ipi_n = imp.find('.//IPI')
                    if ipi_n is not None:
                        cst_i = ipi_n.find('.//CST')
                        if cst_i is not None: linha["CST-IPI"] = cst_i.text.zfill(2)
                        if ipi_n.find('.//vBC') is not None: linha["BC-IPI"] = float(ipi_n.find('.//vBC').text)
                        if ipi_n.find('.//pIPI') is not None: linha["ALQ-IPI"] = float(ipi_n.find('.//pIPI').text)
                        if ipi_n.find('.//vIPI') is not None: linha["VAL-IPI"] = float(ipi_n.find('.//vIPI').text)
                    dif_n = imp.find('.//ICMSUFDest')
                    if dif_n is not None:
                        if dif_n.find('vICMSUFDest') is not None: linha["VAL-DIFAL"] = float(dif_n.find('vICMSUFDest').text)
                        if dif_n.find('vFCPUFDest') is not None: linha["VAL-FCP"] += float(dif_n.find('vFCPUFDest').text)
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_sai, ae_f, as_f, ge_f, gs_f, cod_cliente=""):
    def limpar_txt(v): return str(v).replace('.0', '').strip()
    def format_brl(v): return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    base_file = buscar_base_no_github(cod_cliente); lista_erros = []
    
    try:
        base_icms = pd.read_excel(base_file, sheet_name='ICMS')
        base_icms['NCM_KEY'] = base_icms.iloc[:, 0].apply(limpar_txt).str.replace(r'\D', '', regex=True).str.zfill(8)
        base_icms['CST_KEY'] = base_icms.iloc[:, 1].apply(limpar_txt).str.zfill(2) # Ajustado p/ coluna 1 (CST Interna)
        base_pc = pd.read_excel(base_file, sheet_name='PIS_COFINS')
        base_pc['NCM_KEY'] = base_pc.iloc[:, 0].apply(limpar_txt).str.replace(r'\D', '', regex=True).str.zfill(8)
        base_pc.columns = [c.upper() for c in base_pc.columns]
    except: base_icms = pd.DataFrame(); base_pc = pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # MANUAL COMPLETO (Aba 1)
        pd.DataFrame({
            "RETORNO": ["✅ Correto", "❌ Divergente", "❌ NCM Ausente", "Situação Nota", "Complemento"],
            "EXPLICAÇÃO": [
                "XML e Base conferem.", "Diferença encontrada entre XML e Base.", "NCM não localizado na Base da empresa.",
                "Status real SEFAZ (Autorizado/Cancelado) cruzado com Autenticidade.", "Cálculo do imposto faltante ou a maior com base na divergência."
            ]
        }).to_excel(writer, sheet_name='MANUAL', index=False)

        # Cruzamento Autenticidade
        def cruzar_status(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col_s = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col_s]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_ent = cruzar_status(df_ent, ae_f); df_sai = cruzar_status(df_sai, as_f)

        if not df_sai.empty:
            # --- ABA ICMS (TAGS + AUDITORIA RESTAURADA) ---
            df_i = df_sai.copy(); tem_e = not df_ent.empty
            ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if tem_e else []
            def audit_icms(row):
                ncm = str(row['NCM']).zfill(8); info = base_icms[base_icms['NCM_KEY'] == ncm] if not base_icms.empty else pd.DataFrame()
                st_e = "✅ ST Localizado" if ncm in ncm_st else "❌ Sem ST na Entrada" if tem_e else "⚠️ Sem Entrada"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Erro": "ICMS: NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", format_brl(row['VPROD']), "R$ 0,00", "Cadastrar NCM", "R$ 0,00"])
                cst_e, aliq_e = str(info.iloc[0]['CST_KEY']), float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag, acao = [], []
                if str(row['CST-ICMS']).zfill(2) != cst_e.zfill(2): diag.append("CST: Divergente"); acao.append(f"Cc-e (CST {cst_e})")
                if abs(row['ALQ-ICMS'] - aliq_e) > 0.01: diag.append("Aliq: Divergente"); acao.append("Ajustar Alíquota")
                if diag: lista_erros.append({"NF": row['NUM_NF'], "Erro": f"ICMS: {'; '.join(diag)}"})
                return pd.Series([sit, st_e, "; ".join(diag) if diag else "✅ Correto", format_brl(row['VPROD']), format_brl(row['BC-ICMS']*aliq_e/100), " + ".join(acao) if acao else "✅ Correto", format_brl(max(0, (aliq_e-row['ALQ-ICMS'])*row['BC-ICMS']/100))])
            
            df_i[['Situação Nota', 'ST na Entrada', 'Diagnóstico', 'Valor', 'ICMS Esperado', 'Ação', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS', index=False)

            # --- ABA PIS/COFINS (TAGS + AUDITORIA RESTAURADA) ---
            df_pc = df_sai.copy()
            def audit_pc(row):
                ncm = str(row['NCM']).zfill(8); info = base_pc[base_pc['NCM_KEY'] == ncm] if not base_pc.empty else pd.DataFrame()
                if info.empty: return pd.Series(["❌ NCM Ausente", f"P/C: {row['CST-PIS']}/{row['CST-COF']}", "-", "Cadastrar"])
                try: cp_e, cc_e = str(info.iloc[0]['CST ENTRADA']).zfill(2), str(info.iloc[0]['CST SAÍDA']).zfill(2)
                except: cp_e, cc_e = "01", "01"
                diag, acao = [], []
                if str(row['CST-PIS']) != cp_e: diag.append("PIS: Divergente"); acao.append(f"CST PIS {cp_e}")
                if diag: lista_erros.append({"NF": row['NUM_NF'], "Erro": "P/C: CST Divergente"})
                return pd.Series(["; ".join(diag) if diag else "✅ Correto", f"P/C: {row['CST-PIS']}", f"P/C: {cp_e}", " + ".join(acao) if acao else "OK"])
            df_pc[['Diagnóstico', 'CST XML', 'CST Esperado', 'Ação']] = df_pc.apply(audit_pc, axis=1)
            df_pc.to_excel(writer, sheet_name='PIS_COFINS', index=False)

            # --- DEMAIS ABAS (IPI, DIFAL, DESTINO) ---
            df_sai.to_excel(writer, sheet_name='DIFAL', index=False)
            df_sai.to_excel(writer, sheet_name='IPI', index=False)

        # RESUMO DE ERROS (Checklist dinâmico)
        pd.DataFrame(lista_erros if lista_erros else [{"NF": "-", "Erro": "Nenhuma inconsistência."}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
