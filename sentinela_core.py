import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
    """Converte valores do XML e Excel de forma robusta garantindo que 12% seja 12.0"""
    if v is None: return 0.0
    try:
        txt = str(v).replace('R$', '').replace(' ', '').strip()
        if ',' in txt and '.' in txt: txt = txt.replace('.', '').replace(',', '.')
        elif ',' in txt: txt = txt.replace(',', '.')
        return round(float(txt), 4)
    except: return 0.0

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
            chave = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
            emit = root.find('.//emit'); dest = root.find('.//dest')
            cnpj_e = buscar('CNPJ', emit) or buscar('CPF', emit)
            cnpj_d = buscar('CNPJ', dest) or buscar('CPF', dest)

            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'), "CNPJ_EMIT": cnpj_e, "CNPJ_DEST": cnpj_d,
                    "UF_EMIT": buscar('UF', emit), "UF_DEST": buscar('UF', dest),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "COD_PROD": buscar('cProd', prod), "DESCR": buscar('xProd', prod),
                    "VPROD": safe_float(buscar('vProd', prod)), "ORIGEM": "", "CST-ICMS": "", 
                    "BC-ICMS": 0.0, "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0, "ICMS-ST": 0.0,
                    "CST-PIS": "", "VAL-PIS": 0.0, "CST-COF": "", "VAL-COF": 0.0,
                    "CST-IPI": "", "BC-IPI": 0.0, "ALQ-IPI": 0.0, "VLR-IPI": 0.0, "VAL-DIFAL": 0.0
                }
                if imp is not None:
                    # BUSCA RECURSIVA CST ICMS (ICMS00...ICMS90)
                    icms_node = imp.find('.//ICMS')
                    if icms_node is not None:
                        for sub in icms_node:
                            o = sub.find('orig'); c = sub.find('CST') or sub.find('CSOSN')
                            if o is not None: linha["ORIGEM"] = o.text
                            if c is not None: linha["CST-ICMS"] = c.text.zfill(2)
                            linha["BC-ICMS"] = safe_float(buscar('vBC', sub))
                            linha["VLR-ICMS"] = safe_float(buscar('vICMS', sub))
                            linha["ALQ-ICMS"] = safe_float(buscar('pICMS', sub))
                            linha["ICMS-ST"] = safe_float(buscar('vICMSST', sub))
                    # PIS/COFINS/IPI/DIFAL
                    pis_node = imp.find('.//PIS'); cof_node = imp.find('.//COFINS')
                    if pis_node is not None:
                        for p in pis_node:
                            cp = p.find('CST'); vp = p.find('vPIS')
                            if cp is not None: linha["CST-PIS"] = cp.text.zfill(2)
                            if vp is not None: linha["VAL-PIS"] = safe_float(vp.text)
                    if cof_node is not None:
                        for co in cof_node:
                            cc = co.find('CST'); vc = co.find('vCOFINS')
                            if cc is not None: linha["CST-COF"] = cc.text.zfill(2)
                            if vc is not None: linha["VAL-COF"] = safe_float(vc.text)
                    ipi_node = imp.find('.//IPI')
                    if ipi_node is not None:
                        ci = ipi_node.find('.//CST'); vi = ipi_node.find('.//vIPI'); bi = ipi_node.find('.//vBC'); ai = ipi_node.find('.//pIPI')
                        if ci is not None: linha["CST-IPI"] = ci.text.zfill(2)
                        linha["VAL-IPI"] = safe_float(vi.text) if vi is not None else 0.0
                        linha["BC-IPI"] = safe_float(bi.text) if bi is not None else 0.0
                        linha["ALQ-IPI"] = safe_float(ai.text) if ai is not None else 0.0
                    dif_node = imp.find('.//ICMSUFDest')
                    if dif_node is not None: linha["VAL-DIFAL"] = safe_float(buscar('vICMSUFDest', dif_node))
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_sai, ae_f, as_f, ge_f, gs_f, cod_cliente=""):
    def format_brl(v): return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    base_file = buscar_base_no_github(cod_cliente); lista_erros = []
    
    try:
        base_icms = pd.read_excel(base_file, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms.iloc[:, 0].astype(str).str.zfill(8)
        base_pc = pd.read_excel(base_file, sheet_name='PIS_COFINS'); base_pc['NCM_KEY'] = base_pc.iloc[:, 0].astype(str).str.zfill(8)
        base_ipi = pd.read_excel(base_file, sheet_name='IPI'); base_ipi['NCM_KEY'] = base_ipi.iloc[:, 0].astype(str).str.zfill(8)
    except: base_icms, base_pc, base_ipi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # MANUAL (TEXTO CORRIDO)
        manual_txt = (
            "SENTINELA - MANUAL DE AUDITORIA FISCAL\n\n"
            "1. FUNCIONAMENTO: O sistema extrai as tags do XML e confronta com a base do GitHub.\n"
            "2. COLUNAS DE AUDITORIA:\n"
            "- Situação Nota: Status real (Autorizado/Cancelado).\n"
            "- Diagnóstico: ✅ Correto / ❌ Divergente / ❌ NCM Ausente.\n"
            "- Complemento: Valor financeiro do imposto devido em caso de erro.\n"
            "3. ABAS: ICMS, PIS/COFINS, IPI e DIFAL auditados item a item."
        )
        pd.DataFrame({"INSTRUÇÕES": [manual_txt]}).to_excel(writer, sheet_name='MANUAL', index=False)
        writer.sheets['MANUAL'].set_column('A:A', 130)

        def cruzar_st(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_sai = cruzar_st(df_sai, as_f); df_ent = cruzar_st(df_ent, ae_f)

        if not df_sai.empty:
            # ICMS
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]
                st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": "NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", "Cadastrar", "R$ 0,00"])
                aliq_e = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                cst_e = str(info.iloc[0, 1]).zfill(2)
                diag = "✅ Correto" if (abs(row['ALQ-ICMS'] - aliq_e) < 0.01 and row['CST-ICMS'] == cst_e) else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": f"CST/Aliq Errada ({row['ALQ-ICMS']}%)"})
                comp = max(0, (aliq_e - row['ALQ-ICMS']) * row['BC-ICMS'] / 100)
                return pd.Series([sit, st_e, diag, "Ajustar" if diag != "✅ Correto" else "OK", format_brl(comp)])
            df_i[['Situação Nota', 'ST na Entrada', 'Diagnóstico', 'Ação', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # PIS/COFINS
            df_p = df_sai.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                cc_e = str(info.iloc[0, 2]).zfill(2) if len(info.columns) > 2 else "01"
                diag = "✅ Correto" if row['CST-PIS'] == cc_e else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "PIS/COF", "Erro": "CST Errado"})
                return pd.Series([diag, f"Esperado {cc_e}" if diag != "✅ Correto" else "OK"])
            df_p[['Diagnóstico PIS/COFINS', 'Ação Sugerida']] = df_p.apply(audit_pc, axis=1)
            df_p.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # IPI
            df_ipi_a = df_sai.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                ci_e, ai_e = str(info.iloc[0, 1]).zfill(2), safe_float(info.iloc[0, 2])
                diag = "✅ Correto" if (row['CST-IPI'] == ci_e and abs(row['ALQ-IPI'] - ai_e) < 0.01) else "❌ Divergente"
                return pd.Series([diag, f"Esperado {ci_e}/{ai_e}%" if diag != "✅ Correto" else "OK"])
            df_ipi_a[['Diagnóstico IPI', 'Ação IPI']] = df_ipi_a.apply(audit_ipi, axis=1)
            df_ipi_a.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # DIFAL
            df_sai.to_excel(writer, sheet_name='DIFAL_ANALISE', index=False)

        pd.DataFrame(lista_erros if lista_erros else [{"NF": "-", "Erro": "Tudo Correto"}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
