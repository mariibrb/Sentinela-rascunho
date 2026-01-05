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
            chave = inf_nfe.attrib.get('Id', '')[3:] if inf_nfe is not None else ""
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'),
                    "DATA_EMISSAO": pd.to_datetime(buscar('dhEmi')).replace(tzinfo=None) if buscar('dhEmi') else None,
                    "UF_EMIT": buscar('UF', root.find('.//emit')), "UF_DEST": buscar('UF', root.find('.//dest')),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "VPROD": float(buscar('vProd', prod) or 0.0), "BC-ICMS": 0.0, "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0, 
                    "CST-ICMS": "", "ORIGEM": "", "ICMS-ST": 0.0, "CST-PIS": "", "CST-COF": "", 
                    "CST-IPI": "", "BC-IPI": 0.0, "ALQ-IPI": 0.0, "VLR-IPI": 0.0, "VAL-DIFAL": 0.0
                }
                if imp is not None:
                    icms = imp.find('.//ICMS'); ipi = imp.find('.//IPI'); pis = imp.find('.//PIS'); dif = imp.find('.//ICMSUFDest')
                    if icms is not None:
                        for n in icms:
                            orig = n.find('orig'); cst = n.find('CST') or n.find('CSOSN')
                            if orig is not None: linha["ORIGEM"] = orig.text
                            if cst is not None: linha["CST-ICMS"] = cst.text.zfill(2)
                            if n.find('pICMS') is not None: linha["ALQ-ICMS"] = float(n.find('pICMS').text)
                            if n.find('vICMS') is not None: linha["VLR-ICMS"] = float(n.find('vICMS').text)
                            if n.find('vBC') is not None: linha["BC-ICMS"] = float(n.find('vBC').text)
                            if n.find('vICMSST') is not None: linha["ICMS-ST"] = float(n.find('vICMSST').text)
                    if ipi is not None:
                        cst_i = ipi.find('.//CST')
                        if cst_i is not None: linha["CST-IPI"] = cst_i.text.zfill(2)
                        if ipi.find('.//pIPI') is not None: linha["ALQ-IPI"] = float(ipi.find('.//pIPI').text)
                        if ipi.find('.//vIPI') is not None: linha["VLR-IPI"] = float(ipi.find('.//vIPI').text)
                        if ipi.find('.//vBC') is not None: linha["BC-IPI"] = float(ipi.find('.//vBC').text)
                    if pis is not None:
                        for p in pis:
                            if p.find('CST') is not None: linha["CST-PIS"] = p.find('CST').text.zfill(2)
                    if dif is not None and dif.find('vICMSUFDest') is not None:
                        linha["VAL-DIFAL"] = float(dif.find('vICMSUFDest').text)
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
    except: base_icms = pd.DataFrame(); base_pc = pd.DataFrame(); base_ipi = pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # --- MANUAL DE INSTRUÇÕES COMPLETO ---
        df_manual = pd.DataFrame({
            "COLUNA / RETORNO": [
                "Situação Nota", "ST na Entrada", "Diagnóstico ICMS", "Complemento", 
                "✅ Correto", "❌ Divergente", "❌ NCM Ausente", "⚠️ N/Verif"
            ],
            "DESCRIÇÃO DETALHADA": [
                "Status da nota cruzado com o relatório de Autenticidade (Autorizado/Cancelado).",
                "Check se houve nota de entrada com retenção de ST para o NCM (CST 60 ou vICMSST > 0).",
                "Confronto da Alíquota/CST do XML com as regras cadastradas na sua Base Tributária.",
                "Cálculo matemático do imposto devido (Esperado - Destacado) em caso de divergência.",
                "Indica que as informações tributárias do XML batem 100% com a sua Base.",
                "Indica que existe uma diferença de alíquota ou CST entre a nota e a Base.",
                "O código NCM da nota fiscal não foi localizado na sua Base Tributária.",
                "Não foi possível cruzar a nota com o arquivo de Autenticidade enviado."
            ]
        })
        df_manual.to_excel(writer, sheet_name='MANUAL', index=False)

        # Cruzamento Autenticidade
        def cruzar(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col_st = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col_st]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_sai = cruzar(df_sai, as_f)

        if not df_sai.empty:
            # ANÁLISE ICMS (TAGS + AUDITORIA)
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]; st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Erro": "ICMS: NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", format_brl(row['VPROD']), "Cadastrar NCM", "R$ 0,00"])
                aliq_e = float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - aliq_e) < 0.01 else "❌ Divergente"
                comp = max(0, (aliq_e - row['ALQ-ICMS']) * row['BC-ICMS'] / 100)
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Erro": f"ICMS Aliq {row['ALQ-ICMS']}% (Esperado {aliq_e}%)"})
                return pd.Series([sit, st_e, diag, format_brl(row['VPROD']), "Ajustar" if diag != "✅ Correto" else "OK", format_brl(comp)])
            
            df_i[['Situação Nota', 'Check ST Entrada', 'Diagnóstico ICMS', 'Valor Item', 'Ação', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # ANÁLISE PIS/COFINS
            df_pc_a = df_sai.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                cc_e = str(info.iloc[0, 2]).zfill(2)
                diag = "✅ Correto" if str(row['CST-PIS']) == cc_e else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Erro": f"P/C: CST {row['CST-PIS']} (Esperado {cc_e})"})
                return pd.Series([diag, f"Esperado: {cc_e}" if diag != "✅ Correto" else "OK"])
            df_pc_a[['Diagnóstico PIS/COFINS', 'Ação']] = df_pc_a.apply(audit_pc, axis=1)
            df_pc_a.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # ANÁLISE IPI
            df_ipi_a = df_sai.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                ci_e, ai_e = str(info.iloc[0, 1]).zfill(2), float(info.iloc[0, 2])
                diag = "✅ Correto" if (str(row['CST-IPI']) == ci_e and abs(row['ALQ-IPI'] - ai_e) < 0.01) else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Erro": "IPI Divergente"})
                return pd.Series([diag, f"Esperado {ci_e} / {ai_e}%" if diag != "✅ Correto" else "OK"])
            df_ipi_a[['Diagnóstico IPI', 'Ação Sugerida']] = df_ipi_a.apply(audit_ipi, axis=1)
            df_ipi_a.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

        # RESUMO DE ERROS (Checklist Final)
        df_res = pd.DataFrame(lista_erros) if lista_erros else pd.DataFrame({"NF": ["-"], "Erro": ["Tudo Correto"]})
        df_res.to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
