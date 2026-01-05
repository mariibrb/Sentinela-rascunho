import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
    if v is None: return 0.0
    try:
        txt = str(v).replace('R$', '').replace('.', '').replace(',', '.').strip()
        return float(txt) if any(c.isdigit() for c in txt) else 0.0
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
            
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                ncm_limpo = re.sub(r'\D', '', buscar('NCM', prod)).zfill(8)
                
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'),
                    "UF_EMIT": buscar('UF', root.find('.//emit')), "UF_DEST": buscar('UF', root.find('.//dest')),
                    "NCM": ncm_limpo, "VPROD": safe_float(buscar('vProd', prod)),
                    "ORIGEM": "", "CST-ICMS": "", "BC-ICMS": 0.0, "ALQ-ICMS": 0.0, "VLR-ICMS": 0.0, "ICMS-ST": 0.0,
                    "CST-IPI": "", "BC-IPI": 0.0, "ALQ-IPI": 0.0, "VLR-IPI": 0.0,
                    "CST-PIS": "", "CST-COF": "", "VAL-DIFAL": 0.0
                }
                
                if imp is not None:
                    # CORREÇÃO DA LEITURA DE CST (Hierarquia Blindada)
                    icms_tags = imp.find('.//ICMS')
                    if icms_tags is not None:
                        for subtag in icms_tags:
                            orig_tag = subtag.find('orig')
                            cst_tag = subtag.find('CST') or subtag.find('CSOSN')
                            if orig_tag is not None: linha["ORIGEM"] = orig_tag.text
                            if cst_tag is not None: linha["CST-ICMS"] = cst_tag.text.zfill(2)
                            linha["ALQ-ICMS"] = safe_float(buscar('pICMS', subtag))
                            linha["VLR-ICMS"] = safe_float(buscar('vICMS', subtag))
                            linha["BC-ICMS"] = safe_float(buscar('vBC', subtag))
                            linha["ICMS-ST"] = safe_float(buscar('vICMSST', subtag))

                    ipi_tag = imp.find('.//IPI')
                    if ipi_tag is not None:
                        cst_i = ipi_tag.find('.//CST')
                        if cst_i is not None: linha["CST-IPI"] = cst_i.text.zfill(2)
                        linha["ALQ-IPI"] = safe_float(buscar('pIPI', ipi_tag))
                        linha["VLR-IPI"] = safe_float(buscar('vIPI', ipi_tag))
                        linha["BC-IPI"] = safe_float(buscar('vBC', ipi_tag))

                    pis_tag = imp.find('.//PIS')
                    if pis_tag is not None:
                        cst_p = pis_tag.find('.//CST')
                        if cst_p is not None: linha["CST-PIS"] = cst_p.text.zfill(2)

                    dif_tag = imp.find('.//ICMSUFDest')
                    if dif_tag is not None:
                        linha["VAL-DIFAL"] = safe_float(buscar('vICMSUFDest', dif_tag))
                
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
        # --- MANUAL DE INSTRUÇÕES "DE VERDADE" ---
        pd.DataFrame({
            "SEÇÃO": ["1. OBJETIVO", "2. COMO OPERAR", "3. LEGENDA DE DIAGNÓSTICOS", "4. COLUNAS DE INTELIGÊNCIA"],
            "DETALHAMENTO": [
                "Este relatório audita a conformidade fiscal entre os XMLs de saída e a sua Base Tributária no GitHub.",
                "Suba os XMLs, o Gerencial (CSV) e a Autenticidade (Excel) para processamento completo.",
                "✅ Correto: XML e Base iguais | ❌ Divergente: Diferença de Alíquota ou CST | ❌ NCM Ausente: Item sem cadastro.",
                "Status Nota: Cruzamento SEFAZ | ST na Entrada: Check de custo | Complemento: Valor financeiro do erro."
            ]
        }).to_excel(writer, sheet_name='MANUAL', index=False)

        def cruzar_aut(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col_st = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col_st]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_ent = cruzar_aut(df_ent, ae_f); df_sai = cruzar_aut(df_sai, as_f)

        if not df_sai.empty:
            # AUDITORIA ICMS (Tags + Análise)
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]; st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": "NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", "Cadastrar", "R$ 0,00"])
                aliq_e = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - aliq_e) < 0.01 else "❌ Divergente"
                comp = max(0, (aliq_e - row['ALQ-ICMS']) * row['BC-ICMS'] / 100)
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": f"Aliq {row['ALQ-ICMS']}% (Base: {aliq_e}%)"})
                return pd.Series([sit, st_e, diag, "Ajustar" if diag != "✅ Correto" else "OK", format_brl(comp)])
            
            df_i[['Situação Nota', 'ST Entrada', 'Diagnóstico ICMS', 'Ação', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # AUDITORIA PIS/COFINS (Tags + Análise)
            df_pc_a = df_sai.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "Cadastrar"])
                cc_e = str(info.iloc[0, 2]).zfill(2) # CST Saída
                diag = "✅ Correto" if str(row['CST-PIS']) == cc_e else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "PIS/COF", "Erro": f"CST {row['CST-PIS']} (Esperado: {cc_e})"})
                return pd.Series([diag, f"Correto: {cc_e}" if diag != "✅ Correto" else "OK"])
            df_pc_a[['Diagnóstico PIS/COFINS', 'Ação Sugerida']] = df_pc_a.apply(audit_pc, axis=1)
            df_pc_a.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # AUDITORIA IPI (Tags + Análise)
            df_ipi_a = df_sai.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                ci_e, ai_e = str(info.iloc[0, 1]).zfill(2), safe_float(info.iloc[0, 2])
                diag = "✅ Correto" if (str(row['CST-IPI']) == ci_e and abs(row['ALQ-IPI'] - ai_e) < 0.01) else "❌ Divergente"
                return pd.Series([diag, f"Esperado {ci_e} / {ai_e}%" if diag != "✅ Correto" else "OK"])
            df_ipi_a[['Diagnóstico IPI', 'Ação IPI']] = df_ipi_a.apply(audit_ipi, axis=1)
            df_ipi_a.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # DIFAL
            df_sai[['CHAVE_ACESSO', 'NUM_NF', 'VAL-DIFAL']].to_excel(writer, sheet_name='DIFAL_ANALISE', index=False)

        # RESUMO DE ERROS (Checklist Final)
        df_res = pd.DataFrame(lista_erros) if lista_erros else pd.DataFrame({"NF": ["-"], "Erro": ["Tudo Correto"]})
        df_res.to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
