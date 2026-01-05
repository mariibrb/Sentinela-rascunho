import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
    """Converte valores do XML e Excel de forma robusta (12 ou 12,0 ou 12.0 = 12.0)"""
    if v is None: return 0.0
    try:
        txt = str(v).replace('R$', '').replace(' ', '').strip()
        # Se contiver vírgula e ponto (ex: 1.200,00), limpa o ponto e troca a vírgula
        if ',' in txt and '.' in txt: txt = txt.replace('.', '').replace(',', '.')
        # Se contiver apenas vírgula (ex: 12,0), troca por ponto
        elif ',' in txt: txt = txt.replace(',', '.')
        res = float(txt)
        return res if res < 1000000 else 0.0 # Blindagem contra erros de leitura absurdos
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
                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF'),
                    "DATA_EMISSAO": pd.to_datetime(buscar('dhEmi')).replace(tzinfo=None) if buscar('dhEmi') else None,
                    "UF_EMIT": buscar('UF', root.find('.//emit')), "UF_DEST": buscar('UF', root.find('.//dest')),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "COD_PROD": buscar('cProd', prod), "DESCR": buscar('xProd', prod),
                    "VPROD": safe_float(buscar('vProd', prod)), "ORIGEM": "", "CST-ICMS": "", "BC-ICMS": 0.0, 
                    "VLR-ICMS": 0.0, "ALQ-ICMS": 0.0, "ICMS-ST": 0.0, "CST-PIS": "", "VAL-PIS": 0.0, 
                    "CST-IPI": "", "BC-IPI": 0.0, "ALQ-IPI": 0.0, "VLR-IPI": 0.0, "VAL-DIFAL": 0.0
                }
                if imp is not None:
                    # CORREÇÃO CRÍTICA DO CST E ORIGEM (Busca Profunda)
                    icms_node = imp.find('.//ICMS')
                    if icms_node is not None:
                        for sub_tag in icms_node: # Entra em ICMS00, ICMS10, etc.
                            o = sub_tag.find('orig'); c = sub_tag.find('CST') or sub_tag.find('CSOSN')
                            if o is not None: linha["ORIGEM"] = o.text
                            if c is not None: linha["CST-ICMS"] = c.text.zfill(2)
                            linha["BC-ICMS"] = safe_float(buscar('vBC', sub_tag))
                            linha["VLR-ICMS"] = safe_float(buscar('vICMS', sub_tag))
                            linha["ALQ-ICMS"] = safe_float(buscar('pICMS', sub_tag))
                            linha["ICMS-ST"] = safe_float(buscar('vICMSST', sub_tag))
                    
                    pis_node = imp.find('.//PIS')
                    if pis_node is not None:
                        for p in pis_node:
                            c_p = p.find('CST'); v_p = p.find('vPIS')
                            if c_p is not None: linha["CST-PIS"] = c_p.text.zfill(2)
                            if v_p is not None: linha["VAL-PIS"] = safe_float(v_p.text)
                    
                    ipi_node = imp.find('.//IPI')
                    if ipi_node is not None:
                        c_i = ipi_node.find('.//CST')
                        if c_i is not None: linha["CST-IPI"] = c_i.text.zfill(2)
                        linha["ALQ-IPI"] = safe_float(buscar('pIPI', ipi_node))
                        linha["VLR-IPI"] = safe_float(buscar('vIPI', ipi_node))
                        linha["BC-IPI"] = safe_float(buscar('vBC', ipi_node))

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
        # --- MANUAL DE INSTRUÇÕES (TEXTO CORRIDO) ---
        manual_txt = (
            "SENTINELA - MANUAL DE OPERAÇÃO E AUDITORIA FISCAL\n\n"
            "1. OBJETIVO DO RELATÓRIO:\n"
            "Este arquivo é o resultado do cruzamento automatizado entre os XMLs de sua empresa e as diretrizes tributárias cadastradas na Base GitHub.\n\n"
            "2. ESTRUTURA DAS ABAS:\n"
            "- ICMS_AUDIT: Confronta Alíquotas e CSTs de ICMS, além de validar a retenção de ST via cruzamento com notas de entrada.\n"
            "- PIS_COFINS_AUDIT: Valida se o CST de Saída informado no XML corresponde à regra de crédito/débito da base.\n"
            "- IPI_AUDIT: Audita Alíquotas e CST de IPI destacadas nos itens industriais.\n"
            "- DIFAL_AUDIT: Identifica notas interestaduais que deveriam conter o recolhimento do Diferencial de Alíquotas.\n"
            "- RESUMO_ERROS: Uma lista consolidada de todas as notas que apresentaram qualquer divergência.\n\n"
            "3. LEGENDA DE DIAGNÓSTICOS:\n"
            "- ✅ Correto: Informação no XML está em total conformidade com a Base Tributária.\n"
            "- ❌ Divergente: Foi encontrada uma diferença de alíquota ou código de tributação.\n"
            "- ❌ NCM Ausente: O produto não existe na sua base de dados (requer cadastro imediato).\n"
            "- Status Nota: Indica se a nota está Autorizada ou Cancelada conforme o seu relatório de Autenticidade."
        )
        pd.DataFrame({"SENTINELA": [manual_txt]}).to_excel(writer, sheet_name='MANUAL', index=False)
        writer.sheets['MANUAL'].set_column('A:A', 130)

        # Cruzamento Autenticidade
        def cruzar_aut(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col_s = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col_s]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_sai = cruzar_aut(df_sai, as_f); df_ent = cruzar_aut(df_ent, ae_f)

        if not df_sai.empty:
            # ICMS AUDIT (Tags + Inteligência)
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]; st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": "NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", format_brl(row['VPROD']), "Cadastrar", "R$ 0,00"])
                aliq_e = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                cst_e = str(info.iloc[0, 1]).zfill(2)
                diag = "✅ Correto" if (abs(row['ALQ-ICMS'] - aliq_e) < 0.01 and row['CST-ICMS'] == cst_e) else "❌ Divergente"
                comp = max(0, (aliq_e - row['ALQ-ICMS']) * row['BC-ICMS'] / 100)
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "ICMS", "Erro": f"Aliq {row['ALQ-ICMS']}% (Esperado {aliq_e}%)"})
                return pd.Series([sit, st_e, diag, format_brl(row['VPROD']), "Ajustar" if diag != "✅ Correto" else "OK", format_brl(comp)])
            df_i[['Situação Nota', 'Check ST Entrada', 'Diagnóstico ICMS', 'Valor Item', 'Ação', 'Complemento ICMS']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # PIS/COFINS AUDIT (Tags + Inteligência)
            df_pc_a = df_sai.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "Cadastrar"])
                cc_e = str(info.iloc[0, 2]).zfill(2) if len(info.columns) > 2 else "01"
                diag = "✅ Correto" if row['CST-PIS'] == cc_e else "❌ Divergente"
                if diag != "✅ Correto": lista_erros.append({"NF": row['NUM_NF'], "Tipo": "PIS/COF", "Erro": f"CST {row['CST-PIS']} (Esperado {cc_e})"})
                return pd.Series([diag, f"Esperado {cc_e}" if diag != "✅ Correto" else "OK"])
            df_pc_a[['Diagnóstico PIS/COFINS', 'Ação PIS/COFINS']] = df_pc_a.apply(audit_pc, axis=1)
            df_pc_a.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # IPI AUDIT (Tags + Inteligência)
            df_ipi_a = df_sai.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return pd.Series(["❌ NCM Ausente", "OK"])
                ci_e, ai_e = str(info.iloc[0, 1]).zfill(2), safe_float(info.iloc[0, 2])
                diag = "✅ Correto" if (str(row['CST-IPI']) == ci_e and abs(row['ALQ-IPI'] - ai_e) < 0.01) else "❌ Divergente"
                return pd.Series([diag, f"Esperado {ci_e}/{ai_e}%" if diag != "✅ Correto" else "OK"])
            df_ipi_a[['Diagnóstico IPI', 'Ação IPI']] = df_ipi_a.apply(audit_ipi, axis=1)
            df_ipi_a.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # DIFAL AUDIT
            df_dif = df_sai.copy()
            def audit_dif(row):
                is_i = row['UF_EMIT'] != row['UF_DEST']; cfop = str(row['CFOP'])
                if is_i and cfop in ['6107', '6108', '6933', '6404'] and row['VAL-DIFAL'] == 0:
                    return pd.Series(["❌ DIFAL Obrigatório", "Complementar"])
                return pd.Series(["✅ OK", "OK"])
            df_dif[['Audit DIFAL', 'Ação DIFAL']] = df_dif.apply(audit_dif, axis=1)
            df_dif.to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

        # RESUMO FINAL DE ERROS (Checklist)
        df_res = pd.DataFrame(lista_erros) if lista_erros else pd.DataFrame({"NF": ["-"], "Erro": ["Tudo Correto"]})
        df_res.to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
