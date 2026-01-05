import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
    """Garante que 12% seja lido como 12.0 independente de vírgula ou ponto."""
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
        res = requests.get(url, headers=headers, timeout=10)
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
            xml_data = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace'))
            root = ET.fromstring(xml_data)
            def buscar(tag, raiz):
                alvo = raiz.find(f'.//{tag}')
                return alvo.text if alvo is not None and alvo.text is not None else ""
            def buscar_recursivo(node, tags_alvo):
                for elem in node.iter():
                    tag_limpa = elem.tag.split('}')[-1]
                    if tag_limpa in tags_alvo: return elem.text
                return ""

            inf = root.find('.//infNFe'); chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
            emit = root.find('.//emit'); dest = root.find('.//dest')
            cnpj_e = buscar('CNPJ', emit) or buscar('CPF', emit)
            cnpj_d = buscar('CNPJ', dest) or buscar('CPF', dest)

            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                icms_node = imp.find('.//ICMS') if imp is not None else None
                
                # EXTRAÇÃO DAS TAGS DO XML
                cst_ex = buscar_recursivo(icms_node, ['CST', 'CSOSN']) if icms_node is not None else ""
                orig_ex = buscar_recursivo(icms_node, ['orig']) if icms_node is not None else ""
                
                linha = {
                    "CHAVE_ACESSO": str(chave).strip(),
                    "NUM_NF": buscar('nNF', root),
                    "CNPJ_EMIT": cnpj_e,
                    "CNPJ_DEST": cnpj_d,
                    "UF_EMIT": buscar('UF', emit),
                    "UF_DEST": buscar('UF', dest),
                    "CFOP": buscar('CFOP', prod),
                    "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "VPROD": safe_float(buscar('vProd', prod)),
                    "ORIGEM": orig_ex,
                    "CST-ICMS": cst_ex.zfill(2) if cst_ex else "",
                    "BC-ICMS": safe_float(buscar('vBC', imp)),
                    "ALQ-ICMS": safe_float(buscar('pICMS', imp)),
                    "VLR-ICMS": safe_float(buscar('vICMS', imp)),
                    "ICMS-ST": safe_float(buscar('vICMSST', imp)),
                    "CST-PIS": buscar_recursivo(imp.find('.//PIS'), ['CST']) if imp.find('.//PIS') is not None else "",
                    "VAL-PIS": safe_float(buscar('vPIS', imp)),
                    "CST-COF": buscar_recursivo(imp.find('.//COFINS'), ['CST']) if imp.find('.//COFINS') is not None else "",
                    "VAL-COF": safe_float(buscar('vCOFINS', imp)),
                    "CST-IPI": buscar_recursivo(imp.find('.//IPI'), ['CST']) if imp.find('.//IPI') is not None else "",
                    "BC-IPI": safe_float(buscar('vBC', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "ALQ-IPI": safe_float(buscar('pIPI', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "VAL-IPI": safe_float(buscar('vIPI', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "VAL-DIFAL": safe_float(buscar('vICMSUFDest', imp)),
                    "BC-DIFAL": safe_float(buscar('vBCUFDest', imp))
                }
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_sai, ae_f, as_f, ge_f, gs_f, cod_cliente=""):
    base_f = buscar_base_no_github(cod_cliente); lista_erros = []
    try:
        base_icms = pd.read_excel(base_f, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms.iloc[:, 0].astype(str).str.zfill(8)
        base_pc = pd.read_excel(base_f, sheet_name='PIS_COFINS'); base_pc['NCM_KEY'] = base_pc.iloc[:, 0].astype(str).str.zfill(8)
        base_ipi = pd.read_excel(base_f, sheet_name='IPI'); base_ipi['NCM_KEY'] = base_ipi.iloc[:, 0].astype(str).str.zfill(8)
    except: base_icms, base_pc, base_ipi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # --- MANUAL BÍBLIA MAXIMALISTA ---
        man_l = [
            ["SENTINELA - MANUAL TÉCNICO COMPLETO (BÍBLIA DA AUDITORIA)"],
            [""], ["1. CONCEITO: Auditoria maximalista de XML vs Base Tributária."],
            ["2. CST 60: Se a saída é 60, o sistema busca o NCM nos XMLs de entrada para provar a ST."],
            ["3. SITUAÇÃO: Cruzamento rigoroso com Chave de Acesso para detectar Canceladas."],
            ["4. ALÍQUOTAS: Erros de interpretação de escala foram eliminados (12 = 12.0)."],
            ["5. PIS/COFINS/IPI/DIFAL: Analisados item a item com espelho de tags XML."],
            ["6. COMPLEMENTO: Valor estimado do prejuízo fiscal por linha divergente."]
        ]
        pd.DataFrame(man_l).to_excel(writer, sheet_name='MANUAL', index=False, header=False)

        def cruzar_aut(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col_c = next((c for c in df_a.columns if 'Chave' in str(c)), None)
                col_s = next((c for c in df_a.columns if 'Status' in str(c) or 'Situação' in str(c)), None)
                if col_c and col_s:
                    status_map = df_a.set_index(df_a[col_c].astype(str).str.strip())[col_s].to_dict()
                    df['Situação Nota'] = df['CHAVE_ACESSO'].map(status_map).fillna('⚠️ N/Verificado')
                return df
            except: return df
        df_sai = cruzar_aut(df_sai, as_f)

        if not df_sai.empty:
            # --- ICMS ---
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60") | (df_ent['ICMS-ST'] > 0)]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]
                st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST na Entrada"
                sit = row.get('Situação Nota', '⚠️ N/Verif')
                if row['CST-ICMS'] == "60":
                    diag = "✅ ST Validado" if row['NCM'] in ncm_st else "❌ Irregular: Saída 60 s/ entrada ST"
                    return pd.Series([sit, st_e, diag, "R$ 0,00"])
                if info.empty: return pd.Series([sit, st_e, "❌ NCM Ausente na Base", "R$ 0,00"])
                alq_b = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - alq_b) < 0.01 else f"❌ Aliq XML {row['ALQ-ICMS']}% diverge da Base {alq_b}%"
                comp = max(0, (alq_b - row['ALQ-ICMS']) * row['BC-ICMS'] / 100)
                if "❌" in diag: lista_erros.append({"NF": row['NUM_NF'], "Erro": diag})
                return pd.Series([sit, st_e, diag, f"R$ {comp:,.2f}"])
            df_i[['Situação Nota', 'Check ST Entrada', 'Diagnóstico ICMS', 'Complemento ICMS']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # --- PIS/COFINS ---
            df_p = df_sai[['CHAVE_ACESSO', 'NUM_NF', 'CNPJ_EMIT', 'CNPJ_DEST', 'NCM', 'VPROD', 'CST-PIS', 'VAL-PIS', 'CST-COF', 'VAL-COF']].copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return "❌ NCM Ausente na Base PIS/COF"
                cst_b = str(info.iloc[0, 2]).zfill(2) if len(info.columns) > 2 else "01"
                return "✅ Correto" if row['CST-PIS'] == cst_b else f"❌ CST XML {row['CST-PIS']} diverge da Base {cst_b}"
            df_p['Diagnóstico PIS/COF'] = df_p.apply(audit_pc, axis=1)
            df_p.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # --- IPI ---
            df_ip = df_sai[['CHAVE_ACESSO', 'NUM_NF', 'CNPJ_EMIT', 'CNPJ_DEST', 'NCM', 'VPROD', 'CST-IPI', 'BC-IPI', 'ALQ-IPI', 'VAL-IPI']].copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return "❌ NCM Ausente na Base IPI"
                alq_b = safe_float(info.iloc[0, 2])
                return "✅ Correto" if abs(row['ALQ-IPI'] - alq_b) < 0.01 else f"❌ Aliq IPI XML {row['ALQ-IPI']}% diverge da Base {alq_b}%"
            df_ip['Diagnóstico IPI'] = df_ip.apply(audit_ipi, axis=1)
            df_ip.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # --- DIFAL ---
            df_d = df_sai[['CHAVE_ACESSO', 'NUM_NF', 'CNPJ_EMIT', 'CNPJ_DEST', 'UF_EMIT', 'UF_DEST', 'CFOP', 'VAL-DIFAL', 'BC-DIFAL']].copy()
            df_d['Diagnóstico DIFAL'] = df_d.apply(lambda r: "❌ Alerta: Operação Interestadual sem DIFAL" if r['UF_EMIT'] != r['UF_DEST'] and r['VAL-DIFAL'] == 0 else "✅ OK", axis=1)
            df_d.to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

        pd.DataFrame(lista_erros if lista_erros else [{"NF": "-", "Erro": "Nenhuma inconsistência encontrada."}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
