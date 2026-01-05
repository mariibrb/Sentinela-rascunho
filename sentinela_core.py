import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
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
            # Remove namespaces para facilitar a busca de tags
            xml_data = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', f.read().decode('utf-8', errors='replace'))
            root = ET.fromstring(xml_data)
            
            def buscar(tag, raiz):
                alvo = raiz.find(f'.//{tag}')
                return alvo.text if alvo is not None and alvo.text is not None else ""

            def buscar_recursivo(node, tags_alvo):
                """Busca qualquer tag da lista dentro de um nó, independente da profundidade"""
                for elem in node.iter():
                    # Verifica se o nome da tag termina com algum dos alvos (ex: p:CST vira CST)
                    tag_limpa = elem.tag.split('}')[-1]
                    if tag_limpa in tags_alvo:
                        return elem.text
                return ""

            inf = root.find('.//infNFe')
            chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
            emit = root.find('.//emit'); dest = root.find('.//dest')
            cnpj_e = buscar('CNPJ', emit) or buscar('CPF', emit)
            cnpj_d = buscar('CNPJ', dest) or buscar('CPF', dest)

            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                
                # BUSCA BLINDADA DE CST / CSOSN / ORIGEM
                icms_node = imp.find('.//ICMS') if imp is not None else None
                cst_extraido = ""
                orig_extraida = ""
                if icms_node is not None:
                    cst_extraido = buscar_recursivo(icms_node, ['CST', 'CSOSN'])
                    orig_extraida = buscar_recursivo(icms_node, ['orig'])
                
                if cst_extraido: cst_extraido = cst_extraido.zfill(2)

                # BUSCA PIS/COFINS/IPI
                cst_p = ""; cst_i = ""; alq_i = 0.0
                pis_node = imp.find('.//PIS') if imp is not None else None
                if pis_node is not None: cst_p = buscar_recursivo(pis_node, ['CST'])
                
                ipi_node = imp.find('.//IPI') if imp is not None else None
                if ipi_node is not None: 
                    cst_i = buscar_recursivo(ipi_node, ['CST'])
                    alq_i = safe_float(buscar('pIPI', ipi_node))

                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF', root),
                    "CNPJ_EMIT": cnpj_e, "CNPJ_DEST": cnpj_d,
                    "UF_EMIT": buscar('UF', emit), "UF_DEST": buscar('UF', dest),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "VPROD": safe_float(buscar('vProd', prod)),
                    "ORIGEM": orig_extraida, "CST-ICMS": cst_extraido, 
                    "ALQ-ICMS": safe_float(buscar('pICMS', imp)),
                    "CST-PIS": cst_p.zfill(2) if cst_p else "",
                    "CST-IPI": cst_i.zfill(2) if cst_i else "",
                    "ALQ-IPI": alq_i,
                    "VAL-DIFAL": safe_float(buscar('vICMSUFDest', imp))
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
        # MANUAL COMPLETO LINHA POR LINHA
        man_l = [
            ["SENTINELA - MANUAL COMPLETO DE OPERAÇÃO E AUDITORIA"],
            [""],
            ["1. EXTRAÇÃO DE CST (SOLUÇÃO DEFINITIVA)"],
            ["O motor agora utiliza busca recursiva profunda. Ele ignora namespaces e caminhos fixos,"],
            ["varrendo todo o nó de imposto em busca das tags 'CST' ou 'CSOSN'."],
            [""],
            ["2. IDENTIFICAÇÃO"],
            ["CNPJ e CPF de Emitentes e Destinatários foram incluídos em todas as abas para rastreio."],
            [""],
            ["3. ABAS DE AUDITORIA"],
            ["- ICMS_AUDIT: Confronto de Alíquotas e histórico de ST na entrada."],
            ["- PIS_COFINS_AUDIT: Verificação de CST de PIS contra a Base GitHub."],
            ["- IPI_AUDIT: Verificação de Alíquotas de IPI."],
            ["- DIFAL_AUDIT: Extração de valores de DIFAL para conferência."],
            [""],
            ["4. STATUS DA NOTA"],
            ["A coluna 'Situação Nota' indica se a nota foi Autorizada ou Cancelada conforme"],
            ["o relatório de Autenticidade subido no Passo 2."],
            [""],
            ["5. RESUMO DE ERROS"],
            ["Lista consolidada de todas as divergências para ação imediata."]
        ]
        pd.DataFrame(man_l).to_excel(writer, sheet_name='MANUAL', index=False, header=False)
        writer.sheets['MANUAL'].set_column('A:A', 110)

        def cruzar(df, f):
            if df.empty or not f: return df
            try:
                df_a = pd.read_excel(f); col = 'Status' if 'Status' in df_a.columns else 'Situação'
                return pd.merge(df, df_a[['Chave NF-e', col]], left_on='CHAVE_ACESSO', right_on='Chave NF-e', how='left')
            except: return df
        df_sai = cruzar(df_sai, as_f); df_ent = cruzar(df_ent, ae_f)

        if not df_sai.empty:
            # ICMS AUDIT
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60")]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]
                st_e = "✅ ST Localizado" if row['NCM'] in ncm_st else "❌ Sem ST"
                sit = row.get('Status', row.get('Situação', '⚠️ N/Verif'))
                if info.empty: 
                    lista_erros.append({"NF": row['NUM_NF'], "Erro": "NCM Ausente no ICMS"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente"])
                alq_b = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - alq_b) < 0.01 else f"❌ Aliq {row['ALQ-ICMS']}% (Base {alq_b}%)"
                if "❌" in diag: lista_erros.append({"NF": row['NUM_NF'], "Erro": diag})
                return pd.Series([sit, st_e, diag])
            df_i[['Situação Nota', 'Check ST Entrada', 'Diagnóstico ICMS']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # PIS/COFINS AUDIT
            df_p = df_sai.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']]
                if info.empty: return "❌ NCM Ausente"
                cst_b = str(info.iloc[0, 2]).zfill(2) if len(info.columns) > 2 else "01"
                diag = "✅ Correto" if row['CST-PIS'] == cst_b else f"❌ CST {row['CST-PIS']} (Base {cst_b})"
                return diag
            df_p['Diagnóstico PIS/COFINS'] = df_p.apply(audit_pc, axis=1)
            df_p.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # IPI AUDIT
            df_ipi = df_sai.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']]
                if info.empty: return "❌ NCM Ausente"
                alq_b = safe_float(info.iloc[0, 2])
                diag = "✅ Correto" if abs(row['ALQ-IPI'] - alq_b) < 0.01 else f"❌ Aliq {row['ALQ-IPI']}% (Base {alq_b}%)"
                return diag
            df_ipi['Diagnóstico IPI'] = df_ipi.apply(audit_ipi, axis=1)
            df_ipi.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # DIFAL
            df_sai[['CHAVE_ACESSO', 'NUM_NF', 'CNPJ_EMIT', 'CNPJ_DEST', 'UF_EMIT', 'UF_DEST', 'VAL-DIFAL']].to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

        pd.DataFrame(lista_erros if lista_erros else [{"NF": "-", "Erro": "Tudo OK"}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
