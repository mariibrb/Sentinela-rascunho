import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

# REGRA GERAL ICMS: ALÍQUOTAS INTERNAS PADRÃO
ALIQUOTAS_UF = {
    'AC': 19.0, 'AL': 19.0, 'AM': 20.0, 'AP': 18.0, 'BA': 20.5, 'CE': 20.0,
    'DF': 20.0, 'ES': 17.0, 'GO': 19.0, 'MA': 22.0, 'MG': 18.0, 'MS': 17.0,
    'MT': 17.0, 'PA': 19.0, 'PB': 20.0, 'PE': 20.5, 'PI': 21.0, 'PR': 19.5,
    'RJ': 20.0, 'RN': 20.0, 'RO': 19.5, 'RR': 20.0, 'RS': 17.0, 'SC': 17.0,
    'SE': 19.0, 'SP': 18.0, 'TO': 20.0
}

def safe_float(v):
    if v is None or pd.isna(v) or str(v).strip().upper() == 'NT': return None
    try:
        txt = str(v).replace('R$', '').replace(' ', '').replace('%', '').replace('(', '').replace(')', '').strip()
        if ',' in txt and '.' in txt: txt = txt.replace('.', '').replace(',', '.')
        elif ',' in txt: txt = txt.replace(',', '.')
        return round(float(txt), 4)
    except: return None

def buscar_base_no_repositorio(cod_cliente):
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
            
            def buscar_tag(tag, raiz):
                alvo = raiz.find(f'.//{tag}')
                return alvo.text if alvo is not None and alvo.text is not None else ""
            
            def buscar_recursivo(node, tags_alvo):
                for elem in node.iter():
                    tag_limpa = elem.tag.split('}')[-1]
                    if tag_limpa in tags_alvo: return elem.text
                return ""

            inf = root.find('.//infNFe'); chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
            emit = root.find('.//emit'); dest = root.find('.//dest')
            uf_e = emit.find('.//UF').text if emit is not None else ""
            uf_d = dest.find('.//UF').text if dest is not None else ""

            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                icms_node = imp.find('.//ICMS') if imp is not None else None
                cst_ex = buscar_recursivo(icms_node, ['CST', 'CSOSN']) if icms_node is not None else ""
                orig_ex = buscar_recursivo(icms_node, ['orig']) if icms_node is not None else ""
                
                linha = {
                    "CHAVE_ACESSO": str(chave).strip(), 
                    "NUM_NF": buscar_tag('nNF', root),
                    "CNPJ_EMIT": (emit.find('.//CNPJ').text if emit is not None else ""),
                    "UF_EMIT": uf_e, "UF_DEST": uf_d,
                    "CFOP": prod.find('CFOP').text if prod is not None else "", 
                    "NCM": re.sub(r'\D', '', prod.find('NCM').text).zfill(8) if prod is not None else "",
                    "VPROD": safe_float(prod.find('vProd').text) or 0.0,
                    "ORIGEM": orig_ex, "CST-ICMS": cst_ex.zfill(2) if cst_ex else "",
                    "BC-ICMS": safe_float(buscar_recursivo(imp, ['vBC'])) or 0.0, 
                    "ALQ-ICMS": safe_float(buscar_recursivo(imp, ['pICMS'])), 
                    "VLR-ICMS": safe_float(buscar_recursivo(imp, ['vICMS'])) or 0.0,
                    "VLR-ICMS-ST": safe_float(buscar_recursivo(imp, ['vICMSST'])) or 0.0,
                    "CST-PIS": buscar_recursivo(imp.find('.//PIS'), ['CST']) if imp.find('.//PIS') is not None else "",
                    "VAL-PIS": safe_float(buscar_recursivo(imp.find('.//PIS'), ['vPIS'])) or 0.0,
                    "CST-COF": buscar_recursivo(imp.find('.//COFINS'), ['CST']) if imp.find('.//COFINS') is not None else "",
                    "VAL-COF": safe_float(buscar_recursivo(imp.find('.//COFINS'), ['vCOFINS'])) or 0.0,
                    "CST-IPI": buscar_recursivo(imp.find('.//IPI'), ['CST']) if imp.find('.//IPI') is not None else "",
                    "ALQ-IPI": safe_float(buscar_recursivo(imp.find('.//IPI'), ['pIPI'])),
                    "VAL-IPI": safe_float(buscar_recursivo(imp.find('.//IPI'), ['vIPI'])) or 0.0,
                    "VAL-DIFAL": safe_float(buscar_recursivo(imp, ['vICMSUFDest'])) or 0.0
                }
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_xs, ae_f, as_f, ge_f, gs_f, cod_cliente=""):
    base_f = buscar_base_no_repositorio(cod_cliente)
    try:
        base_icms = pd.read_excel(base_f, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms['NCM'].astype(str).str.zfill(8)
        base_pc = pd.read_excel(base_f, sheet_name='PIS_COFINS'); base_pc['NCM_KEY'] = base_pc['NCM'].astype(str).str.zfill(8)
        base_ipi = pd.read_excel(base_f, sheet_name='IPI'); base_ipi['NCM_KEY'] = base_ipi['NCM_TIPI'].astype(str).str.zfill(8)
    except: base_icms, base_pc, base_ipi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try: 
        tipi_padrao = pd.read_csv('394-Bases_Tributarias.xlsx - IPI.csv')
        tipi_padrao['NCM_KEY'] = tipi_padrao['NCM_TIPI'].astype(str).str.zfill(8)
    except: tipi_padrao = pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        man_l = [
            ["SENTINELA - AUDITORIA MAXIMALISTA TOTAL"], [""],
            ["1. PRIORIDADE GITHUB: NCMs na base do cliente têm precedência absoluta."],
            ["2. REGRA GERAL ICMS: Aplicada via Fallback UF quando o NCM está vazio."],
            ["3. REGRA GERAL IPI: Utiliza a TIPI padrão enviada na ausência de regra específica."],
            ["4. CST 10 / ST: Auditoria rígida de destaque e conformidade de CFOP."],
            ["5. PIS/COFINS: Confronto de CST de entrada e saída por NCM."],
            ["6. GERENCIAIS E AUTENTICIDADE: Cruzamento de dados externos com XMLs."]
        ]
        pd.DataFrame(man_l).to_excel(writer, sheet_name='MANUAL', index=False, header=False)

        def cruzar_aut(df, f):
            if df.empty or f is None: return df
            try:
                df_a = pd.read_excel(f, header=None)
                df_a[0] = df_a[0].astype(str).str.replace('NFe', '').str.strip()
                status_map = df_a.set_index(0)[5].to_dict()
                df['Situação Nota'] = df['CHAVE_ACESSO'].map(status_map).fillna('⚠️ N/Encontrada')
                return df
            except: return df
        df_xs = cruzar_aut(df_xs, as_f)

        if not df_xs.empty:
            # --- ICMS ---
            df_i = df_xs.copy()
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']] if not base_icms.empty else pd.DataFrame()
                sit = row.get('Situação Nota', '⚠️ N/V')
                cfop_st = ['5401', '5403', '5405', '6401', '6403', '6404']
                diag_st = "✅ OK"
                if row['CST-ICMS'] == '10' and row['VLR-ICMS-ST'] == 0: diag_st = "❌ Alerta: CST 10 sem destaque ST"
                if row['CFOP'] in cfop_st and row['CST-ICMS'] in ['00', '20']: diag_st = "❌ Conflito: CFOP ST com CST Trib"
                
                val_git = safe_float(info['ALIQ (INTERNA)'].iloc[0]) if not info.empty else None
                if val_git is None:
                    if row['UF_EMIT'] != row['UF_DEST']:
                        alq_esp = 4.0 if str(row['ORIGEM']) in ['1', '2', '3', '8'] else 12.0
                    else: alq_esp = ALIQUOTAS_UF.get(row['UF_EMIT'], 18.0)
                else: alq_esp = val_git

                alq_xml = row['ALQ-ICMS'] or 0.0
                diag_alq = "✅ Alq OK" if abs(alq_xml - alq_esp) < 0.01 else f"❌ XML {alq_xml}% vs {alq_esp}%"
                comp = max(0, (alq_esp - alq_xml) * row['BC-ICMS'] / 100)
                return pd.Series([sit, diag_st, diag_alq, f"R$ {comp:,.2f}"])
            
            df_i[['Situação Nota', 'Check ST', 'Diagnóstico ICMS', 'Complemento ICMS']] = df_i.apply(audit_icms, axis=1)
            df_i['Carga Efetiva (%)'] = ((df_i['VLR-ICMS'] + df_i['VAL-PIS'] + df_i['VAL-COF'] + df_i['VAL-IPI']) / (row['VPROD'] if row['VPROD'] > 0 else 1) * 100).round(2)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # --- PIS/COFINS ---
            df_pc = df_xs.copy()
            def audit_pc(row):
                info = base_pc[base_pc['NCM_KEY'] == row['NCM']] if not base_pc.empty else pd.DataFrame()
                if info.empty: return "❌ NCM ausente na Base"
                cst_b = str(info['CST Saída'].iloc[0]).zfill(2)
                return "✅ CST OK" if row['CST-PIS'] == cst_b else f"❌ XML {row['CST-PIS']} vs Base {cst_b}"
            df_pc['Diagnóstico PIS/COF'] = df_pc.apply(audit_pc, axis=1)
            df_pc.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # --- IPI ---
            df_ip = df_xs.copy()
            def audit_ipi(row):
                info = base_ipi[base_ipi['NCM_KEY'] == row['NCM']] if not base_ipi.empty else pd.DataFrame()
                info_p = tipi_padrao[tipi_padrao['NCM_KEY'] == row['NCM']] if not tipi_padrao.empty else pd.DataFrame()
                val_git = safe_float(info['ALÍQUOTA (%)'].iloc[0]) if not info.empty else None
                val_pad = safe_float(info_p['ALÍQUOTA (%)'].iloc[0]) if not info_p.empty else 0.0
                alq_esp = val_git if val_git is not None else (val_pad or 0.0)
                alq_xml = row['ALQ-IPI'] or 0.0
                return "✅ Alq OK" if abs(alq_xml - alq_esp) < 0.01 else f"❌ XML {alq_xml}% vs TIPI {alq_esp}%"
            df_ip['Diagnóstico IPI'] = df_ip.apply(audit_ipi, axis=1)
            df_ip.to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # --- DIFAL ---
            df_dif = df_xs.copy()
            df_dif['Check DIFAL'] = df_dif.apply(lambda r: "⚠️ Faltando DIFAL" if r['UF_EMIT'] != r['UF_DEST'] and r['VAL-DIFAL'] == 0 else "✅ OK", axis=1)
            df_dif.to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

    return output.getvalue()
