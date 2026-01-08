import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st, zipfile

ALIQUOTAS_UF = {
    'AC': 19.0, 'AL': 19.0, 'AM': 20.0, 'AP': 18.0, 'BA': 20.5, 'CE': 20.0,
    'DF': 20.0, 'ES': 17.0, 'GO': 19.0, 'MA': 22.0, 'MG': 18.0, 'MS': 17.0,
    'MT': 17.0, 'PA': 19.0, 'PB': 20.0, 'PE': 20.5, 'PI': 21.0, 'PR': 19.5,
    'RJ': 20.0, 'RN': 20.0, 'RO': 19.5, 'RR': 20.0, 'RS': 17.0, 'SC': 17.0,
    'SE': 19.0, 'SP': 18.0, 'TO': 20.0
}

def safe_float(v):
    if v is None or pd.isna(v) or str(v).strip().upper() in ['NT', '']: return 0.0
    try:
        txt = str(v).replace('R$', '').replace(' ', '').replace('%', '').strip()
        if ',' in txt and '.' in txt: txt = txt.replace('.', '').replace(',', '.')
        elif ',' in txt: txt = txt.replace(',', '.')
        return round(float(txt), 4)
    except: return 0.0

def buscar_github(nome_arquivo):
    token = st.secrets.get("GITHUB_TOKEN")
    repo = st.secrets.get("GITHUB_REPO")
    url = f"https://api.github.com/repos/{repo}/contents/Bases_Tributárias/{nome_arquivo}"
    headers = {"Authorization": f"token {token}"}
    try:
        res = requests.get(url, headers=headers, timeout=25)
        if res.status_code == 200:
            f_res = requests.get(res.json()['download_url'], headers=headers)
            return io.BytesIO(f_res.content)
    except: pass
    return None

def processar_conteudo_xml(content, dados_lista):
    try:
        xml_str = content.decode('utf-8', errors='replace')
        xml_str = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', xml_str)
        root = ET.fromstring(xml_str)
        def tag_val(t, n):
            v = n.find(f'.//{t}')
            return v.text if v is not None and v.text else ""
        def rec_val(n, ts):
            if n is None: return ""
            for e in n.iter():
                if e.tag.split('}')[-1] in ts: return e.text
            return ""
        inf = root.find('.//infNFe'); emit = root.find('.//emit'); dest = root.find('.//dest')
        chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
        for det in root.findall('.//det'):
            prod = det.find('prod'); imp = det.find('imposto')
            icms = imp.find('.//ICMS') if imp is not None else None
            # ORDEM DAS TAGS RESTAURADA EXATAMENTE COMO VOCÊ DEFINEU
            linha = {
                "CHAVE_ACESSO": str(chave).strip(),
                "NUM_NF": tag_val('nNF', root),
                "CNPJ_EMIT": tag_val('CNPJ', emit),
                "CNPJ_DEST": tag_val('CNPJ', dest),
                "CPF_DEST": tag_val('CPF', dest),
                "UF_EMIT": tag_val('UF', emit),
                "UF_DEST": tag_val('UF', dest),
                "indIEDest": tag_val('indIEDest', dest),
                "CFOP": tag_val('CFOP', prod),
                "NCM": re.sub(r'\D', '', tag_val('NCM', prod)).zfill(8),
                "VPROD": safe_float(tag_val('vProd', prod)),
                "ORIGEM": rec_val(icms, ['orig']),
                "CST-ICMS": rec_val(icms, ['CST', 'CSOSN']).zfill(2),
                "BC-ICMS": safe_float(rec_val(imp, ['vBC'])),
                "ALQ-ICMS": safe_float(rec_val(imp, ['pICMS'])),
                "VLR-ICMS": safe_float(rec_val(imp, ['vICMS'])),
                "CST-PIS": rec_val(imp.find('.//PIS'), ['CST']),
                "VAL-PIS": safe_float(rec_val(imp.find('.//PIS'), ['vPIS'])),
                "CST-COF": rec_val(imp.find('.//COFINS'), ['CST']),
                "VAL-COF": safe_float(rec_val(imp.find('.//COFINS'), ['vCOFINS'])),
                "CST-IPI": rec_val(imp.find('.//IPI'), ['CST']),
                "ALQ-IPI": safe_float(rec_val(imp.find('.//IPI'), ['pIPI'])),
                "VAL-IPI": safe_float(rec_val(imp.find('.//IPI'), ['vIPI'])),
                "VAL-DIFAL": safe_float(rec_val(imp, ['vICMSUFDest'])),
                "VAL-FCP-DEST": safe_float(rec_val(imp, ['vFCPUFDest'])),
                "VAL-ICMS-ST": safe_float(rec_val(imp, ['vICMSST'])),
                "BC-ICMS-ST": safe_float(rec_val(imp, ['vBCST'])),
                "VAL-FCP-ST": safe_float(rec_val(imp, ['vFCPST'])),
                "VAL-FCP": safe_float(rec_val(imp, ['vFCP'])),
                "IE_SUBST": tag_val('IEST', icms) if icms is not None else "",
                "VAL-IBS": safe_float(rec_val(imp, ['vIBS'])),
                "ALQ-IBS": safe_float(rec_val(imp, ['pIBS'])),
                "VAL-CBS": safe_float(rec_val(imp, ['vCBS'])),
                "ALQ-CBS": safe_float(rec_val(imp, ['pCBS']))
            }
            dados_lista.append(linha)
    except: pass

def extrair_dados_xml(f):
    dados_lista = []
    if not f: return pd.DataFrame()
    with zipfile.ZipFile(f) as z:
        for filename in z.namelist():
            if filename.endswith('.xml'):
                with z.open(filename) as xml_file: processar_conteudo_xml(xml_file.read(), dados_lista)
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, ae, as_f, ge, gs, cod_cliente):
    f_cliente = buscar_github(f"{cod_cliente}-Bases_Tributárias.xlsx")
    f_tipi = buscar_github("TIPI.csv")
    try:
        base_icms = pd.read_excel(f_cliente, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms['NCM'].astype(str).str.zfill(8)
        base_pc = pd.read_excel(f_cliente, sheet_name='PIS_COFINS'); base_pc['NCM_KEY'] = base_pc['NCM'].astype(str).str.zfill(8)
    except: base_icms, base_pc = pd.DataFrame(), pd.DataFrame()
    try: tipi_df = pd.read_csv(f_tipi); tipi_df['NCM_KEY'] = tipi_df['NCM'].astype(str).str.replace('.', '').str.strip().str.zfill(8)
    except: tipi_df = pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([["RELATÓRIO SENTINELA"]]).to_excel(writer, sheet_name='RESUMO', index=False, header=False)
        
        # ABAS GERENCIAIS (FIXAS)
        for f_obj, s_name in [(ge, 'GERENCIAL_ENTRADA'), (gs, 'GERENCIAL_SAIDA')]:
            if f_obj:
                try:
                    f_obj.seek(0)
                    df_g = pd.read_excel(f_obj) if f_obj.name.endswith('.xlsx') else pd.read_csv(f_obj)
                    df_g.to_excel(writer, sheet_name=s_name, index=False)
                except: pass

        st_map = {}
        if as_f:
            try:
                as_f.seek(0)
                df_auth = pd.read_excel(as_f, header=None) if as_f.name.endswith('.xlsx') else pd.read_csv(as_f, header=None)
                df_auth[0] = df_auth[0].astype(str).str.replace('NFe', '').str.strip(); st_map = df_auth.set_index(0)[5].to_dict()
            except: pass

        if not df_xs.empty:
            df_xs['Situação Nota'] = df_xs['CHAVE_ACESSO'].map(st_map).fillna('⚠️ N/Encontrada')
            
            # 1. ICMS_AUDIT
            df_i = df_xs.copy()
            def audit_icms(r):
                info = base_icms[base_icms['NCM_KEY'] == r['NCM']] if not base_icms.empty else pd.DataFrame()
                val_b = safe_float(info['ALIQ (INTERNA)'].iloc[0]) if not info.empty else 0.0
                if val_b == 0:
                    if r['UF_EMIT'] != r['UF_DEST']: alq_e = 4.0 if str(r['ORIGEM']) in ['1', '2', '3', '8'] else 12.0
                    else: alq_e = ALIQUOTAS_UF.get(r['UF_EMIT'], 18.0)
                else: alq_e = val_b
                diag = "✅ Alq OK" if abs(r['ALQ-ICMS'] - alq_e) < 0.01 else f"❌ XML {r['ALQ-ICMS']}% vs {alq_e}%"
                comp = max(0, (alq_e - r['ALQ-ICMS']) * r['BC-ICMS'] / 100)
                return pd.Series([diag, f"R$ {comp:,.2f}"])
            df_i[['Diagnóstico', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            # COLOCO O DIAGNÓSTICO E O STATUS NO INÍCIO, MANTENDO A ORDEM DAS TAGS DEPOIS
            cols_fixed = ['Situação Nota', 'Diagnóstico', 'Complemento']
            cols_final = cols_fixed + [c for c in df_i.columns if c not in cols_fixed]
            df_i[cols_final].to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # 2. IPI_AUDIT
            df_ip = df_xs.copy()
            def audit_ipi(r):
                match = tipi_df[tipi_df['NCM_KEY'] == r['NCM']] if not tipi_df.empty else pd.DataFrame()
                val_p = safe_float(match['ALÍQUOTA (%)'].iloc[0]) if not match.empty else 0.0
                return "✅ Alq OK" if abs(r['ALQ-IPI'] - val_p) < 0.01 else f"❌ XML {r['ALQ-IPI']}% vs TIPI {val_p}%"
            df_ip['Diagnóstico IPI'] = df_ip.apply(audit_ipi, axis=1)
            cols_fixed_ipi = ['Situação Nota', 'Diagnóstico IPI']
            cols_final_ipi = cols_fixed_ipi + [c for c in df_ip.columns if c not in cols_fixed_ipi]
            df_ip[cols_final_ipi].to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # 3. PIS_COFINS_AUDIT
            df_pc = df_xs.copy()
            def audit_pc(r):
                info = base_pc[base_pc['NCM_KEY'] == r['NCM']] if not base_pc.empty else pd.DataFrame()
                if info.empty: return "❌ NCM ausente na Base"
                cst_b = str(info['CST Saída'].iloc[0]).zfill(2)
                return "✅ CST OK" if r['CST-PIS'] == cst_b else f"❌ XML {r['CST-PIS']} vs Base {cst_b}"
            df_pc['Diagnóstico P/C'] = df_pc.apply(audit_pc, axis=1)
            cols_fixed_pc = ['Situação Nota', 'Diagnóstico P/C']
            cols_final_pc = cols_fixed_pc + [c for c in df_pc.columns if c not in cols_fixed_pc]
            df_pc[cols_final_pc].to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # 4. DIFAL_ST_FECP (TABELA POR ESTADO)
            df_res_uf = df_xs.groupby('UF_DEST').agg({'IE_SUBST': 'first', 'VAL-ICMS-ST': 'sum', 'VAL-DIFAL': 'sum', 'VAL-FCP': 'sum', 'VAL-FCP-ST': 'sum'}).reset_index()
            df_res_uf.columns = ['ESTADO', 'IE SUBST.', 'ST', 'DIFAL', 'FCP', 'FCP-ST']
            df_res_uf.to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False)

    return output.getvalue()
