import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

# TABELA DE ALÍQUOTAS INTERNAS PADRÃO
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
            def buscar_tag(tag, node):
                alvo = node.find(f'.//{tag}')
                return alvo.text if alvo is not None and alvo.text is not None else ""
            def buscar_recursivo(node, tags_alvo):
                if node is None: return ""
                for elem in node.iter():
                    tag_limpa = elem.tag.split('}')[-1]
                    if tag_limpa in tags_alvo: return elem.text
                return ""
            
            inf = root.find('.//infNFe'); emit = root.find('.//emit'); dest = root.find('.//dest')
            chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
            
            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                icms_node = imp.find('.//ICMS') if imp is not None else None
                linha = {
                    "CHAVE_ACESSO": str(chave).strip(), 
                    "NUM_NF": buscar_tag('nNF', root),
                    "CNPJ_EMIT": buscar_tag('CNPJ', emit),
                    "CNPJ_DEST": buscar_tag('CNPJ', dest),
                    "CPF_DEST": buscar_tag('CPF', dest),
                    "UF_EMIT": buscar_tag('UF', emit),
                    "UF_DEST": buscar_tag('UF', dest),
                    "indIEDest": buscar_tag('indIEDest', dest),
                    "CFOP": buscar_tag('CFOP', prod),
                    "NCM": re.sub(r'\D', '', buscar_tag('NCM', prod)).zfill(8),
                    "VPROD": safe_float(buscar_tag('vProd', prod)),
                    "ORIGEM": buscar_recursivo(icms_node, ['orig']),
                    "CST-ICMS": buscar_recursivo(icms_node, ['CST', 'CSOSN']).zfill(2),
                    "BC-ICMS": safe_float(buscar_recursivo(imp, ['vBC'])),
                    "ALQ-ICMS": safe_float(buscar_recursivo(imp, ['pICMS'])),
                    "VLR-ICMS": safe_float(buscar_recursivo(imp, ['vICMS'])),
                    "VLR-ICMS-ST": safe_float(buscar_recursivo(imp, ['vICMSST'])),
                    "CST-PIS": buscar_recursivo(imp.find('.//PIS'), ['CST']),
                    "VAL-PIS": safe_float(buscar_recursivo(imp.find('.//PIS'), ['vPIS'])),
                    "CST-COF": buscar_recursivo(imp.find('.//COFINS'), ['CST']),
                    "VAL-COF": safe_float(buscar_recursivo(imp.find('.//COFINS'), ['vCOFINS'])),
                    "CST-IPI": buscar_recursivo(imp.find('.//IPI'), ['CST']),
                    "ALQ-IPI": safe_float(buscar_recursivo(imp.find('.//IPI'), ['pIPI'])),
                    "VAL-IPI": safe_float(buscar_recursivo(imp.find('.//IPI'), ['vIPI'])),
                    "VAL-DIFAL": safe_float(buscar_recursivo(imp, ['vICMSUFDest'])),
                    "VAL-FCP-DEST": safe_float(buscar_recursivo(imp, ['vFCPUFDest']))
                }
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_xe, df_xs, ae, as_f, ge, gs, cod_cliente):
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

    def load_status_map(f):
        if not f: return {}
        try:
            f.seek(0)
            df = pd.read_excel(f, header=None) if f.name.endswith('.xlsx') else pd.read_csv(f, header=None)
            df[0] = df[0].astype(str).str.replace('NFe', '').str.strip()
            return df.set_index(0)[5].to_dict()
        except: return {}
    
    status_map = load_status_map(as_f)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([["AUDITORIA FISCAL SENTINELA"]]).to_excel(writer, sheet_name='MANUAL', index=False, header=False)
        
        # 👣 GERENCIAIS COMO ABAS REAIS
        for f, s in [(ge, 'GERENCIAL_ENTRADA'), (gs, 'GERENCIAL_SAIDA')]:
            if f:
                try:
                    f.seek(0)
                    df_g = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
                    df_g.to_excel(writer, sheet_name=s, index=False)
                except: pass

        if not df_xs.empty:
            df_xs['Situação Nota'] = df_xs['CHAVE_ACESSO'].map(status_map).fillna('⚠️ N/Encontrada')
            
            # --- 1. ICMS AUDIT (4% IMPORTADOS + REGRA GERAL) ---
            df_i = df_xs.copy()
            def audit_icms(r):
                info = base_icms[base_icms['NCM_KEY'] == r['NCM']] if not base_icms.empty else pd.DataFrame()
                diag_st = "✅ OK"
                if r['CST-ICMS'] == '10' and r['VLR-ICMS-ST'] == 0: diag_st = "❌ Alerta: CST 10 sem destaque ST"
                
                val_git = safe_float(info['ALIQ (INTERNA)'].iloc[0]) if not info.empty else 0.0
                if val_git == 0:
                    if r['UF_EMIT'] != r['UF_DEST']:
                        # Regra de Ouro: Origem 1,2,3,8 interestadual trava em 4%
                        alq_esp = 4.0 if str(r['ORIGEM']) in ['1', '2', '3', '8'] else 12.0
                        fonte = "Interestadual (Regra Geral)"
                    else:
                        alq_esp = ALIQUOTAS_UF.get(r['UF_EMIT'], 18.0)
                        fonte = f"Interna {r['UF_EMIT']} (Regra Geral)"
                else: 
                    alq_esp = val_git; fonte = "Base Cadastrada"

                diag_alq = "✅ Alq OK" if abs(r['ALQ-ICMS'] - alq_esp) < 0.01 else f"❌ XML {r['ALQ-ICMS']}% vs {alq_esp}%"
                comp = max(0, (alq_esp - r['ALQ-ICMS']) * r['BC-ICMS'] / 100)
                return pd.Series([r['Situação Nota'], fonte, diag_st, diag_alq, f"R$ {comp:,.2f}"])
            
            df_i[['Situação Nota', 'Fonte', 'Check ST', 'Diagnóstico', 'Complemento']] = df_i.apply(audit_icms, axis=1)
            cols_i = ['Situação Nota', 'Fonte', 'Check ST', 'Diagnóstico', 'Complemento'] + [c for c in df_i.columns if c not in ['Situação Nota', 'Fonte', 'Check ST', 'Diagnóstico', 'Complemento']]
            df_i[cols_i].to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # --- 2. PIS/COFINS AUDIT ---
            df_pc = df_xs.copy()
            def audit_pc(r):
                info = base_pc[base_pc['NCM_KEY'] == r['NCM']] if not base_pc.empty else pd.DataFrame()
                if info.empty: return "❌ NCM ausente na Base"
                cst_b = str(info['CST Saída'].iloc[0]).zfill(2)
                return "✅ CST OK" if r['CST-PIS'] == cst_b else f"❌ XML {r['CST-PIS']} vs Base {cst_b}"
            df_pc['Check PIS/COF'] = df_pc.apply(audit_pc, axis=1)
            cols_pc = ['Situação Nota', 'Check PIS/COF'] + [c for c in df_pc.columns if c not in ['Situação Nota', 'Check PIS/COF']]
            df_pc[cols_pc].to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)

            # --- 3. IPI AUDIT (TIPI PADRÃO) ---
            df_ip = df_xs.copy()
            def audit_ipi(r):
                info_p = tipi_padrao[tipi_padrao['NCM_KEY'] == r['NCM']] if not tipi_padrao.empty else pd.DataFrame()
                val_p = safe_float(info_p['ALÍQUOTA (%)'].iloc[0]) if not info_p.empty else 0.0
                return "✅ Alq OK" if abs(r['ALQ-IPI'] - val_p) < 0.01 else f"❌ XML {r['ALQ-IPI']}% vs TIPI {val_p}%"
            df_ip['Check IPI'] = df_ip.apply(audit_ipi, axis=1)
            cols_ip = ['Situação Nota', 'Check IPI'] + [c for c in df_ip.columns if c not in ['Situação Nota', 'Check IPI']]
            df_ip[cols_ip].to_excel(writer, sheet_name='IPI_AUDIT', index=False)

            # --- 4. DIFAL AUDIT (BASEADO EM CPF/CNPJ) ---
            df_dif = df_xs.copy()
            def audit_difal(r):
                if r['UF_EMIT'] != r['UF_DEST']:
                    v_xml = r['VAL-DIFAL'] + r['VAL-FCP-DEST']
                    # CPF preenchido ou indIEDest 9 = Não Contribuinte
                    if (r['CPF_DEST'] and len(str(r['CPF_DEST'])) > 5) or r['indIEDest'] == '9':
                        return "✅ DIFAL OK" if v_xml > 0 else "⚠️ Alerta: Não-Contribuinte sem DIFAL destacado"
                    return "Contribuinte: Verificar Diferencial Interno"
                return "Operação Interna"
            df_dif['Análise DIFAL'] = df_dif.apply(audit_difal, axis=1)
            cols_d = ['Situação Nota', 'Análise DIFAL'] + [c for c in df_dif.columns if c not in ['Situação Nota', 'Análise DIFAL']]
            df_dif[cols_d].to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

    return output.getvalue()
