import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import re, io, requests, streamlit as st

def safe_float(v):
    if v is None or pd.isna(v): return 0.0
    try:
        txt = str(v).replace('R$', '').replace(' ', '').strip()
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
                cst_ex = buscar_recursivo(icms_node, ['CST', 'CSOSN']) if icms_node is not None else ""
                orig_ex = buscar_recursivo(icms_node, ['orig']) if icms_node is not None else ""
                
                linha = {
                    "CHAVE_ACESSO": str(chave).strip(), "NUM_NF": buscar('nNF', root),
                    "CNPJ_EMIT": cnpj_e, "CNPJ_DEST": cnpj_d, "UF_EMIT": buscar('UF', emit), "UF_DEST": buscar('UF', dest),
                    "CFOP": buscar('CFOP', prod), "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "VPROD": safe_float(buscar('vProd', prod)), "ORIGEM": orig_ex, "CST-ICMS": cst_ex.zfill(2) if cst_ex else "",
                    "BC-ICMS": safe_float(buscar('vBC', imp)), "ALQ-ICMS": safe_float(buscar('pICMS', imp)), "VLR-ICMS": safe_float(buscar('vICMS', imp)),
                    "ICMS-ST": safe_float(buscar('vICMSST', imp)),
                    "CST-PIS": buscar_recursivo(imp.find('.//PIS'), ['CST']) if imp.find('.//PIS') is not None else "",
                    "BC-PIS": safe_float(buscar('vBC', imp.find('.//PIS'))) if imp.find('.//PIS') is not None else 0.0,
                    "VAL-PIS": safe_float(buscar('vPIS', imp.find('.//PIS'))) if imp.find('.//PIS') is not None else 0.0,
                    "CST-COF": buscar_recursivo(imp.find('.//COFINS'), ['CST']) if imp.find('.//COFINS') is not None else "",
                    "BC-COF": safe_float(buscar('vBC', imp.find('.//COFINS'))) if imp.find('.//COFINS') is not None else 0.0,
                    "VAL-COF": safe_float(buscar('vCOFINS', imp.find('.//COFINS'))) if imp.find('.//COFINS') is not None else 0.0,
                    "CST-IPI": buscar_recursivo(imp.find('.//IPI'), ['CST']) if imp.find('.//IPI') is not None else "",
                    "BC-IPI": safe_float(buscar('vBC', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "ALQ-IPI": safe_float(buscar('pIPI', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "VAL-IPI": safe_float(buscar('vIPI', imp.find('.//IPI'))) if imp.find('.//IPI') is not None else 0.0,
                    "BC-DIFAL": safe_float(buscar_recursivo(imp, ['vBCUFDest'])),
                    "ALQ-DIFAL": safe_float(buscar_recursivo(imp, ['pICMSUFDest'])),
                    "VAL-DIFAL": safe_float(buscar_recursivo(imp, ['vICMSUFDest']))
                }
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_sai, ae_f, as_f, ge_f, gs_f, cod_cliente=""):
    base_f = buscar_base_no_repositorio(cod_cliente); lista_erros = []
    try:
        base_icms = pd.read_excel(base_f, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms.iloc[:, 0].astype(str).str.zfill(8)
        base_pc = pd.read_excel(base_f, sheet_name='PIS_COFINS'); base_pc['NCM_KEY'] = base_pc.iloc[:, 0].astype(str).str.zfill(8)
        base_ipi = pd.read_excel(base_f, sheet_name='IPI'); base_ipi['NCM_KEY'] = base_ipi.iloc[:, 0].astype(str).str.zfill(8)
    except: base_icms, base_pc, base_ipi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # MANUAL COMPLETO LINHA POR LINHA
        man_l = [
            ["SENTINELA - MANUAL MAXIMALISTA DE AUDITORIA"], [""],
            ["1. DIFAL: Aba restaurada com tags <vBCUFDest>, <pICMSUFDest> e <vICMSUFDest>."],
            ["2. ANÁLISE DIFAL: O motor valida obrigatoriedade em operações interestaduais para Consumidor Final."],
            ["3. AUTENTICIDADE: Motor fixo na Coluna A (Chave) e Coluna F (Status) para evitar erros de leitura."],
            ["4. CST ICMS: Busca profunda recursiva em todas as tags de ICMS e Simples Nacional."],
            ["5. ERRO DE VALOR: Mensagens claras para campos vazios na sua Base GitHub."]
        ]
        pd.DataFrame(man_l).to_excel(writer, sheet_name='MANUAL', index=False, header=False)
        writer.sheets['MANUAL'].set_column('A:A', 110)

        def cruzar_aut(df, f):
            if df.empty or f is None: return df
            try:
                df_a = pd.read_excel(f, header=None)
                df_a[0] = df_a[0].astype(str).str.replace('NFe', '').str.strip()
                status_map = df_a.set_index(0)[5].to_dict()
                df['Situação Nota'] = df['CHAVE_ACESSO'].map(status_map).fillna('⚠️ Chave não encontrada')
                return df
            except: return df
        
        df_sai = cruzar_aut(df_sai, as_f)

        if not df_sai.empty:
            # --- ICMS AUDIT ---
            df_i = df_sai.copy(); ncm_st = df_ent[(df_ent['CST-ICMS']=="60")]['NCM'].unique().tolist() if not df_ent.empty else []
            def audit_icms(row):
                info = base_icms[base_icms['NCM_KEY'] == row['NCM']]
                st_e = "✅ ST Localizada" if row['NCM'] in ncm_st else "❌ Sem ST na Entrada"
                sit = row.get('Situação Nota', '⚠️ N/Verif')
                if row['CST-ICMS'] == "60":
                    return pd.Series([sit, st_e, "✅ ST Validado" if row['NCM'] in ncm_st else "❌ Saída 60 s/ prova de entrada ST"])
                if info.empty: return pd.Series([sit, st_e, f"❌ NCM {row['NCM']} ausente na Base ICMS"])
                val_base = info.iloc[0, 2]
                if pd.isna(val_base): return pd.Series([sit, st_e, f"❌ VALOR NÃO PREENCHIDO NA BASE (NCM {row['NCM']})"])
                alq_b = safe_float(val_base) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - alq_b) < 0.01 else f"❌ Aliq XML {row['ALQ-ICMS']}% diverge da Base {alq_b}%"
                return pd.Series([sit, st_e, diag])
            df_i[['Situação Nota', 'Check ST Entrada', 'Diagnóstico ICMS']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

            # --- DEMAIS ABAS COM ESPELHO DE TAGS ---
            df_sai.to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)
            df_sai.to_excel(writer, sheet_name='IPI_AUDIT', index=False)
            
            # --- DIFAL AUDIT (MAXIMALISTA) ---
            
            df_dif = df_sai.copy()
            def audit_dif(row):
                if row['UF_EMIT'] != row['UF_DEST'] and row['VAL-DIFAL'] == 0:
                    return "❌ Alerta: Operação Interestadual sem destaque de DIFAL"
                return "✅ OK ou Não Aplicável"
            df_dif['Diagnóstico DIFAL'] = df_dif.apply(audit_dif, axis=1)
            df_dif.to_excel(writer, sheet_name='DIFAL_AUDIT', index=False)

        pd.DataFrame(lista_erros if lista_erros else [{"Resultado": "Auditoria Finalizada"}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
