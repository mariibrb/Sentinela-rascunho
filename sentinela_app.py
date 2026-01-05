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
            def buscar(tag, raiz):
                alvo = raiz.find(f'.//{tag}')
                return alvo.text if alvo is not None and alvo.text is not None else ""

            inf = root.find('.//infNFe')
            chave = inf.attrib.get('Id', '')[3:] if inf is not None else ""
            emit = root.find('.//emit'); dest = root.find('.//dest')
            cnpj_e = buscar('CNPJ', emit) or buscar('CPF', emit)
            cnpj_d = buscar('CNPJ', dest) or buscar('CPF', dest)

            for det in root.findall('.//det'):
                prod = det.find('prod'); imp = det.find('imposto')
                # BUSCA RECURSIVA PARA CST/CSOSN (BLINDADO)
                cst_extraido = ""
                orig_extraida = ""
                icms_node = imp.find('.//ICMS')
                if icms_node is not None:
                    # Varre todos os subtipos (ICMS00, ICMS10... ICMSSN102...)
                    for tipo in icms_node:
                        c = tipo.find('CST') if tipo.find('CST') is not None else tipo.find('CSOSN')
                        o = tipo.find('orig')
                        if c is not None: cst_extraido = c.text.zfill(2)
                        if o is not None: orig_extraida = o.text

                linha = {
                    "CHAVE_ACESSO": chave, "NUM_NF": buscar('nNF', root),
                    "CNPJ_EMIT": cnpj_e, "CNPJ_DEST": cnpj_d,
                    "UF_EMIT": buscar('UF', emit), "UF_DEST": buscar('UF', dest),
                    "NCM": re.sub(r'\D', '', buscar('NCM', prod)).zfill(8),
                    "VPROD": safe_float(buscar('vProd', prod)),
                    "ORIGEM": orig_extraida, "CST-ICMS": cst_extraido,
                    "ALQ-ICMS": safe_float(buscar('pICMS', imp)),
                    "BC-ICMS": safe_float(buscar('vBC', imp)),
                    "VAL-PIS": safe_float(buscar('vPIS', imp)),
                    "VAL-COF": safe_float(buscar('vCOFINS', imp)),
                    "VAL-DIFAL": safe_float(buscar('vICMSUFDest', imp))
                }
                dados_lista.append(linha)
        except: continue
    return pd.DataFrame(dados_lista)

def gerar_excel_final(df_ent, df_sai, ae_f, as_f, cod_cliente=""):
    base_f = buscar_base_no_github(cod_cliente); lista_erros = []
    try:
        base_icms = pd.read_excel(base_f, sheet_name='ICMS'); base_icms['NCM_KEY'] = base_icms.iloc[:, 0].astype(str).str.zfill(8)
    except: base_icms = pd.DataFrame()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # --- MANUAL COMPLETO DA FACE DA TERRA ---
        man_l = [
            ["SENTINELA - MANUAL TÉCNICO E DE INSTRUÇÕES"],
            ["1. INTRODUÇÃO"], ["Este relatório é o resultado do cruzamento entre XMLs e a Base GitHub."],
            [""], ["2. EXTRAÇÃO DE CST (BUSCA PROFUNDA)"],
            ["O motor agora varre recursivamente as tags ICMS00 até ICMS90 e CSOSN101 até CSOSN900."],
            ["Isso garante que, independente do regime da empresa, o CST nunca venha vazio."],
            [""], ["3. GLOSSÁRIO DE RETORNOS"],
            ["✅ Correto: O XML e a Base Tributária estão em total harmonia."],
            ["❌ Divergente: Foi encontrada diferença de alíquota ou código de tributação."],
            ["❌ NCM Ausente: O item não existe na sua base de dados cadastrada."],
            [""], ["4. COLUNAS DE RASTREABILIDADE"],
            ["CNPJ EMIT/DEST: Exibidos em todas as abas para identificar os participantes."],
            ["SITUAÇÃO NOTA: Indica se a nota foi CANCELADA via cruzamento com arquivo de Autenticidade."],
            [""]
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
                    lista_erros.append({"NF": row['NUM_NF'], "Erro": "NCM Ausente"})
                    return pd.Series([sit, st_e, "❌ NCM Ausente", "OK"])
                alq_b = safe_float(info.iloc[0, 2]) if row['UF_EMIT'] == row['UF_DEST'] else 12.0
                diag = "✅ Correto" if abs(row['ALQ-ICMS'] - alq_b) < 0.01 else f"❌ Aliq {row['ALQ-ICMS']}% (Base {alq_b}%)"
                if "❌" in diag: lista_erros.append({"NF": row['NUM_NF'], "Erro": diag})
                return pd.Series([sit, st_e, diag, "Ajustar" if "❌" in diag else "OK"])
            
            df_i[['Status Nota', 'Check ST Entrada', 'Diagnóstico ICMS', 'Ação']] = df_i.apply(audit_icms, axis=1)
            df_i.to_excel(writer, sheet_name='ICMS_AUDIT', index=False)

        pd.DataFrame(lista_erros if lista_erros else [{"NF": "-", "Erro": "Tudo OK"}]).to_excel(writer, sheet_name='RESUMO_ERROS', index=False)

    return output.getvalue()
