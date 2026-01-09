import pandas as pd

UFS_BRASIL = ['AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MG', 'MS', 'MT', 'PA', 'PB', 'PE', 'PI', 'PR', 'RJ', 'RN', 'RO', 'RR', 'RS', 'SC', 'SE', 'SP', 'TO']
# Lista atualizada conforme sua última homologação
UFS_COM_IE = ['BA', 'DF', 'ES', 'GO', 'MT', 'MS', 'MG', 'PR', 'PE', 'RJ', 'RS', 'SC', 'SP']

def gerar_resumo_uf(df, writer):
    if df.empty: return
    
    df_temp = df.copy()
    
    # 1. Filtro rigoroso: Somente notas autorizadas e não canceladas
    df_validas = df_temp[
        (df_temp['Situação Nota'].astype(str).str.upper().str.contains('AUTORIZAD', na=False)) &
        (~df_temp['Situação Nota'].astype(str).str.upper().str.contains('CANCEL', na=False))
    ].copy()

    # 2. Identificação de Sentido pelo CFOP
    def identificar_sentido(cfop):
        c = str(cfop).strip()[0]
        if c in ['1', '2', '3']: return 'ENTRADA'
        if c in ['5', '6', '7']: return 'SAÍDA'
        return 'OUTROS'
    
    df_validas['SENTIDO'] = df_validas['CFOP'].apply(identificar_sentido)

    # 3. Preparação da Tabela Completa (27 UFs)
    def preparar_tabela_completa(df_origem):
        # Soma tudo o que pode compor o saldo devedor/credor por UF
        agrupado = df_origem.groupby(['UF_DEST']).agg({
            'VAL-ICMS-ST': 'sum', 
            'VAL-DIFAL': 'sum', 
            'VAL-FCP-DEST': 'sum',
            'VAL-FCP': 'sum', 
            'VAL-FCP-ST': 'sum'
        }).reset_index()
        
        # Consolidação: DIFAL TOTAL = DIFAL Puro + FCP Destino
        agrupado['DIFAL_CONSOLIDADO'] = agrupado['VAL-DIFAL'] + agrupado['VAL-FCP-DEST']
        
        base = pd.DataFrame({'UF_DEST': UFS_BRASIL})
        final = pd.merge(base, agrupado, on='UF_DEST', how='left').fillna(0)
        
        # Busca a IE de Substituto na UF (se existir em alguma nota)
        ie_map = df_origem[df_origem['IE_SUBST'] != ""].groupby('UF_DEST')['IE_SUBST'].first().to_dict()
        final['IE_SUBST'] = final['UF_DEST'].map(ie_map).fillna("")
        
        return final[['UF_DEST', 'IE_SUBST', 'VAL-ICMS-ST', 'DIFAL_CONSOLIDADO', 'VAL-FCP', 'VAL-FCP-ST']]

    res_s = preparar_tabela_completa(df_validas[df_validas['SENTIDO'] == 'SAÍDA'])
    res_e = preparar_tabela_completa(df_validas[df_validas['SENTIDO'] == 'ENTRADA'])

    # 4. Cálculo do Saldo Líquido (Apenas estados da lista abatem devolução)
    res_saldo = pd.DataFrame({'ESTADO (UF)': UFS_BRASIL})
    res_saldo['IE SUBSTITUTO'] = res_s['IE_SUBST']
    
    cols_calculo = [
        ('VAL-ICMS-ST', 'ST LÍQUIDO'), 
        ('DIFAL_CONSOLIDADO', 'DIFAL LÍQUIDO'), 
        ('VAL-FCP', 'FCP LÍQUIDO'), 
        ('VAL-FCP-ST', 'FCP-ST LÍQUIDO')
    ]
    
    for c_xml, c_fin in cols_calculo:
        res_saldo[c_fin] = res_saldo['ESTADO (UF)'].apply(
            lambda x: (res_s.loc[res_s['UF_DEST']==x, c_xml].values[0] - res_e.loc[res_e['UF_DEST']==x, c_xml].values[0])
            if x in UFS_COM_IE else res_s.loc[res_s['UF_DEST']==x, c_xml].values[0]
        )

    # 5. Gravação Excel (Layout Side-by-Side sem Gridlines)
    workbook = writer.book
    worksheet = workbook.add_worksheet('DIFAL_ST_FECP')
    writer.sheets['DIFAL_ST_FECP'] = worksheet
    worksheet.hide_gridlines(2)

    # Formatos
    title_fmt = workbook.add_format({'bold': True, 'font_color': '#FF6F00', 'font_size': 11})
    orange_fill = workbook.add_format({'bg_color': '#FFDAB9', 'border': 1})
    orange_num = workbook.add_format({'bg_color': '#FFDAB9', 'border': 1, 'num_format': '#,##0.00'})
    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E0E0E0'})
    num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
    border_fmt = workbook.add_format({'border': 1})
    total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'num_format': '#,##0.00'})

    # Cabeçalhos
    h_s = ['UF', 'IE SUBST', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
    h_e = ['UF ', 'IE SUBST ', 'ST TOTAL ', 'DIFAL TOTAL ', 'FCP TOTAL ', 'FCP-ST TOTAL ']
    h_l = ['UF  ', 'IE SUBST  ', 'ST LÍQUIDO', 'DIFAL LÍQUIDO', 'FCP LÍQUIDO', 'FCP-ST LÍQUIDO']

    for df_t, start_c, title, heads in [(res_s, 0, "1. SAÍDAS", h_s), (res_e, 7, "2. ENTRADAS", h_e), (res_saldo, 14, "3. SALDO", h_l)]:
        worksheet.write(0, start_c, title, title_fmt)
        for i, h in enumerate(heads): worksheet.write(2, start_c+i, h, header_fmt)
        for r_idx, row in enumerate(df_t.values):
            uf = str(row[0]).strip()
            for c_idx, val in enumerate(row):
                fmt = (orange_num if isinstance(val, (int,float)) else orange_fill) if uf in UFS_COM_IE else (num_fmt if isinstance(val, (int,float)) else border_fmt)
                worksheet.write(r_idx+3, start_c+c_idx, val, fmt)
        # Totais
        worksheet.write(30, start_c, "TOTAL GERAL", total_fmt)
        worksheet.write(30, start_c+1, "", total_fmt)
        for i in range(2, 6):
            let = chr(65+start_c+i) if (start_c+i)<26 else f"A{chr(65+start_c+i-26)}"
            worksheet.write(30, start_c+i, f'=SUM({let}4:{let}30)', total_fmt)
