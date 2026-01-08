import pandas as pd

def gerar_resumo_uf(df, writer):
    """
    Gera a aba DIFAL_ST_FECP separando Entradas e Saídas.
    Utiliza o CFOP para identificar o sentido da operação.
    """
    if df.empty:
        return

    df_temp = df.copy()

    # 1. Filtro de Notas Autorizadas
    df_aut = df_temp[
        df_temp['Situação Nota'].astype(str).str.upper().str.contains('AUTORIZAD', na=False)
    ].copy()

    if df_aut.empty:
        pd.DataFrame([["Aviso:", "Nenhuma nota AUTORIZADA encontrada."]]).to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False, header=False)
        return

    # 2. Identificação de Sentido (Entrada vs Saída) pelo CFOP
    # CFOP iniciando em 1, 2 ou 3 = ENTRADA
    # CFOP iniciando em 5, 6 ou 7 = SAÍDA
    def identificar_sentido(cfop):
        c = str(cfop).strip()[0]
        if c in ['1', '2', '3']: return 'ENTRADA'
        if c in ['5', '6', '7']: return 'SAÍDA'
        return 'OUTROS'

    df_aut['SENTIDO'] = df_aut['CFOP'].apply(identificar_sentido)

    # 3. Função para agrupar e formatar
    def agrupar_dados(dataframe):
        return dataframe.groupby(['UF_DEST', 'IE_SUBST']).agg({
            'VAL-ICMS-ST': 'sum',
            'VAL-DIFAL': 'sum',
            'VAL-FCP': 'sum',
            'VAL-FCP-ST': 'sum'
        }).reset_index()

    # Separa os dataframes
    df_saidas = df_aut[df_aut['SENTIDO'] == 'SAÍDA']
    df_entradas = df_aut[df_aut['SENTIDO'] == 'ENTRADA']

    # 4. Gravação Física na Aba
    start_row = 0
    workbook = writer.book
    worksheet = workbook.add_worksheet('DIFAL_ST_FECP')
    writer.sheets['DIFAL_ST_FECP'] = worksheet

    # Formatos
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6F00', 'font_color': 'white', 'border': 1})
    title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#FF6F00'})

    # --- TABELA DE SAÍDAS ---
    worksheet.write(start_row, 0, "RESUMO DE SAÍDAS (VENDAS)", title_fmt)
    start_row += 2
    
    if not df_saidas.empty:
        res_s = agrupar_dados(df_saidas)
        res_s.columns = ['ESTADO (UF)', 'IE SUBSTITUTO', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
        res_s.to_excel(writer, sheet_name='DIFAL_ST_FECP', startrow=start_row, index=False)
        start_row += len(res_s) + 4
    else:
        worksheet.write(start_row, 0, "Nenhuma saída encontrada.")
        start_row += 3

    # --- TABELA DE ENTRADAS (DEVOLUÇÕES) ---
    worksheet.write(start_row, 0, "RESUMO DE ENTRADAS (DEVOLUÇÕES/COMPRAS)", title_fmt)
    start_row += 2

    if not df_entradas.empty:
        res_e = agrupar_dados(df_entradas)
        res_e.columns = ['ESTADO (UF)', 'IE SUBSTITUTO', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
        res_e.to_excel(writer, sheet_name='DIFAL_ST_FECP', startrow=start_row, index=False)
    else:
        worksheet.write(start_row, 0, "Nenhuma entrada encontrada.")
