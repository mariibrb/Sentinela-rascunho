import pandas as pd

# Lista oficial das 27 UFs do Brasil para garantir a presença de todos os estados
UFS_BRASIL = [
    'AC', 'AL', 'AM', 'AP', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MG', 'MS', 'MT',
    'PA', 'PB', 'PE', 'PI', 'PR', 'RJ', 'RN', 'RO', 'RR', 'RS', 'SC', 'SE', 'SP', 'TO'
]

def gerar_resumo_uf(df, writer):
    """
    Gera a aba DIFAL_ST_FECP com três tabelas lado a lado:
    SAÍDAS | ENTRADAS | SALDO LÍQUIDO
    Filtro rigoroso: Ignora notas canceladas (ex: Cancelamento de NF-e homologado).
    Preenche todos os estados do Brasil, mesmo que zerados.
    """
    if df.empty:
        return

    df_temp = df.copy()

    # 1. Filtro Rigoroso de Notas Válidas
    # Aceita variações de 'Autorizada' e descarta 'Cancelamento', 'Inutilizada', etc.
    df_validas = df_temp[
        (df_temp['Situação Nota'].astype(str).str.upper().str.contains('AUTORIZAD', na=False)) &
        (~df_temp['Situação Nota'].astype(str).str.upper().str.contains('CANCEL', na=False))
    ].copy()

    if df_validas.empty:
        # Se não houver notas válidas, cria a estrutura zerada para não quebrar o Excel
        pd.DataFrame([["Aviso:", "Nenhuma nota AUTORIZADA e VÁLIDA encontrada."]]).to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False, header=False)
        return

    # 2. Identificação de Sentido pelo CFOP (1,2,3 Entrada | 5,6,7 Saída)
    def identificar_sentido(cfop):
        c = str(cfop).strip()[0]
        if c in ['1', '2', '3']: return 'ENTRADA'
        if c in ['5', '6', '7']: return 'SAÍDA'
        return 'OUTROS'

    df_validas['SENTIDO'] = df_validas['CFOP'].apply(identificar_sentido)

    # 3. Função Especialista para preparar a tabela completa (27 UFs)
    def preparar_tabela_completa(dataframe_origem):
        # Agrupamento dos valores
        agrupado = dataframe_origem.groupby(['UF_DEST']).agg({
            'VAL-ICMS-ST': 'sum', 'VAL-DIFAL': 'sum', 'VAL-FCP': 'sum', 'VAL-FCP-ST': 'sum'
        }).reset_index()
        
        # Mapeamento da primeira IE de substituto encontrada para cada UF
        ie_map = dataframe_origem.groupby('UF_DEST')['IE_SUBST'].first().to_dict()
        
        # Base fixa com todos os estados brasileiros
        base_completa = pd.DataFrame({'UF_DEST': UFS_BRASIL})
        
        # Merge (Cruzamento) para garantir que todos os estados apareçam
        final = pd.merge(base_completa, agrupado, on='UF_DEST', how='left').fillna(0)
        
        # Inclusão da IE no mapeamento
        final['IE_SUBST'] = final['UF_DEST'].map(ie_map).fillna("")
        
        return final[['UF_DEST', 'IE_SUBST', 'VAL-ICMS-ST', 'VAL-DIFAL', 'VAL-FCP', 'VAL-FCP-ST']]

    # 4. Processamento dos blocos de Saída e Entrada
    df_s = df_validas[df_validas['SENTIDO'] == 'SAÍDA']
    df_e = df_validas[df_validas['SENTIDO'] == 'ENTRADA']

    res_s = preparar_tabela_completa(df_s)
    res_e = preparar_tabela_completa(df_e)

    # 5. Cálculo da Tabela de Saldo Líquido (Saída - Entrada)
    res_saldo = pd.DataFrame({'ESTADO (UF)': UFS_BRASIL})
    res_saldo['IE SUBSTITUTO'] = res_s['IE_SUBST']
    res_saldo['ST LÍQUIDO'] = res_s['VAL-ICMS-ST'] - res_e['VAL-ICMS-ST']
    res_saldo['DIFAL LÍQUIDO'] = res_s['VAL-DIFAL'] - res_e['VAL-DIFAL']
    res_saldo['FCP LÍQUIDO'] = res_s['VAL-FCP'] - res_e['VAL-FCP']
    res_saldo['FCP-ST LÍQUIDO'] = res_s['VAL-FCP-ST'] - res_e['VAL-FCP-ST']

    # Renomeando colunas para gravação lado a lado
    res_s.columns = ['ESTADO (UF)', 'IE SUBSTITUTO', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
    res_e.columns = ['ESTADO (UF) ', 'IE SUBSTITUTO ', 'ST TOTAL ', 'DIFAL TOTAL ', 'FCP TOTAL ', 'FCP-ST TOTAL ']

    # 6. Gravação Física no Excel (Lado a Lado)
    workbook = writer.book
    worksheet = workbook.add_worksheet('DIFAL_ST_FECP')
    writer.sheets['DIFAL_ST_FECP'] = worksheet

    # Formatação de Títulos
    title_fmt = workbook.add_format({'bold': True, 'font_color': '#FF6F00', 'font_size': 12})
    total_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'num_format': '#,##0.00'})

    # Escrevendo Títulos das Seções
    worksheet.write(0, 0, "1. SAÍDAS (DÉBITO)", title_fmt)
    worksheet.write(0, 8, "2. ENTRADAS (CRÉDITO)", title_fmt)
    worksheet.write(0, 16, "3. SALDO LÍQUIDO (RECOLHER)", title_fmt)

    # Gravando as Tabelas
    res_s.to_excel(writer, sheet_name='DIFAL_ST_FECP', startrow=2, startcol=0, index=False)
    res_e.to_excel(writer, sheet_name='DIFAL_ST_FECP', startrow=2, startcol=8, index=False)
    res_saldo.to_excel(writer, sheet_name='DIFAL_ST_FECP', startrow=2, startcol=16, index=False)

    # Adicionando Linha de Totais no Rodapé (Opcional - Linha 30 do Excel)
    for col_set in [0, 8, 16]:
        for i, col_name in enumerate(['ST', 'DIFAL', 'FCP', 'FCP-ST']):
            col_idx = col_set + 2 + i
            # Soma das linhas 4 a 30 (27 estados)
            worksheet.write(30, col_idx, f'=SUM({chr(65+col_idx)}4:{chr(65+col_idx)}30)', total_fmt)
        worksheet.write(30, col_set, "TOTAL GERAL", total_fmt)
