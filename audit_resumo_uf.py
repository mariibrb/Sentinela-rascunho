import pandas as pd

def gerar_resumo_uf(df, writer):
    """
    Gera a aba DIFAL_ST_FECP com somatória por UF e IE Substituto.
    Ignora notas canceladas ou não encontradas.
    """
    # Filtro rigoroso: Apenas Notas Autorizadas
    df_aut = df[df['Situação Nota'].str.upper().str.contains('AUTORIZADA', na=False)].copy()
    
    if not df_aut.empty:
        # Agrupamento por Estado e Inscrição Estadual de Substituto
        res = df_aut.groupby(['UF_DEST', 'IE_SUBST']).agg({
            'VAL-ICMS-ST': 'sum',
            'VAL-DIFAL': 'sum',
            'VAL-FCP': 'sum',
            'VAL-FCP-ST': 'sum'
        }).reset_index()
        
        # Renomeando colunas para o padrão aprovado
        res.columns = ['ESTADO (UF)', 'IE SUBSTITUTO', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
        
        # Grava no Excel
        res.to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False)
    else:
        # Se não houver notas autorizadas, cria a aba com aviso para não quebrar o Excel
        pd.DataFrame([["Aviso:", "Nenhuma nota AUTORIZADA encontrada para somatória."]]).to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False, header=False)
