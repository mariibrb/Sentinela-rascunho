import pandas as pd

def gerar_resumo_uf(df, writer):
    # FILTRO: Apenas Autorizadas
    df_aut = df[df['Situação Nota'].str.upper().str.contains('AUTORIZADA', na=False)]
    if not df_aut.empty:
        res = df_aut.groupby(['UF_DEST', 'IE_SUBST']).agg({
            'VAL-ICMS-ST': 'sum', 'VAL-DIFAL': 'sum', 'VAL-FCP': 'sum', 'VAL-FCP-ST': 'sum'
        }).reset_index()
        res.columns = ['UF', 'IE SUBST.', 'ST TOTAL', 'DIFAL TOTAL', 'FCP TOTAL', 'FCP-ST TOTAL']
        res.to_excel(writer, sheet_name='DIFAL_ST_FECP', index=False)
