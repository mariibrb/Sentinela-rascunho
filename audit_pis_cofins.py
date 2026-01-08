import pandas as pd

def processar_pc(df, writer, cod_cliente):
    df_pc = df.copy()
    df_pc['Diagnóstico P/C'] = "✅ Analisado"
    tags = [c for c in df_pc.columns if c not in ['Situação Nota', 'Diagnóstico P/C']]
    df_pc[tags + ['Situação Nota', 'Diagnóstico P/C']].to_excel(writer, sheet_name='PIS_COFINS_AUDIT', index=False)
