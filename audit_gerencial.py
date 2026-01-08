import pandas as pd

def gerar_abas_gerenciais(writer, ge, gs):
    for f_obj, s_name in [(ge, 'GERENCIAL_ENTRADA'), (gs, 'GERENCIAL_SAIDA')]:
        if f_obj:
            f_obj.seek(0)
            df = pd.read_excel(f_obj) if f_obj.name.endswith('.xlsx') else pd.read_csv(f_obj)
            df.to_excel(writer, sheet_name=s_name, index=False)
