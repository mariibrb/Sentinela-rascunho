import pandas as pd

def gerar_abas_gerenciais(writer, ge, gs):
    for f_obj, s_name in [(ge, 'GERENCIAL_ENTRADA'), (gs, 'GERENCIAL_SAIDA')]:
        if f_obj:
            try:
                f_obj.seek(0)
                if f_obj.name.endswith('.xlsx'):
                    df = pd.read_excel(f_obj)
                else:
                    # sep=None detecta automaticamente se é vírgula ou ponto-e-vírgula
                    # on_bad_lines='skip' ignora a linha 15 que está com erro no seu arquivo
                    df = pd.read_csv(f_obj, sep=None, engine='python', on_bad_lines='skip', encoding='utf-8')
                
                df.to_excel(writer, sheet_name=s_name, index=False)
            except Exception as e:
                # Se ainda der erro, cria aba com aviso em vez de travar o app
                pd.DataFrame([["Erro ao ler arquivo", str(e)]]).to_excel(writer, sheet_name=s_name, index=False, header=False)
