import glob
import pandas as pd

def arquivos():
    arquivos = glob.glob('*.xls')
    if arquivos:
        arquivo_original = arquivos[0]
    else:
        raise FileNotFoundError("Nenhum arquivo .xls encontrado.")

    df = pd.read_excel(arquivo_original, engine="xlrd")
    novo_nome = "Relatório.xlsx"
    df.to_excel(novo_nome, index=False)