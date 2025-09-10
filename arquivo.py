import glob
import pandas as pd
import sys
import os
import traceback
from tkinter import messagebox

def resource_path(relative_path):
    """Pega o caminho correto do recurso, seja em dev ou em exe."""
    try:
        # PyInstaller cria uma pasta temporária _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)  


def tratar_erro(e, nome_funcao=""):
    # e → é a exceção que aconteceu.
    # nome_funcao → opcional, só pra aparecer no título do popup qual função deu problema.
    """Mostra erro em popup sem salvar log"""

    # Resumo do erro
    # type(e).__name__ → pega só o nome do erro (ValueError, ZeroDivisionError, etc).
    # e → traz a mensagem do erro.
    erro_resumido = f"{type(e).__name__}: {e}"

    # Detalhes completos (traceback formatado)
    # Isso cria o erro completo com traceback (ou seja, aquela árvore de chamadas que o Python mostra no console).
    erro_detalhado = "".join(traceback.format_exception(type(e), e, e.__traceback__))

    # Título do popup
    # Se você passou o nome da função, o título do popup fica, por exemplo:
    # Erro em gerar_m_finalizado
    # Se não passou nada, fica só Erro.
    titulo = f"Erro em {nome_funcao}" if nome_funcao else "Erro"

    # Mostra o popup com erro resumido
    # Um só com o resumo (mais amigável pro usuário).
    # Outro com os detalhes técnicos completos (mais útil pra você, dev).
    messagebox.showerror(titulo, erro_resumido)
    messagebox.showerror(titulo, erro_detalhado)

    # (Opcional) → Se você quiser também exibir detalhes completos no terminal:
    print(erro_detalhado)


def captura_erros(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            tratar_erro(e, func.__name__)
    return wrapper

    # captura_erros(func) → recebe a função que você decorou.
    # wrapper → é uma função interna que vai substituir a original.
    # *args, **kwargs → garante que qualquer argumento passado pra função original ainda funcione.
    # Dentro do try → executa a função normalmente.
    # Se der erro → chama o tratar_erro, já passando o nome da função (func.__name__).

    # Quando você chamar gerar_m_finalizado(), quem de fato roda é o wrapper.
    # Se der erro → tratar_erro é chamado → abre popup pro usuário.

    # Resumindo
    # tratar_erro = centraliza exibição do erro.
    # captura_erros = evita duplicar try/except em todas as funções, automatizando isso.
    # @captura_erros = um “atalho mágico” pra englobar tudo.


@captura_erros
def arquivos():
    arquivos = glob.glob('*.xls')
    if arquivos:
        arquivo_original = arquivos[0]
    else:
        raise FileNotFoundError("Nenhum arquivo .xls encontrado.")

    df = pd.read_excel(arquivo_original, engine="xlrd")
    novo_nome = "Relatório.xlsx"
    df.to_excel(novo_nome, index=False)