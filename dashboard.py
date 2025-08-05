import customtkinter as ctk
from tkinter import messagebox
from arquivo import arquivos
from peracio import gerar_m_relatorio, gerar_m_pedido, gerar_m_cod_errado, gerar_m_zero_estoque, gerar_m_finalizado
from nartic import gerar_n_relatorio, gerar_n_pedido, gerar_n_cod_errado, gerar_n_zero_estoque, gerar_n_finalizado

def iniciar_programa():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.geometry("550x510")
    app.title("Transferência de Mercadorias")

    # Frame de cabeçalho (título visual)
    header = ctk.CTkFrame(app)
    header.pack(fill="x", padx=10, pady=10)
    # Label à esquerda
    titulo_esquerda = ctk.CTkLabel(header, text="Transferência - Moeda | Nartic", font=("Arial", 14, "bold"))
    titulo_esquerda.pack(side="left", padx=10)
    # Label à direita
    titulo_direita = ctk.CTkLabel(header, text="Desenvolvido por Lauro Júnior", font=("Arial", 10), text_color="gray")
    titulo_direita.pack(side="right", padx=10)



    janela = ctk.CTkFrame(app)
    janela.pack(fill="both", expand=True, padx=20, pady=20)

    # Criando os botões como itens do menu
    arquivo = ctk.CTkButton(janela, text=".XLS  p/  .XLSX", width=200, height=50, font=("Arial", 12), command=lambda:arquivos())
    arquivo.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    m_final = ctk.CTkButton(janela, text="MOEDA FINALIZADO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_m_finalizado())
    m_final.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
    m_relatorio = ctk.CTkButton(janela, text="MOEDA RELATÓRIO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_m_relatorio())
    m_relatorio.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
    m_pedido = ctk.CTkButton(janela, text="MOEDA PEDIDO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_m_pedido())
    m_pedido.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
    m_zero_etq = ctk.CTkButton(janela, text="MOEDA ZERO ESTOQUE", width=200, height=50, font=("Arial", 12), command=lambda:gerar_m_zero_estoque())
    m_zero_etq.grid(row=4, column=0, padx=10, pady=10, sticky="ew")
    m_cod_errado = ctk.CTkButton(janela, text="TRANSF. CÓDIGO ERRADO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_m_cod_errado())
    m_cod_errado.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

    n_final = ctk.CTkButton(janela, text="NARTIC FINALIZADO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_n_finalizado())
    n_final.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
    n_relatorio = ctk.CTkButton(janela, text="NARTIC RELATÓRIO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_n_relatorio())
    n_relatorio.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
    n_pedido = ctk.CTkButton(janela, text="NARTIC PEDIDO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_n_pedido())
    n_pedido.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
    n_zero_etq = ctk.CTkButton(janela, text="NARTIC ZERO ESTOQUE", width=200, height=50, font=("Arial", 12), command=lambda:gerar_n_zero_estoque())
    n_zero_etq.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
    n_cod_errado = ctk.CTkButton(janela, text="TRANSF. CÓDIGO ERRADO", width=200, height=50, font=("Arial", 12), command=lambda:gerar_n_cod_errado())
    n_cod_errado.grid(row=5, column=1, padx=10, pady=10, sticky="ew")
    
    janela.grid_columnconfigure(0, weight=1)
    janela.grid_columnconfigure(1, weight=1)

    app.mainloop()

    