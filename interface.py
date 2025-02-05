import customtkinter as ctk
from tkinter import filedialog, messagebox
from relatorio import gerar_relatorio

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def iniciar_interface():
    root = ctk.CTk()
    root.title("Gerador de Relatórios de Gastos")
    root.geometry("500x400")

    def selecionar_arquivo():
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminho:
            entrada_arquivo.delete(0, "end")
            entrada_arquivo.insert(0, caminho)

    def selecionar_pasta():
        pasta = filedialog.askdirectory()
        if pasta:
            entrada_pasta.delete(0, "end")
            entrada_pasta.insert(0, pasta)

    def chamar_gerar_relatorio():
        arquivo_excel = entrada_arquivo.get()
        pasta_destino = entrada_pasta.get()
        gerar_relatorio(arquivo_excel, pasta_destino)

    frame = ctk.CTkFrame(root)
    frame.pack(pady=20, padx=20, fill="both", expand=True)

    ctk.CTkLabel(frame, text="Selecione o arquivo Excel:").pack()
    entrada_arquivo = ctk.CTkEntry(frame, width=300)
    entrada_arquivo.pack(pady=5)
    ctk.CTkButton(frame, text="Procurar", command=selecionar_arquivo).pack()

    ctk.CTkLabel(frame, text="Selecione a pasta de destino:").pack(pady=(10, 0))
    entrada_pasta = ctk.CTkEntry(frame, width=300)
    entrada_pasta.pack(pady=5)
    ctk.CTkButton(frame, text="Procurar", command=selecionar_pasta).pack()

    ctk.CTkButton(frame, text="Gerar Relatório", fg_color="green", hover_color="darkgreen", command=chamar_gerar_relatorio).pack(pady=20)

    root.mainloop()
