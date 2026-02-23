import os
import tkinter as tk
from tkinter import filedialog, messagebox
from reader import processarTXT
from xlGenerator import gerarExcel

class RelatorioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatorio Contábil")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        self.caminhoTXT = ""

        # Título
        titulo = tk.Label(root, text="Gerador de Relatório Contábil", font=("Arial", 16, "bold"))
        titulo.pack(pady=10)

        # Botão selecionar
        btnSelec = tk.Button(root, text="Selecionar Arquivo TXT", command=self.selecArq, width=30)
        btnSelec.pack(pady=10)

        # Label caminho
        self.LabelCaminho = tk.Label(root, text="Nenhum arquivo selecionado", wraplength=450)
        self.LabelCaminho.pack(pady=5)

        # Botão gerar
        btnGerar = tk.Button(root, text="Gerar Relatório", command=self.gerarRelatorio, width=30, bg="#1F4E79",
                              fg="white")
        btnGerar.pack(pady=20)

        # Status
        self.label_status = tk.Label(root, text="")
        self.label_status.pack(pady=5)

    def selecArq(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("Arquivos TXT", "*.txt")]
        )

        if caminho:
            self.caminhoTXT = caminho
            self.LabelCaminho.config(text=caminho)

    def gerarRelatorio(self):
        if not self.caminhoTXT:
            messagebox.showwarning("Aviso", "Selecione um arquivo TXT primeiro.")
            return

        try:
            grupos = processarTXT(self.caminhoTXT)

            pasta = os.path.dirname(self.caminhoTXT)
            nomeBase = os.path.splitext(os.path.basename(self.caminhoTXT))[0]

            caminhoExcel = os.path.join(pasta, f"{nomeBase}.xlsx")

            gerarExcel(grupos, caminhoExcel)

            self.label_status.config(text="Relatório gerado com sucesso!", fg="green",)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {caminhoExcel}")
        except Exception as e:
            self.label_status.config(text="Erro ao gerar o relatório.", fg="red")
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__": #executa o programa
    root = tk.Tk()
    app = RelatorioApp(root)
    root.mainloop()
