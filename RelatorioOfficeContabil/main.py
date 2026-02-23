import os
from tkinter import Tk, filedialog
from reader import processarTXT
from xlGenerator import gerarExcel

def selecArq():
    Tk().withdraw() # Esconde a janela principal

    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos TXT", "*.txt")]
    )

    return caminho

def main():
    caminhoTXT = selecArq()

    if not caminhoTXT: #verifica a existência do arquivo
        print("Nenhum arquivo selecionado!")
        return

    grupos = processarTXT(caminhoTXT) #processa o TXT

    # Define o nome da planilha
    pasta = os.path.dirname(caminhoTXT)
    nomeBase = os.path.splitext(os.path.basename(caminhoTXT))[0]
    caminhoExcel = os.path.join(pasta, f"{nomeBase}.xlsx")

    # Gera o relatorio no local do arquivo
    gerarExcel(grupos, caminhoExcel)
    print("Relatorio Gerado!")
    print(f"Arquivo salvo em: {caminhoExcel}")
    print(grupos)

if __name__ == "__main__": #executa o programa
    main()
