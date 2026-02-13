import os
from reader import processarTXT

def main():
    caminhoTXT = r"RelatorioOfficeContabil\LivrosFiscais\livro.txt" #nome do arquivo (Teste)

    if not os.path.exists(caminhoTXT): #verifica a existência do arquivo
        print("Arquivo não encontrado!")
        print("Diretório atual:", os.getcwd())
        return

    grupos = processarTXT(caminhoTXT)
    print(grupos)

if __name__ == "__main__": #executa o programa
    main()
