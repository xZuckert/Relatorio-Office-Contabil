import time
import pandas as pd
from pywinauto import Application
import os

delay = 0.25
PARAR_AUTOMACAO = False

# Pega o mes e o ano do documento TXT
def extrairMesAno(caminhoTXT):
    nome = os.path.basename(caminhoTXT)
    base = os.path.splitext(nome)[0]

    mes = base[1:3]
    ano = base[3:7]

    return mes, ano

# Pega o codigo da empresa da pasta que esta o TXT
def extrairCodEmpresa(caminhoTXT):
    pasta = os.path.dirname(caminhoTXT)
    return os.path.basename(pasta)

# Conecta ao office contábil ja aberto
def conectarOffice():
    app = Application(backend="uia").connect(title_re="^Office Contábil")
    janela = app.window(title_re="^Office Contábil")
    janela.set_focus()
    return janela

# Lê o relatório excel
def lerRelatorio(caminhoExcel):
    df = pd.read_excel(caminhoExcel)
    df = df[df["Dia"].notna()]
    print(df)
    print("Quantidade de linhas:", len(df))
    return df

def ativarEmpresa(janela, codigoEmpresa, mes, ano):
    # ESC 5 vezes para fechar qualquer janela aberta no office
    for i in range(5):
        janela.type_keys("{ESC}")
        time.sleep(delay)

    # Abrir Ativação de Empresa
    janela.type_keys("E")
    time.sleep(1)

    # Codigo da Empresa
    janela.type_keys(codigoEmpresa)
    janela.type_keys("{ENTER}")

    # Mês/Ano
    janela.type_keys(f"{mes}{ano}")
    janela.type_keys("{ENTER}")

    # Pular Digitador e Ativar
    janela.type_keys("{TAB}")
    janela.type_keys("{ENTER}")

    time.sleep(4)

    # Ir para Digitação
    janela.type_keys("D")
    time.sleep(1)

def lancarProvisao(janela, dia, valor, numero):
    janela.type_keys("N") # Inicia novo lancamento
    time.sleep(delay)

    janela.type_keys("{TAB}") # Pula Lcto

    janela.type_keys(str(dia)) # Dia provisão
    janela.type_keys("{ENTER}")

    janela.type_keys("{TAB}") # Pula Hist.

    janela.type_keys("5") # Débito provisão
    janela.type_keys("{ENTER}")

    janela.type_keys("104") # Crédito provisão
    janela.type_keys("{ENTER}")

    janela.type_keys(str(valor)) #Valor provisão
    janela.type_keys("{ENTER}")

    janela.type_keys(str(numero)) #Número das notas
    janela.type_keys("{PGDN}") # Gravar

    time.sleep(delay)

def lancarPagamento(janela, dia, valor, numero):
    janela.type_keys("N")  # Inicia novo lancamento
    time.sleep(delay)

    janela.type_keys("{TAB}")  # Pula Lcto

    janela.type_keys(str(dia))  # Dia pagamento
    janela.type_keys("{ENTER}")

    janela.type_keys("{TAB}")  # Pula Hist.

    janela.type_keys("1")  # Débito pagamento
    janela.type_keys("{ENTER}")

    janela.type_keys("5")  # Crédito pagamento
    janela.type_keys("{ENTER}")

    janela.type_keys(str(valor))  # Valor pagamento
    janela.type_keys("{ENTER}")

    janela.type_keys(str(numero))  # Número das notas
    janela.type_keys("{PGDN}")  # Gravar

    time.sleep(delay)

def executarAutomacao(caminhoTXT, caminhoExcel):
    global PARAR_AUTOMACAO
    PARAR_AUTOMACAO = False

    mes, ano = extrairMesAno(caminhoTXT)
    codigoEmpresa = extrairCodEmpresa(caminhoTXT)

    df = lerRelatorio(caminhoExcel)

    janela = conectarOffice()

    ativarEmpresa(janela, codigoEmpresa, mes, ano)

    for i, row in df.iterrows():
        if PARAR_AUTOMACAO:
            print("Automação interrompida")
            break

        janela.set_focus()
        time.sleep(0.1)
        dia = int(row["Dia"])
        valor = float(row["valor contabil (R$)"])
        numero = row["número"]

        lancarProvisao(janela, dia, valor, numero)
        if PARAR_AUTOMACAO:
            print("Automação interrompida")
            break

        janela.set_focus()
        time.sleep(0.1)
        lancarPagamento(janela, dia, valor, numero)

def pararAutomacao():
    global PARAR_AUTOMACAO
    PARAR_AUTOMACAO = True