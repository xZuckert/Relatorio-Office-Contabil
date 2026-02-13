import re
from collections import defaultdict

cfopValid = {"5.102", "6.102", "5.405", "6.405"} #CFOPs de saida que devem ser leançados

#compila a linha em tipo, serie, numero, dia, uf, valor e CFOP, ficando assim: |NFCE|1|000123|05|SP| 1.234,56 | ... |5.102|
padrao = re.compile(
    r'\|(?P<tipo>NFCE|NFE)\s*\|'
    r'(?P<serie>\d+)\s*\|'
    r'(?P<numero>\d+)\|'
    r'(?P<dia>\d{2})\s*\|'
    r'(?P<uf>[A-Z]{2})\|'
    r'\s*(?P<valor>[\d\.,]+)\|.*\|'
    r'\s*(?P<cfop>\d\.\d{3})'
)

def converterValor(valorString):
    return float(valorString.replace('.', '').replace(',', '.')) #converte "1.234,56" em 1234,56

def processarTXT(caminhoArquivo):
    grupos = defaultdict(lambda: {
        "min": None,
        "max": None,
        "total": 0.0
    })
    with open(caminhoArquivo, "r", encoding="latin-1", errors="ignore") as f:
        for linha in f:
            match = padrao.search(linha)
            if not match:
                continue

            cfop = match.group("cfop") #filtra os CFOPs de venda de mercadorias para extrair dia, serie, numero e valor da nota
            if cfop not in cfopValid:
                continue
            if match:
                dia = int(match.group("dia"))
                serie = match.group("serie")
                numero = int(match.group("numero").lstrip("0") or "0")
                valor = converterValor(match.group("valor"))

            chave = (dia, serie) #agrupa por dia e serie

            #faz o calculo dos numeros das notas para o relatorio diario
            if grupos[chave]["min"] is None:
                grupos[chave]["min"] = numero
                grupos[chave]["max"] = numero
            else:
                grupos[chave]["min"] = min(grupos[chave]["min"], numero)
                grupos[chave]["max"] = max(grupos[chave]["max"], numero)

            grupos[chave]["total"] += valor
    return grupos
