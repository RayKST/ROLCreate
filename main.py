import pandas as pd
from openpyxl import Workbook
import string

alfabeto = []
for letra in string.ascii_uppercase:
    alfabeto.append(letra)

excel = pd.ExcelFile("input/input.xlsx")
df = excel.parse("Planilha1")
colunaBase = input("Insira o nome da coluna para ser ordenada: ")
df = df.sort_values(by=colunaBase)


workbook = Workbook()
sheet = workbook.active

indexNumeros = 0
for index, valor in enumerate(df[colunaBase].to_list()):
    if index%5 == 0:
        indexNumeros += 1
        indexLetra = 0
    
    cell = alfabeto[indexLetra] + str(indexNumeros)
    sheet[cell] = valor
    indexLetra += 1

workbook.save(filename="output/"+ input("Insira o nome da planilha a ser criada: ") +".xlsx")
print("Planilha criada")