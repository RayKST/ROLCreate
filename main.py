import pandas as pd
from openpyxl import Workbook
import string
import math

alfabeto = []
for letra in string.ascii_uppercase:
    alfabeto.append(letra)

excel = pd.ExcelFile("input/input.xlsx")
df = excel.parse("Planilha1")
colunaBase = input("Insira o nome da coluna para ser ordenada: ")
df = df.sort_values(by=colunaBase)


workbook = Workbook()
sheet = workbook.active
dfList = df[colunaBase].to_list()
indexNumeros = 0
for index, valor in enumerate(dfList):
    if index%5 == 0:
        indexNumeros += 1
        indexLetra = 0
    
    cell = alfabeto[indexLetra] + str(indexNumeros)
    sheet[cell] = valor
    indexLetra += 1

amplitudeAmostra = dfList[-1] - dfList[0]
numeroClasses = round((1 + math.log2(dfList.__len__()))+0.5)
amplitudeClasse = round((amplitudeAmostra/numeroClasses)+0.5)
sheet['G1'] = 'Amplitude da amostra - AA:'
sheet['G2'] = 'Numero de classes - K:'
sheet['G3'] = 'Amplitude da classe - H:'
sheet['H1'] = amplitudeAmostra
sheet['H2'] = numeroClasses
sheet['H3'] = amplitudeClasse

def verificaIntervalo(intervaloMenor, intervaloMaior):
    contador = 0
    for value in dfList:
        if intervaloMenor <= value < intervaloMaior:
            contador += 1
    return contador

inicioTabela = dfList[0]
inicioIntervalo = dfList[0]
for index in range(numeroClasses):
    if index == 0:
        sheet[alfabeto[9]+ str(index + 1)] = f'{inicioTabela} <--- {inicioTabela + amplitudeClasse}'
        frequencia = verificaIntervalo(inicioTabela, inicioTabela + amplitudeClasse)
        sheet[alfabeto[10] + str(index + 1)] = frequencia
    else:
          sheet[alfabeto[9]+ str(index + 1)] = f'{inicioIntervalo} <--- {inicioIntervalo + amplitudeClasse}'  
          frequencia = verificaIntervalo(inicioIntervalo, inicioIntervalo + amplitudeClasse)
          sheet[alfabeto[10] + str(index + 1)] = frequencia
   
    inicioIntervalo += amplitudeClasse

workbook.save(filename="output/"+ input("Insira o nome da planilha a ser criada: ") +".xlsx")
print("Planilha criada")