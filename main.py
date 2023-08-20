import pandas as pd

excel = pd.ExcelFile("input/input.xlsx")
df = excel.parse("Planilha1")
colunaBase = 'Header Row'#input("Insira o nome da coluna para ser ordenada: ")
df = df.sort_values(by=colunaBase)
writer = pd.ExcelWriter('output/output.xlsx')
df.to_excel(writer,sheet_name='Planilha1',columns=[colunaBase],index=False)
writer.close()

print("Planilha criada")