import pandas as pd

excel = pd.ExcelFile("input/input.xlsx")
df = excel.parse("Planilha1")
df = df.sort_values(by="Header Row")
writer = pd.ExcelWriter('output/output.xlsx')
df.to_excel(writer,sheet_name='Planilha1',columns=["Header Row"],index=False)
writer.close()