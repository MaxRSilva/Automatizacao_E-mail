import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel("Vendas.xlsx")
pd.set_option("display.max_columns", None)


faturamento = (tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum())
print(faturamento)

print ("-" * 50)

faturamento_de_produtos = (tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum())
print(faturamento_de_produtos)

print ("-" * 50)

ticket_medio = (faturamento["Valor Final"] / faturamento_de_produtos ["Quantidade"]).to_frame()
print(ticket_medio)