import pandas as pd
import win32com.client as win32

#importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#faturamento por loja
# primeiro criamos o filtro das colunas que queremos olhar 'tabela_vendas[['ID Loja', 'Valor Final']]'
# Depois agrupamos os resultados pela coluna ID Loja com o 'groupby' assim o resultado aparece só uma vez.
# Por fim somamos as outras colunas, no caso 'Valor Final' com o comando .sum().
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)

#quantidade de produtos vendidos por loja
quantidades = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidades)

print('-' * 50)

#ticket médio por produto em cada loja
#quando fazemos operações entre colunas o resultado não é dado como uma tabela
#Por isso colocamos o código entre parenteses e colocamos '.to_frame()' no final do código, convertendo o resultado para uma tabela
ticket_medio = (faturamento['Valor Final'] / quantidades['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

