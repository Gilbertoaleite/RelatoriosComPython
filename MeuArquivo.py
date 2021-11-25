import win32com.client as win32
import pandas as pd

#import a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")  #nome da tabela que vai importar

#visualizar a base de dados
pd.set_option("display.max_columns", None) #configurar a visualizacao (opcao, valor)
#print(tabela_vendas)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)
#quantidade de produtos vendidos por loja
Quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(Quantidade)
print('-' * 50)

#ticket medio por produto em casa loja
ticket_medio = (faturamento['Valor Final'] / Quantidade['Quantidade']).to_frame()
print(ticket_medio)

#envia email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gilbertoaleite@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{Quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Gilberto A Leite</p>
'''
mail.Send()
print('Email Enviado')


