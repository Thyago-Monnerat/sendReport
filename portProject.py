import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Verificar a base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Calcular o ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# Enviar um relatório por email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Email destino'
mail.Subject = 'Relatório de vendas'
mail.HTMLBody = f'''
<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>

<p>Att.,Thyago</p>

'''
mail.Send()