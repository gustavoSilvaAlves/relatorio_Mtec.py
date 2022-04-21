import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_mtec = pd.read_excel('mtec.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_mtec)

#Planilha toda

planilha = tabela_mtec[['Pv', 'Data', 'Tecnico','Produto','Quantidade','Valor Unitário','Valor Final','Observação','Função']]
print(planilha)

# faturamento por loja
faturamento = tabela_mtec[['Tecnico', 'Valor Final']].groupby('Tecnico').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_mtec[['Tecnico', 'Quantidade']].groupby('Tecnico').sum()
print(quantidade)

print('-' * 50)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gu.silvaalves@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Claúdio,</p>
<p>Segue o Relatório de cada colaborador.</p>
<p> Planilha total:</p>
{planilha.to_html(formatters={'Valor Final':'R${:,.2F}'.format})}
<p>Faturamento do dia:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2F}'.format})}
<p>Máquinas finalizadas:</p>
{quantidade.to_html()}
<p>Qualquer dúvida estou à disposição.</p>
<p>Att.</p>
<p>Gustavo</p>
'''

mail.Send()

print('Email Enviado')
