import pandas as pd
import smtplib
import email.message


# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja

# tabela_vendas[['ID Loja', 'Valor Final']]
# tabela_vendas.groupby('ID Loja').sum()
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)

# Enviar email com relatorio

email_content = f''' 
<html>

<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Qantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Medio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposição.</p>

<p>At..</p>
<p>Romulo Conceição</p>

</html>
 '''
msg = email.message.Message()
msg['Subject'] = 'Tabelas'

msg['From'] = '@outlook.com' # - Coloque seu e-mail
msg['To'] = '@outlook.com' # - E-mails de destinatarios
password = '*******' # - Senha do seu e-mail
msg.add_header('Content-Type', 'text/html')
msg.set_payload(email_content)

mail = smtplib.SMTP('smtp-mail.outlook.com', 587) # ('smtp.gmail.com', 587) - Gmail
mail.starttls()

mail.login(msg['From'], password)

mail.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
mail.quit()

print('E-mail enviado')

# Se o seu Sistema Operacional for Windows

# import win32com.client as win32
#
# outlook = win32.Dispatch('outlook.application')
# mail = outlook.CreateItem(0)
# mail.To = 'pythonimpressionador@gmail.com'
# mail.Subject = 'Relatório de Vendas por Loja'
# mail.HTMLBody = f'''
#
# <p>Prezados,</p>
#
# <p>Segue o Relatório de Vendas por cada Loja.</p>
#
# <p>Faturamento:</p>
# {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
#
# <p>Qantidade Vendida:</p>
# {quantidade.to_html()}
#
# <p>Ticket Medio dos Produtos em cada Loja:</p>
# {ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}
#
# <p>Qualquer duvida estou a disposição.</p>
#
# <p>At..</p>
# <p>Romulo Conceição</p>
#
#  '''