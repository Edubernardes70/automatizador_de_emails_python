import openpyxl as openpyxl
import pandas as pd
#import pywin32 as pywin32
#precisa instalar o openpyxl
#incorporando a base de dados
tabela_vendas=pd.read_excel('Vendas.xlsx')
#visaualizar a base de dados
pd.set_option('display.max_columns', None)#aqui vamos selecionar para mostrar o máximo de colunas
#none para não botar limites no que vai aparecer
#print(tabela_vendas[['ID Loja', 'Valor Final']])posso usar isto para filtrar

#FATURAMENTO POR LOJA
faturamento=(tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum())
#tabela_vendas [Id loja', valor final] pega os itens que irão aparecer
#groupby vai juntar as lojas com o mesmo nome para não repetir
#sum= realizada a soma dos elementos de cada loja. Veja que depois do groupby de apresentar que quero ver o ID loja o restante eu disse que era para somar
print(faturamento)

print('-'*50)
#QUANTIDADE DE PRODUTOS VENDIDOS
quantidade=tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
#Aqui fazemos igual ao faturamento da loja, porém eu indico a quantidade ao invés do valor final

print(quantidade)
print('-'*50)

#TICKET MÉDIO
ticket=(faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
#Aqui estou pegando a coluna valor final de faturamento para pegar os itens apenas desta coluna, por isso uso apenas 1 []
#Depos faço a divisão do valor final pela tabela quantidade
#entre parenteses com .to_frame() transforma os dados em numeros. Isso quando vamos dividir uma tabela por outra.
ticket=ticket.rename(columns={0: 'Ticket Médio'})#Aqui mudamos o ) que aparecia como titulo para ticket médio
print(ticket)
print('-'*50)

#ENVIAR E-MAIL COM RELATÓRIO
#precisa instalar o pywin32

# In[ ]:


import smtplib
import email.message
def enviar_email():
    corpo_email = f"""
    <p>Prezados, </p>

    <p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Eduardo Bernardes</p>

     """

    msg = email.message.Message()
    msg['Subject'] = "Relatório de vendas"
    msg['From'] = 'profeduardobernardes@gmail.com'
    msg['To'] = 'profeduardobernardes@gmail.com'
    password = 'bwjqqfmwterkamxz'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


# In[ ]:


enviar_email()