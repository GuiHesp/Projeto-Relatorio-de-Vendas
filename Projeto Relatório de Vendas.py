#!/usr/bin/env python
# coding: utf-8

# # Desafio:
# 
# Você faz parte da equipe de Analytics de uma grande marca de vestuário com mais de 25 lojas espalhadas em Shoppings de todo o Brasil.
# 
# Toda semana você precisa enviar para a diretoria um ranking atualizado com as 25 lojas contendo 3 informações:
# - Faturamento de cada Loja
# - Quantidade de Produtos Vendidos de cada Loja
# - Ticket Médio dos Produto de cada Loja
# 
# Além disso, cada loja tem 1 gerente que precisa receber o resumo das informações da loja dele. Por isso, cada gerente deve receber no e-mail:
# - Faturamento da sua loja
# - Quantidade de Produtos Vendidos da sua loja
# - Ticket Médio dos Produto da sua Loja
# 
# Esse relatório é sempre enviado como um resumo de todos os dados disponíveis no ano.

# ### Passo 1 - Importando a Base de Dados + Passo 2 - Visualizando os Dados

# In[ ]:


import pandas as pd

df = pd.read_excel(r'C:\Users\Gui\CURSO PYTHON\29- Intensivão python\Aula 01\Vendas.xlsx')

display(df)


# ### Passo 3.1 - Calculando o Faturamento por Loja

# In[ ]:


faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
faturamento = faturamento.sort_values(by='Valor Final', ascending=False)
display(faturamento)


# ### Passo 3.2 - Calculando a Quantidade Vendida por Loja

# In[ ]:


quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='ID Loja', ascending=False)
display(quantidade)


# ### Passo 3.3 - Calculando o Ticket Médio dos Produtos por Loja

# In[ ]:


ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Medio', ascending=False)
display(ticket_medio)


# ### Criando a função de enviar e-mail

# In[ ]:


# função enviar_email
import smtplib
import email.message

def enviar_email(resumo_loja, loja):

  server = smtplib.SMTP('smtp.gmail.com:587')  
  email_content = f'''
  <p>Olá, aqui vai o resumo das lojas:</p>
  {resumo_loja}
  <p>Abraços</p>'''
  
  msg = email.message.Message()
  msg['Subject'] = f'Loja: {loja}'
  
  
  msg['From'] = 'email'
  msg['To'] = 'email'
  password = 'passwor'
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(email_content)
  
  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))


# In[ ]:


enviar_email('Diretoria', faturamento)


# ### Calculando Indicadores por Loja + Enviar E-mail para todas as lojas

# In[ ]:


# email para diretoria

tabela_diretoria = faturamento.join(quantidade).join(ticket_medio)
enviar_email(tabela_diretoria, 'Todas as Lojas')


# In[ ]:


lojas = df['ID Loja'].unique()

for loja in lojas:
  tabela_loja = df.loc[df['ID Loja'] == loja, ['ID Loja', 'Quantidade', 'Valor Final']]
  resumo_loja = tabela_loja.groupby('ID Loja').sum()
  resumo_loja['Ticket Médio'] = resumo_loja['Valor Final'] / resumo_loja['Quantidade']
  enviar_email(resumo_loja, loja)


# ### Enviar e-mail para a diretoria
