#!/usr/bin/env python
# coding: utf-8
# 1 Passo - Criar um Navegador 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time



# 2 Passo - Importar/visualizar base de dados
import pandas as pd 
tabela_produtos = pd.read_excel('buscas.xlsx')
display(tabela_produtos)


#funcao google
def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    # entrar no google
    nav = webdriver.Chrome()
    nav.get("https://www.google.com/")
    
    # tratar os valores que vieram da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    

    # pesquisar o nome do produto no google
    nav.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(produto, Keys.ENTER)

    time.sleep(5)
    # clicar na aba shopping
    nav.find_element(By.XPATH, '//*[@id="hdtb-msb"]/div[1]/div/div[2]/a').click()

    # pegar a lista de resultados da busca no google shopping
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'KZmu8e')
    
    # para cada resultado, ele vai verificar se o resultado corresponde a todas as nossas condicoes
    lista_ofertas = [] # lista que a função vai me dar como resposta
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'translate-content').text
        nome = nome.lower()

        # verificacao do nome - se no nome tem algum termo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        
        # verificar se no nome tem todos os termos do nome do produto
        tem_todos_termos_produto = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_todos_termos_produto = False

        if not tem_termos_banidos and tem_todos_termos_produto: # verificando o nome
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'T14wmb').text
                preco = preco.replace("R$", "")
                preco = preco.replace(" ", "")
                preco = preco.replace(".", "")
                preco = preco.replace(",", ".")
                preco = preco.replace(".5", "")
                preco = preco.replace("Recondicionado", "")
                preco = preco.replace(".7", "")
                preco = preco.replace(".003", "")
                preco = float(preco)
                # verificando se o preco ta dentro do minimo e maximo
                if preco_minimo <= preco <= preco_maximo:
                    elemento_referencia = resultado.find_element(By.CLASS_NAME, 'ROMz4c')
                    elemento_pai = elemento_referencia.find_element(By.XPATH, '..')  
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
            except:
                continue

            
    return lista_ofertas


nav = webdriver.Chrome()
lista_ofertas_google = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
display(lista_ofertas_google)


tabela_ofertas = pd.DataFrame()
for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome'] 
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']
    
    lista_ofertas_google = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google:
        tabela_google_shop = pd.DataFrame(lista_ofertas_google, columns = ['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shop])
    else:
        tabela_google_shop = None
display(tabela_ofertas)

tabela_ofertas.to_excel('Ofertas.xlsx', index = False)


# Enviando E-mail

# In[18]:


if len(tabela_ofertas) > 0:
    import smtplib
    import email.message
    
    def enviar_email():  
        corpo_email = f"""
        <p>Olá, Meu amorzão esse e-mail ta sendo enviado de forma automática apenas para teste de um robô que eu criei
        que procura ofertas na internet dos produtos desejados na faixa de preço desejada, no seu caso procurei algumas das maquiagens mais baratas
        beijos te amo <3.</p>

        <p>{tabela_ofertas.to_html(index = False)}</p>
        <p>Atenciosamente, Caio Toscano (mestre do python)</p>
        """
    
        msg = email.message.Message()
        msg['Subject'] = "Ofertas Produtos pro meu Amorzão"
        msg['From'] = 'caio.toscano345@gmail.com'
        msg['To'] = 'caiofirst2@gmail.com'
        password = 'tnox bfyx mpdd pzuf' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )
    
        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')

    enviar_email()

