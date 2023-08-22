from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import win32com.client as win32

#criar o navegador
servico = Service()
opcoes = Options()
opcoes.add_argument("--start-maximized")
navegador = webdriver.Chrome(service=servico, options=opcoes)

#importar/visualizar a base de dados
tabela_produtos = pd.read_excel('buscas.xlsx')
display(tabela_produtos)

def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(lista_termos_produto, nome):
    tem_todos_termos_produto = True
    for palavra in lista_termos_produto:
        if palavra not in nome:
            tem_todos_termos_produto = False
    return tem_todos_termos_produto

def busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produto = produto.split(' ')
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    #entrar no google
    navegador.get('https://www.google.com.br/')
    navegador.find_element(By.CLASS_NAME, 'gLFyf').send_keys(produto, Keys.ENTER)
    time.sleep(3)

    #entrar na aba shopping
    navegador.find_element(By.CLASS_NAME, 'UqcIvb').click()

    #pegar as informações dos produtos
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'i0X6df')

    for resultado in lista_resultados:
        #Tratamento do nome
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()
        
        #analisar se ele não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        
        #analisar se ele tem TODOS os termos do nome do produto
        tem_todos_termos_produto = verificar_tem_todos_termos_produto(lista_termos_produto, nome)
        
        #selecionar os elementos com tem_termos_banidos=False e tem_todos_termos_produto = True
        if not tem_termos_banidos and tem_todos_termos_produto:

            #Tratamento do preco
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace('R$', '').replace(' ','').replace('.','').replace(',','.')
            preco = float(preco)
            
            #verificar se o preco está entre preco_minimo e preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                elemento_referencia = resultado.find_element(By.CLASS_NAME, 'bONr3b')
                elemento_pai = elemento_referencia.find_element(By.XPATH, '..')
                link = elemento_pai.get_attribute('href')
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas
       
def busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo):
    #tratar os valores
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_produto = produto.split(' ')
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    #buscar no buscapé
    navegador.get('https://www.buscape.com.br')
    navegador.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
    
    #pegar os resultados
    while len(navegador.find_elements(By.CLASS_NAME, 'Select_Select__1S7HV')) < 1: #esperando a página após a busca carregar
        time.sleep(1)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb') #aqui estamos pegando o bloco de um anuncio completo
    
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Name__ZaO5o').text
        preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__Zxam2').text
        link = resultado.get_attribute('href')
        
        #Tratamento do nome
        nome = nome.lower()
    
        #analisar se o resultado tem termos banidos e tem todos os termos do produto
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_termos_produto = verificar_tem_todos_termos_produto(lista_termos_produto, nome)
        if not tem_termos_banidos and tem_todos_termos_produto:
            
            #Tratamento do preco
            preco = preco.replace('R$', '').replace(' ','').replace('.','').replace(',','.')
            preco = float(preco)
            
            #analisar se o preço está entre preco_minimo e preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                lista_ofertas.append((nome, preco, link))

    #retornar lista de ofertas do buscapé
    return lista_ofertas


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    #pesquisar o produto
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping: #verifica se tem algum item na lista
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
        
    lista_ofertas_buscape = busca_buscape(navegador, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None
        
display(tabela_ofertas)
navegador.close()

#Exportando para Excel
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '136493140+Hugo-Hattori@users.noreply.github.com'
    mail.Subject = 'Produto(s) encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados,</p>
    <p>Segue a lista de produtos encontrados dentro da faixa de preço desejada.</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Att.,</p>
    '''
    mail.Send()