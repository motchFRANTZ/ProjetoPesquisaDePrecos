from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from IPython.display import display
import time
import win32com.client as win32

nav = webdriver.Chrome()

tabela_buscas = pd.read_excel('buscas.xlsx')
display(tabela_buscas)

def termos_proibidos(lista_termo_banidos, nome):
    tem_termo_banido = False
    for palavra in lista_termo_banidos:
        if palavra in nome:
            tem_termo_banido = True
    return tem_termo_banido


def todas_palavras(lista_nome, nome):
    tem_todas_as_palavras = True
    for palavra in lista_nome:
        if palavra not in nome:
            tem_todas_as_palavras = False
    return tem_todas_as_palavras


def busca_google(nav, nome_produto, termos_banidos, preco_min, preco_max):
    url = 'https://www.google.com.br/'
    nav.get(url)
    nav.find_element(By.CLASS_NAME, 'gLFyf').click()
    nav.find_element(By.CLASS_NAME, 'gLFyf').send_keys(
        nome_produto, Keys.ENTER)
    time.sleep(2)

    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')

    nome_produto = nome_produto.lower()
    lista_nome_produto = nome_produto.split(' ')

    elementos = nav.find_elements('class name', 'YmvwI')
    for item in elementos:
        if "Shopping" in item.text:
            item.click()
            break

    lista_resultado = nav.find_elements('class name', 'i0X6df')
    lista_ofertas = []
    for elemento in lista_resultado:
        nome = elemento.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()

        tem_termo_banido = termos_proibidos(lista_termos_banidos, nome)
        tem_todas_as_palavras = todas_palavras(lista_nome_produto, nome)

        if not tem_termo_banido and tem_todas_as_palavras:
            try:
                preco = elemento.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace('R$', '').replace(' ', '').replace(
                    '.', '').replace(',', '.').replace('+impostos', '')
                preco = float(preco)

                preco_max = float(preco_max)
                preco_min = float(preco_min)

                if preco_min <= preco <= preco_max:
                    elemento_referencia = elemento.find_element(
                        'class name', 'bONr3b')
                    elemento_pai = elemento_referencia.find_element('xpath', '..')
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
            except:
                continue       
    return lista_ofertas


# Busca no Buscapé
def busca_buscape(nav, nome_produto, termos_banidos, preco_min, preco_max):
    url = 'https://www.buscape.com.br/'

    nome_produto = nome_produto.lower()
    lista_nome = nome_produto.split(' ')

    termos_banidos = 'zota galax'
    termos_banidos = termos_banidos.lower()
    lista_termo_banidos = termos_banidos.split(' ')

    nav.get(url)
    nav.find_element(
        By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(nome_produto, Keys.ENTER)

    lista_resultados_buscape = nav.find_elements(
        By.CLASS_NAME,
        'ProductCard_ProductCard__WWKKW')
    lista_ofertas = []
    for produto in lista_resultados_buscape:
        try:
            nome = produto.find_element(
                By.CLASS_NAME, 'ProductCard_ProductCard_NameWrapper__45Z01').text
            nome = nome.lower()

            tem_termo_banido = termos_proibidos(lista_termo_banidos, nome)
            tem_todas_as_palavras = todas_palavras(lista_nome, nome)

            if not tem_termo_banido and tem_todas_as_palavras:
                preco = produto.find_element(
                    By.CLASS_NAME, 'Text_Text__ARJdp.Text_MobileHeadingS__HEz7L').text
                preco = preco.replace('R$', '').replace(' ', '').replace(
                    '.', '').replace(',', '.').replace('+impostos', '')
                preco = float(preco)
    
                preco_max = float(preco_max)
                preco_min = float(preco_min)
    
                if preco_min <= preco <= preco_max:
                    link = produto.find_element(
                        By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh').get_attribute('href')
    
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas


tabela_ofertas = pd.DataFrame()


for linha in tabela_buscas.index:
    nome_produto = tabela_buscas.loc[linha, 'Nome']
    termos_banidos = tabela_buscas.loc[linha, 'Termos banidos']
    preco_max = tabela_buscas.loc[linha, 'Preço máximo']
    preco_min = tabela_buscas.loc[linha, 'Preço mínimo']

    lista_ofertas_google = busca_google(nav, nome_produto, termos_banidos, preco_min, preco_max)
    lista_ofertas_buscape = busca_buscape(nav, nome_produto, termos_banidos, preco_min, preco_max)

    if lista_ofertas_google:
        tabela_google = pd.DataFrame(lista_ofertas_google, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google], ignore_index=True)
    else:
        tabela_google = pd.DataFrame()

    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape], ignore_index=True)
    else:
        tabela_buscape = pd.DataFrame()

display(tabela_ofertas)

tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)


if len(tabela_ofertas.index) > 0:
    # vou enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'seu_email'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    """
    
    mail.Send()

nav.quit()
