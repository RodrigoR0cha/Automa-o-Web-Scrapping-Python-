import time
import pandas as pd
from selenium import webdriver
from msilib.schema import tables
from selenium.webdriver.common.keys import Keys # Click do mouse ou teclado

# abrir o navegador
navegador = webdriver.Chrome()

# entrar no google (ou outro navegador)
navegador.get("https://www.google.com.br/")

# pesquisar cotação do dólar no google
navegador.find_element('xpath', './html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação do dólar")
navegador.find_element('xpath', './html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# pegar a cotação do dólar
cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

                                    
# Passo 2: pegar a cotação do Euro

# entrar no google (ou outro navegador)

navegador.get("https://www.google.com.br/")

# pesquisar cotação do euro no google
navegador.find_element('xpath', './html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação do Euro")
navegador.find_element('xpath', './html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# pegar a cotação do euro
cotacao_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)
                                 

# Passo 3: pegar a cotação do Ouro

# entrar no site melhor câmbio
navegador.get("https://www.melhorcambio.com/ouro-hoje")

# # pegar a cotação do Ouro no site melhor câmbio
cotacao_ouro = navegador.find_element('xpath', './/*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",",".")
navegador.quit()
print(cotacao_ouro)

# Passo 4: Atualizar a base de dados

tabela = pd.read_excel("Produtos.xlsx")
print(tabela)

# Passo 5: Recalcular os preços 

# Atualizar as cotações
tabela.loc[tabela["Moeda"]=="Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"]=="Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"]=="Ouro", "Cotação"] = float(cotacao_ouro)

# Preço de compra = Preço Original * Cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

# Preço de venda = Preço de compra * Margem
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]
print(tabela)

# Passo 6: Exportar a base de dados
tabela.to_excel("Produtos Novo.xlsx", index=False) # Criar uma nova base de dados atualizada em excel
