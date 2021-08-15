# Entrar a internet
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

driver = webdriver.Chrome()

# Pegar a cotação do Dólar
driver.get('https://www.google.com/')
driver.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dolar")
driver.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = driver.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')   

# Pegar a cotação do Euro
driver.get('https://www.google.com/')
driver.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
driver.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = driver.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')   

# Pegar a cotação do Lastro - Ouro
driver.get('https://www.melhorcambio.com/ouro-hoje')
cotacao_ouro = driver.find_element_by_xpath(
    '//*[@id="comercial"]').get_attribute('value').replace(',', '.')

driver.quit()

# Importar e atualizar a base de dados
tabela = pd.read_excel('./produtos.xlsx')

# Atualizar a cotação
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = round(float(cotacao_dolar), 2)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = round(float(cotacao_euro), 2)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = round(float(cotacao_ouro), 2)

# Calcular os Preços Base Reais
tabela['Preço Base Reais'] = round(tabela['Preço Base Original'] * tabela['Cotação'], 2)

# Calcular os Preços Finais
tabela['Preço Final'] = round(tabela['Preço Base Reais'] * tabela['Margem'], 2)

# Exportar a base de dados atualizada
tabela.to_excel('./ProdutosAtualizados.xlsx', index=False)