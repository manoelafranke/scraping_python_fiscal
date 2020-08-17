import time
import requests
import pandas as pd
import json
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import xlsxwriter

def pegar_link(link_entrada):
    print("pegando link")
    option = Options()
    option.headless = True
    driver = webdriver.Firefox(executable_path=r'./geckodriver')
    driver.get(link_entrada)

    time.sleep(5)
    elemento = driver.find_element_by_xpath("//iframe[@id='iframeConteudo']")
    dados_compras = elemento.get_attribute('src')
    return dados_compras


def gerar_planilha(e,nome):
    print("gerando a planilha")
    imprimir_ano = e
    excel = pd.ExcelWriter(nome+'.xlsx', engine='xlsxwriter')
    imprimir_ano.to_excel(excel, sheet_name='Dados de um ANO específico')
    excel.save()  

#link_entrada = "https://www.sefaz.rs.gov.br/NFCE/NFCE-COM.aspx?p=43200801874166000108651230002707301002707314|2|1|1|A71A11D1AA86C0BC7280908857249DBB91AD3795"

link_entrada = input("entre com o link: ")
url = pegar_link(link_entrada)

option = Options()
option.headless = True
driver = webdriver.Firefox( executable_path=r'./geckodriver')
driver.get(url)

time.sleep(5)

posts = driver.find_elements_by_class_name('NFCCabecalho')


data_emisao = driver.find_elements_by_class_name('NFCCabecalho_SubTitulo')
data_emisao_arquivo = data_emisao[2].get_attribute('outerHTML')

data_dia = data_emisao_arquivo.split()[11]
data_hora = data_emisao_arquivo.split()[12]


dados_compras = posts[3].get_attribute('outerHTML')
soup = BeautifulSoup(dados_compras,'html.parser')
table = soup.find(name="table")
dados = pd.read_html(str(table))[0]
#print("")
e = dados[[0,1,2,3,4,5]]
e.columns = ['Código','Descrição','Qtde','Un','Vl Unit','Vl Total']
print(e)
gerar_planilha(e,data_dia.replace("/","_"))

driver.quit()


