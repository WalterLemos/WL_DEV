from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import openpyxl
import pyautogui
import os

numero_CNPJ = '40434458000173'

## Configurando o caminho do executável como variável de ambiente
chrome_driver_path = r'C:\ProjetorPython\dev\chromedriver-win64\chromedriver.exe'
os.environ["webdriver.chrome.driver"] = chrome_driver_path
# Inicializando o driver do Chrome
driver = webdriver.Chrome()
driver.get('https://eproc.jfrj.jus.br/eproc/externo_controlador.php?acao=processo_consulta_publica&acao_origem=&acao_retorno=processo_consulta_publica')

pyautogui.click(1472,50)
sleep(1)
CNPJ = str(numero_CNPJ)
pyautogui.click(453,712)
sleep(1)
pyautogui.click(904,695)
sleep(1)
pyautogui.write(CNPJ)
sleep(1)
pyautogui.click(1760,334)
sleep(10)

 # extrair o n° do processo 
numero_processo = driver.find_element(By.XPATH,"//*[@id='divInfraAreaTabela']/table/tbody/tr[1]/td[1]/span").span
# extrair a data da distribuição
data_distribuicao = driver.find_element(By.XPATH,"//*[@id='divInfraAreaTabela']/table/tbody/tr[1]/td[2]/span").span
# extrair o Órgão Julgador
orgao_jugaldor = driver.find_element(By.XPATH,"//*[@id='divInfraAreaTabela']/table/tbody/tr[1]/td[3]/span/span" ).span
# extrair o Órgão Julgador
valor_causa = driver.find_element(By.XPATH,"//*[@id='fldInformacoesAdicionais']/table/tbody/tr[1]/td[2]/label" ).span
print("O Nº do Processo é:  ", numero_processo)
print("O Órgão de Distribuição é:  ", orgao_jugaldor)
print("A Data de Distribuição é:  ", data_distribuicao)
print("O Valor da Causa é:  ", valor_causa)

        
       

