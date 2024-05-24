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

numero_CNPJ = '40434458000173'

# entrar no site da - https://pje-consulta-publica.tjmg.jus.br/
driver = webdriver.Chrome(r'C:\ProjetorPython\dev\chromedriver-win64\chromedriver.exe')
driver.get('https://eproc.jfrj.jus.br/eproc/externo_controlador.php?acao=processo_consulta_publica&acao_origem=&acao_retorno=processo_consulta_publica')
sleep(3)
CNPJ = str(numero_CNPJ)
pyautogui.click(985,24)
sleep(1)
pyautogui.click(301,525)
sleep(1)
pyautogui.click(615,516)
sleep(1)
pyautogui.write(CNPJ)
sleep(1)
pyautogui.click(1249,270)
sleep(1)
# Consultar Registro
href_Registro = driver.find_element(By.XPATH, "/html/body/div/div[3]/div[2]/div/div/form/div[3]/div/table/tbody/tr[2]/td[1]/a")
href_Registro.click()

processo = driver.find_elements(By.XPATH,"//*[@id='divInfraAreaTabela']/table/tbody/tr")
processo.click()

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

        
       

