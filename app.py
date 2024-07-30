"""
Site que o programa vai usar
https://contabilidade-devaprender.netlify.app/

"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

#Navegar até o site
driver = webdriver.Chrome() #Abrir o navegador
driver.get('https://contabilidade-devaprender.netlify.app/')
sleep(5)
#XPATH -> //tag[@atributo='valor']
#Digitar o e-mail
email = driver.find_element(By.XPATH,"//input[@id='email']")
sleep(3)
email.send_keys('admin@contabilidade.com')
sleep(3)
#Digitar a senha
senha = driver.find_element(By.XPATH,"//input[@id='senha']")
sleep(3)
senha.send_keys('contabilidade123456')
sleep(3)
#Clicar em entrar
entrar = driver.find_element(By.XPATH,"//input[@id='Entrar]")
sleep(2)
entrar.click()
sleep(5)
#Extraída da planilha
empresas = openpyxl.load_workbook('./empresas.xlsx')
paginaEmpresas = empresas['dados empresas']
for linha in paginaEmpresas.iter_rows(min_row=2, max_row=2, values_only=True):
    nomeEmpresa, email, telefone, endereco, cnpf, areaAtuacao, quantidadeFuncionaro, dataFundacao = linha
    #Clicar em cada campo e preencher com a informação 
    #Só é possível porque todos tem ID
    driver.find_element(By.ID,'nomeEmpresa').send_keys(nomeEmpresa)
    sleep(2)
    driver.find_element(By.ID,'emailEmpresa').send_keys(email)
    sleep(2)
    driver.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(2)
    driver.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(2)
    driver.find_element(By.ID,'cnpj').send_keys(cnpf)
    sleep(2)
    driver.find_element(By.ID,'areaAtuacao').send_keys(areaAtuacao)
    sleep(2)
    driver.find_element(By.ID,'numeroFuncionario').send_keys(quantidadeFuncionaro)
    sleep(2)
    driver.find_element(By.ID,'dataFundacao').send_keys(dataFundacao)
    sleep(2)
#Clicar em cadastrar
    cadastrar = driver.find_element(By.ID,'Cadastrar')
    sleep(2)
    cadastrar.click()
#Repito até acabar as informações da planilha

