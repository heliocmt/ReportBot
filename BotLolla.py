#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
from IPython import get_ipython
import pyautogui as py
from selenium import webdriver
import time
import pandas as pd
from twilio.rest import Client
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
import dataframe_image as dfi
from selenium.webdriver import ActionChains
import clipboard

py.PAUSE = 1.2

# Alerta inicial
py.alert("O código vai começar a rodar. Clique em OK e aguarde terminar.")

# Definindo datas e formatos que serão usados no código:
dia_hoje = datetime.datetime.today().date()
dia_hoje2 = dia_hoje.strftime("%d_%m_%Y") # O Bling ERP baixa os relatórios com esse formato no nome
dia_hoje3 = dia_hoje.strftime("%d/%m/%Y") # Modelo BR para ser enviado na mensagem de WhatsApp
dia_mes_passado = date.today() + relativedelta(months=-1)
dia_mes_passado = dia_mes_passado.strftime("%d/%m/%y") # Mesmo dia do mês passado para configurar busca do relatório

# Definindo todas as funções do código:

def open_website(website):
    # abrir o navegador e maximizar e entrar em um site
    py.PAUSE = 1
    global navegador 
    navegador = webdriver.Chrome("chromedriver.exe")
    navegador.maximize_window()
    navegador.get(website)
    time.sleep(2)
#------------------------------------------------------------------------------------------
def erp_login(erp_report_area):
    # Abrir erp na área de relatórios
    py.PAUSE = 1
    open_website(erp_report_area)
    # Fazer login
    navegador.find_element("xpath", '//*[@id="username"]').click()
    navegador.find_element("xpath", '//*[@id="username"]').send_keys("meulogin")
    navegador.find_element("xpath", '//*[@id="senha"]').send_keys("minhasenha")
    navegador.find_element("xpath", '//*[@id="login-buttons-site"]/button').click()
    time.sleep(3)
#------------------------------------------------------------------------------------------
def rename_report(new_name):
    py.PAUSE = 1
    # Abrir pasta downloads
    open_downloads()
    # Procurar a base de dados na pasta e renomear
    py.hotkey("ctrl", "f")
    time.sleep(1.3)
    py.write(f"relatorio_{dia_hoje2}.csv")
    py.press("enter")
    time.sleep(2)
    py.press("tab")
    time.sleep(1.5)
    py.press("up")
    py.press("F2")
    py.write(new_name)
    py.press("enter")
    py.hotkey("alt","F4")
#------------------------------------------------------------------------------------------
def upload_png():
    # Abrir pasta da base de dados em PNG
    py.press("win")
    time.sleep(2)
    py.write("daily_report.png")
    py.press("right")
    py.press("down")
    py.press("enter")
    # Fazer upload do PNG no ImgUr
    py.hotkey("ctrl","c")
    open_website("https://imgur.com/upload")
    py.hotkey("ctrl", "v")
    time.sleep(8)
    action = ActionChains(navegador)         
    source = navegador.find_element("xpath", '//*[@id="root"]/div/div[1]/div/div[4]/div[1]/div/div[2]/div[1]/div[2]/img')
    action.context_click(source).perform()
    time.sleep(4)
    py.press(["down","down","down","down"])
    py.press("enter")
    # Armazenar URL
    link = clipboard.paste()
    py.hotkey("alt","F4")
    return link
#------------------------------------------------------------------------------------------
def etl_daily_report():
    # Ler base de dados na pasta downloads
    tabela = pd.read_csv(r"C:\Users\helio\Downloads\daily_report.csv", sep=';')
    # Tratamento da base de dados
    tabela.rename(columns={'Valor': 'Valor Total'}, inplace = True)
    tabela['Qtde'] = tabela['Qtde'].str.replace(',00','').astype(int)
    tabela['Preço Médio'] = tabela['Preço Médio'].str.replace('.','')
    tabela['Preço Médio'] = tabela['Preço Médio'].str.replace(',','.').astype(float)
    tabela['Valor Total'] = tabela['Valor Total'].str.replace('.','')
    tabela['Valor Total'] = tabela['Valor Total'].str.replace(',','.').astype(float)
    tabela['Desconto'] = tabela['Desconto'].str.replace('.','')
    tabela['Desconto'] = tabela['Desconto'].str.replace(',','.').astype(float)
    tabela['Frete'] = tabela['Frete'].str.replace('.','')
    tabela['Frete'] = tabela['Frete'].str.replace(',','.').astype(float)
    tabela['Outras despesas'] = tabela['Outras despesas'].str.replace('.','')
    tabela['Outras despesas'] = tabela['Outras despesas'].str.replace(',','.').astype(float)
    tabela['Total Venda'] = tabela['Total Venda'].str.replace('.','')
    tabela['Total Venda'] = tabela['Total Venda'].str.replace(',','.').astype(float)
    tabela['Descontado'] = tabela['Valor Total'] - tabela['Desconto']
    tabela = tabela[['Produto','Código','Qtde','Preço Médio','Valor Total','Desconto', 'Descontado','Frete','Outras despesas', 'Total Venda']]
    # Exportar base de dados em png
    dfi.export(tabela, "daily_report.png")
#------------------------------------------------------------------------------------------
def get_total_value():
    py.PAUSE = 1.5
    # ERP login
    erp_login("https://www.bling.com.br/login?r=https%3A%2F%2Fwww.bling.com.br%2Frelatorio.vendas.php")
    # Configurar busca do ERP e baixar banco de dados
    navegador.find_element("xpath", '//*[@id="campo1"]').click()
    py.press("down")
    py.press("enter")
    navegador.find_element("xpath", '//*[@id="btn_visualizar"]').click()
    time.sleep(3)
    try:
        # Baixar relatório mensal
        navegador.find_element("xpath", '//*[@id="exportarRelatorio"]').click()
        time.sleep(1.2)
        navegador.find_element("xpath", '/html/body/div[15]/div[3]/div/button[1]').click()
        py.hotkey("alt","F4")
        # Renomear e ler relatório mensal
        rename_report("total_value")
        total_value_df = pd.read_csv(r"C:\Users\helio\Downloads\total_value.csv", sep=';')
        total_value_df = total_value_df.set_index('Produto')
        # Armazenar receita mensal total
        total_value = total_value_df.loc['Totais','Total Venda'] 
    except:
        total_value = "Não houve receita ainda."
        py.hotkey("alt","F4")
    return total_value
#------------------------------------------------------------------------------------------
def get_last_month():
    py.PAUSE = 1.5
    # ERP login
    erp_login("https://www.bling.com.br/login?r=https%3A%2F%2Fwww.bling.com.br%2Frelatorio.vendas.php")
    # Configurar parâmetros da pesquisa do relatório do mesmo dia do mês passado
    navegador.find_element("xpath", '//*[@id="periodoPesq"]').click()
    py.press("down")
    py.press("enter")
    navegador.find_element("xpath", '//*[@id="campo1"]').click()
    py.press("down")
    py.press("enter")
    navegador.find_element("xpath", '//*[@id="p-data"]').click()
    py.hotkey("ctrl","a")
    py.write(dia_mes_passado)
    navegador.find_element("xpath", '//*[@id="btn_visualizar"]').click()
    time.sleep(4)
    try:
        # Baixar o relatório do mesmo dia do mês passado
        navegador.find_element("xpath", '//*[@id="exportarRelatorio"]').click()
        time.sleep(1.2)
        navegador.find_element("xpath", '/html/body/div[15]/div[3]/div/button[1]').click()
        py.hotkey("alt","F4")
        # Renomear e ler relatório
        rename_report("last_month")
        last_month_df = pd.read_csv(r"C:\Users\helio\Downloads\last_month.csv", sep=';')
        last_month_df = last_month_df.set_index('Produto')
        # Armazenar receita do mesmo dia do mês passado
        last_month = last_month_df.loc['Totais','Total Venda']
    except:
        # Se não houver ou cair em final de semana, a receita será igual a zero
        last_month = 0
        py.hotkey("alt","F4")
    return last_month

#------------------------------------------------------------------------------------------
def get_report():
    py.PAUSE = 1.5
    # ERP Login
    erp_login("https://www.bling.com.br/login?r=https%3A%2F%2Fwww.bling.com.br%2Frelatorio.vendas.php")
    # Configurar parâmetros da busca do relatório no ERP
    navegador.find_element("xpath", '//*[@id="periodoPesq"]').click()
    py.press("down")
    py.press("enter")
    navegador.find_element("xpath", '//*[@id="campo1"]').click()
    py.press("down")
    py.press("enter")
    navegador.find_element("xpath", '//*[@id="btn_visualizar"]').click()
    time.sleep(4)
    try:
        # Baixar o relatório diário
        navegador.find_element("xpath", '//*[@id="exportarRelatorio"]').click()
        time.sleep(1.2)
        navegador.find_element("xpath", '/html/body/div[15]/div[3]/div/button[1]').click()
        py.hotkey("alt","F4")
        # Renomear o relatório diário
        rename_report("daily_report")
        # Manipular os dados do relatório diário
        etl_daily_report()
        # Hospedar o relatório na internet
        link = upload_png()
    except:
        # Link da logo para caso não exista receita no dia
        link = "https://uselollafit.com.br/wp-content/uploads/2022/05/Logo-Lolla-fit.svg" 
        py.hotkey("alt","F4")
    return link
  
#------------------------------------------------------------------------------------------
def get_expenses():
    # Login no ERP na área de finanças e buscar relatório
    erp_login("https://www.bling.com.br/relatorio.financas.resumo.categoria.php")
    time.sleep(3)
    navegador.find_element("xpath", '//*[@id="btn-pesquisar"]').click()
    time.sleep(5)
    # Armazenar valor total de despesa mensal
    expenses = navegador.find_element("xpath", '//*[@id="resultado_sai"]/tfoot/tr/th[2]').text
    py.hotkey("alt","F4")
    return expenses

#----------------------------------------------------------------------------------------------
def open_downloads():
    # Abrir pasta downloads e apagar relatório antigo
    py.hotkey("win", "r")
    py.write("shell:downloads")
    py.press("enter")
    time.sleep(2)
    
#----------------------------------------------------------------------------------------------
def delete_report(report_name):
    # Abrir pasta Downloads
    open_downloads()
    # Buscar e deletar relatório
    py.hotkey("ctrl", "f")
    py.write(report_name)
    py.press("enter")
    py.press("tab")
    py.press("down")
    py.press("del")
    py.hotkey("alt","F4")
    
#----------------------------------------------------------------------------------------------
def delete_img():
    # Abrindo destino e deletando png
    py.press("win")
    time.sleep(2)
    py.write("daily_report.png")
    py.press("right")
    py.press("down")
    py.press("enter")
    py.press("del")
    py.hotkey("alt","F4")    
#----------------------------------------------------------------------------------------------

# INÍCIO DO CÓDIGO

py.press("numlock")
link = get_report()
total_value = get_total_value()
last_month = get_last_month()
expenses = get_expenses()
py.press("numlock")

account_sid = 'MY TWILIO ACCOUNT SID'
auth_token = 'MY TWILIO AUTH TOKEN'
client = Client(account_sid, auth_token)

message = client.messages.create(
                              body= f"*Relatório Diário: {dia_hoje3}*\nReceita em {dia_mes_passado}: *R${last_month}*\nReceita Mensal: *R${total_value}*\nDespesa Mensal: *R${expenses}*",
                              media_url= link,
                              from_='whatsapp: mytwilionumber',
                              to='whatsapp: mywhatsappnumber'
                          )
print(message.sid)

py.press("numlock")
delete_report("last_month.csv")
delete_report("daily_report.csv")
delete_report("total_value.csv")
delete_img()
py.press("numlock")

# Alerta Final
py.alert("Relatório enviado!\nClique em OK para finalizar. :)") 

