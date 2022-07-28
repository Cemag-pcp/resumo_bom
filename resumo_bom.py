import pandas as pd
import time
import gspread
import numpy as np
import glob
import os

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#from gspread_dataframe import set_with_dataframe
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-features=NetworkService")
options.add_argument("--window-size=1920x1080")
options.add_argument("--disable-features=VizDisplayCompositor")


def delete_selenium_log():
    if os.path.exists('selenium.log'):
        os.remove('selenium.log')


def show_selenium_log():
    if os.path.exists('selenium.log'):
        with open('selenium.log') as f:
            content = f.read()
            st.code(content)

def run_selenium():
    name = str()
    with webdriver.Chrome(options=options, service_log_path='selenium.log') as driver:
        url = "http://192.168.3.141/"
        driver.get(url)
        xpath = '//*[@class="ui-mainview-block eventpath-wrapper"]'
        # Wait for the element to be rendered:
        element = WebDriverWait(driver, 10).until(lambda x: x.find_elements(by=By.XPATH, value=xpath))
        name = element[0].get_property('attributes')[0]['name']
    return name


if __name__ == "__main__":
    delete_selenium_log()
    st.set_page_config(page_title="Selenium Test", page_icon='✅',
        initial_sidebar_state='collapsed')
    st.title('🔨 Selenium Test for Streamlit Sharing')

  
name_sheet = 'calculo de custo'
worksheet1 = 'Base de carretas'
worksheet2 = 'Extração do BOM'
worksheet3 = '% de perda'
worksheet4 = 'Custo contabil'

filename = "service_account.json"

sa = gspread.service_account(filename)
sh = sa.open(name_sheet)

wks1 = sh.worksheet(worksheet1)
wks2 = sh.worksheet(worksheet2)
wks3 = sh.worksheet(worksheet3)
wks4 = sh.worksheet(worksheet4)

#obtendo todos os valores da planilha
list1 = wks1.get_all_records()
list2 = wks2.get_all_records()
list3 = wks3.get_all_records()
list4 = wks4.get_all_records()

#transformando em dataframe
table = pd.DataFrame(list1)
base_perda = pd.DataFrame(list3)
base_contabil = pd.DataFrame(list4)

#itens para retirar
itens = table['Código'].values.tolist()
itens = str(itens)
itens = itens.strip('][').split(', ')

carretas = table
carretas = carretas.replace("", np.nan)
carretas.dropna(subset=['Carretas'], inplace=True)

############CHROME##################### --------------

link1 = "http://192.168.3.141/"
nav = webdriver.Chrome()
nav.get(link1)

#definindo data para exportação
data = table['Data'][0]

#logando 
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("luan araujo")
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo1234")

time.sleep(2)

WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

time.sleep(2)

#abrindo menu
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1892603865"]/table/tbody/tr/td[2]'))).click()

time.sleep(2)

#clicando em producao
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]'))).click()

time.sleep(1.5)

#clicando em cons gerenciais
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div[2]/div[9]/span[2]'))).click()

time.sleep(1.5)

#clicando em BOM
WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTreeNavegation"]/div[15]/span[2]'))).click()

iframe = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
nav.switch_to.frame(iframe)

#colocando data
data_input = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')))
data_input.click()
time.sleep(1)
data_input.clear()
time.sleep(1)
data_input.send_keys(data)
data_input.send_keys(Keys.TAB)

tabela1 = pd.DataFrame()

nav.switch_to.default_content()

#inputando carreta
#14:34
#13 minutos = 32 carretas

time.sleep(2)

#PODE DAR ERRO AQUI!!!!!!

for i in range(len(carretas['Carretas'])):
    
    tabelona = pd.DataFrame()
    
    name_carreta = carretas['Carretas'][i]
    
    iframe = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe)
    
    carreta = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
    campo_recurso = carreta.get_attribute('value')
    
    time.sleep(2)
    carreta.click()
    time.sleep(2)
    carreta.clear()
    time.sleep(2)
    carreta.click()
    time.sleep(2)
    carreta.send_keys(str(name_carreta))
    time.sleep(2)
    carreta.send_keys(Keys.TAB)
    time.sleep(2)
       
    try:
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[1]/input'))).click()
        time.sleep(1)
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[2]/div/table/tbody/tr/td[1]/span[2]'))).click()
    except:
        pass
    
    try:
        nav.switch_to.default_content()
        executar = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]')))
        executar.click()
        executar.click()
        
        time.sleep(10)
        
        nav.switch_to.default_content()
        
        iframe2 = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
        nav.switch_to.frame(iframe2)
        
        time.sleep(2)

        table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]')))
        table_html_prod = table_prod.get_attribute('outerHTML')
        
        time.sleep(2)
        
        tabelona = pd.read_html(str(table_html_prod), header=None)
        tabelona = tabelona[0]
        tabelona = tabelona.dropna()
        
        tabelona = tabelona.reset_index(drop=True)
        
        tabelona = tabelona.droplevel(level=0,axis=1)
        
        tabelona['Carreta'] = name_carreta
        tabelona['DataRef'] = data
        
        for i in range(len(tabelona)):
            if len(tabelona['Custo'][i]) > 6 :
                tabelona['Custo'][i] = tabelona['Custo'][i].replace(',','')
                tabelona['Custo'][i] = tabelona['Custo'][i].replace('.','')
        
        try:
            for j in range(len(tabelona)):
                if len(tabelona['Custo'][j]) >= 3 :
                    tabelona['Custo'][j] = float(tabelona['Custo'][j]) / 100
                else:
                    tabelona['Custo'][j] = float(tabelona['Custo'][j]) / 10
        except:
            pass
        
        try:
            for j in range(len(tabelona)):
                if len(tabelona['Qtd.'][j]) >= 3 :
                    tabelona['Qtd.'][j] = float(tabelona['Qtd.'][j]) / 100
                else:
                    tabelona['Qtd.'][j] = float(tabelona['Qtd.'][j]) / 10
        except:
            pass

        
        #tabelona['Código'] = tabelona['Código'].astype(str)
    
        #tabela1 = tabela1.append(tabelona)
        
        time.sleep(2)
        
        nav.switch_to.default_content()
        
        WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="buttonsCell"]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div'))).click()
        
        time.sleep(1.5)
        
        mask = tabelona['Código'].str.contains('Total')    
        mask1 = tabelona['Código'].str.contains('/Menu')

        tabelona.drop(tabelona[mask].index, axis=0, inplace=True)
        tabelona.drop(tabelona[mask1].index, axis=0, inplace=True)
        
        #filtrando linhas que contenham os codigos
        #filtro = tabelona.query("Código in [@itens]")
        filtro = tabelona[tabelona['Código'].isin(itens)]

        #excluindo linhas com os códigos da lista
        tabelona.drop(filtro.index, axis=0, inplace=True)
        
        tabelona.dtypes
        type(itens)
                
        tabelona = tabelona.reset_index(drop=True)

        tabela1_list = tabelona.values.tolist()

        sh.values_append('Extração do BOM', {'valueInputOption': 'RAW'}, {'values': tabela1_list})
        
    except:
        pass

nav.quit()
