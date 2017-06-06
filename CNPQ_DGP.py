import os
import time
import csv
import codecs
import requests 
import time
import xlwt
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import win32com.client

comecocode = time.time()
wb = xlwt.Workbook() 
ws = wb.add_sheet('Grupos')
grupos_list= [['Nome do Grupo','Instituição','Líder(es)','Área','Localidade','Ano de Formação','Grande Área']]
ws.write(0,0,'Nome do Grupo')
ws.write(0,1,'Instituição')
ws.write(0,2,'Líder(es)')
ws.write(0,3,'Área')
ws.write(0,4,'Localidade')
ws.write(0,5,'Ano de Formação')
ws.write(0,6,'Grande Área')
print("\nDiretório dos Grupos de Pesquisa no Brasil - CNPQ")

options = Options()
options.add_argument("--start-maximized")

caminho = os.getcwd() + "\chromedriver.exe"

if caminho.find("Desktop"):
 shell = win32com.client.Dispatch("WScript.Shell")
 shortcut = shell.CreateShortCut("/Users/Public/Desktop/CNPQ_DGP.lnk")
 tcam = shortcut.Targetpath
tmn = len("CNPQ_DGP.exe")
tmnG = len(tcam)
corte = tmnG - tmn
if tcam:
 caminho = tcam[:corte] + "chromedriver.exe"
chromeDriverPath = caminho[2:].replace("\\","/")
os.environ["webdriver.chrome.driver"] = chromeDriverPath


browser = webdriver.Chrome(chromeDriverPath, chrome_options=options)
browser.get('http://dgp.cnpq.br/dgp/faces/consulta/consulta_parametrizada.jsf')



browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)
browser.find_element_by_id("idFormConsultaParametrizada:buscaRefinada").click()

time.sleep(1)
browser.execute_script("window.scrollTo(0, 900);")

#Select Sudeste
while True:
 try:
  triggerDropDown = browser.find_element_by_class_name("ui-selectonemenu-trigger")
  triggerDropDown.click();
  item = triggerDropDown.find_element_by_xpath("//*[@id= 'idFormConsultaParametrizada:idRegiao_panel']/div/ul/li[5]")
  item.click()
  break
 except:
  pass


#Select Rio de Janeiro
while True:
 try:
  triggerDropDown = browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:idUF']/div[3]")
  triggerDropDown.click();
  item = triggerDropDown.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:idUF_panel']/div/ul/li[4]")
  item.click() 
  break
 except:
  pass



time.sleep(1)
browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)
browser.find_element_by_id("idFormConsultaParametrizada:idPesquisar").click()

#time.sleep(10)
while True:
 try:
  browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
  select = Select(browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList_paginator_bottom']/select"))
  select.select_by_visible_text('100')
  total=browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList_paginator_bottom']/span[6]").text
  eula = browser.find_element_by_id("idFormConsultaParametrizada:resultadoDataList:%d:idBtnVisualizarEspelhoGrupo"%(0))
  browser.execute_script('arguments[0].scrollTo = arguments[0].scrollIntoView(true)', eula)
  break
 except:
  pass
#time.sleep(10)

numero1=int(total[20:])
numero = int(numero1/100)
numero = numero +1

time.sleep(10)

for z in range(0,numero):
 for x in range(0, 100):
   y = z*100

   y= x + y 
   time.sleep(1)
   eula = browser.find_element_by_id("idFormConsultaParametrizada:resultadoDataList:%d:idBtnVisualizarEspelhoGrupo"%(y))
   
   browser.execute_script('arguments[0].scrollTo = arguments[0].scrollIntoView(true)', eula)
   action = ActionChains(browser)
   action.move_to_element(eula).perform()
   time.sleep(2)
   wait = WebDriverWait(browser, 2).until(
       EC.presence_of_element_located((By.XPATH,"//*[@id='idFormConsultaParametrizada:resultadoDataList:%d:idBtnVisualizarEspelhoGrupo']" %(y)))
    )
	
   print('\n %d'%(y+1))
 
   try:
    grupo= browser.find_element_by_id("idFormConsultaParametrizada:resultadoDataList:%d:idBtnVisualizarEspelhoGrupo"%(y))
    grupot = grupo.get_attribute('textContent')
    
    print (grupot)
   except:
    grupot= 'N/A'
   
   try: 
    instituicao= browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList_list']/li[%d]/div/div[2]/div"%(x+1))
    instituicaot = instituicao.text
    print (instituicaot)
   except:
    instituicaot= 'N/A'
	
   try: 
    lider= browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList:%d:idBtnVisualizarEspelhoLider1']"%(y))
    lidert = lider.text
    print (lidert)
   except:
    lidert= 'N/A'
	
   try: 
    area= browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList_list']/li[%d]/div/div[5]/div"%(x+1))
    areat = area.text
    print (areat)
   except:
    areat= 'N/A'
	
	
   curWindowHndl = browser.current_window_handle

   grupo.click()

   time.sleep(1)
   browser.switch_to_window(browser.window_handles[1])

   try:
    endereco = browser.find_element_by_xpath("//*[@id='endereco']/fieldset/div[6]/div")
    enderecot= endereco.text
    print(enderecot)
   except:
    enderecot = 'N/A'

   try:
    garea = browser.find_element_by_xpath("//*[@id='identificacao']/fieldset/div[6]/div")
    gareat= garea.text
    print(gareat)
   except:
    gareat = 'N/A'
   
   
   try:
    ano = browser.find_element_by_xpath("//*[@id='identificacao']/fieldset/div[2]/div")
    anot= ano.text
    print(anot)
   except:
    anot = 'N/A'
   
   listinha = [grupot,instituicaot,lidert,areat,enderecot,anot,gareat]
   grupos_list.append(listinha)
   ws.write(y+1,0,grupot)
   ws.write(y+1,1,instituicaot)
   ws.write(y+1,2,lidert)
   ws.write(y+1,3,areat)
   ws.write(y+1,4,enderecot)
   ws.write(y+1,5,anot)
   ws.write(y+1,6,gareat)



   time.sleep(1)
   browser.close() #closes new tab
   browser.switch_to_window(curWindowHndl)
   t = y + 1
   if t == numero1:
    break
 
 #click next:
 time.sleep(5)
 browser.find_element_by_xpath("//*[@id='idFormConsultaParametrizada:resultadoDataList_paginator_bottom']/span[4]").click()
 time.sleep(10)
   
finalcode = time.time()   
listinha = ['Nome do Grupo','Instituição','Líder(es)','Área',str(comecocode),str(finalcode),str(finalcode-comecocode)]
grupos_list.append(listinha)

with open ('Grupos_CSV.csv','w', newline='') as file:
    writer=csv.writer(file)
    for row in grupos_list:
        writer.writerow([row])


#time.sleep(10)

wb.save('Grupos_Lista.xls')
os.system('Grupos_Lista.xls')
browser.close()