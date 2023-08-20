from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from datetime import date
from datetime import timedelta
from datetime import date
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import zipfile
import wget
import os

import time


usuario = "aanaya@topitop.com.pe"
contrasenia = "83676656"

driver = webdriver.Chrome()
driver.get('https://escritorio.acepta.pe/')
driver.maximize_window()
time.sleep(2)

#usuario
user = driver.find_element("xpath",'//*[@id="loginrut"]')
user.send_keys(usuario)

#password
password = driver.find_element("xpath",'//*[@id="container_login_form"]/div/div[2]/form/fieldset/div[2]/div/div[2]/div/input[1]')
password.send_keys(contrasenia)

#iniciar sesion
mBox = driver.find_element("xpath",'//*[@id="container_login_form"]/div/div[2]/form/fieldset/div[2]/div/div[3]/input').click()

#emitidos
mBox = driver.find_element("xpath",'//*[@id="badge_185"]').click()

#busqueda avanzada
mBox = driver.find_element("xpath",'//*[@id="form_params"]/div[5]/a').click()
time.sleep(2)
#tipofecha=Emision
mBox = driver.find_element("xpath",'//*[@id="TIPO_FECHA"]/option[2]').click()

#periodoDesde--1 dia menos
fIni = date.today() - timedelta(days=1)

fechaIni = driver.find_element("xpath","//input[@name='FSTART']")
fechaIni.send_keys(fIni.strftime("%d%m%Y"))

#periodoHasta--1 dia menos
fechaFin = driver.find_element("xpath","//input[@name='FEND']")
fechaFin.send_keys(fIni.strftime("%d%m%Y"))

#buton Buscar
buscar = driver.find_element("xpath","//input[@name='buscar2']").click()
time.sleep(8)

#esperar que boton "exportar" este disponible
#click boton exportar
ex=False
while ex==False:
    try:
        expor=driver.find_element("xpath","(//input[@name='Exportar'])[1]").click()
        ex=True
    except:
        pass

#enter
WebDriverWait(driver, 10).until(EC.alert_is_present())
driver.switch_to.alert.accept()
#ir a reportes
reportes = driver.find_element("xpath",'//*[@id="badge_193"]').click()
time.sleep(180)
#descargar reportes
btemp=False

while btemp==False:
    try:
        descg=driver.find_element("xpath","//a[@id='1Descargargrilla_reportesNEW']").get_attribute('href')
        #driver.get(descg)
        wget.download(descg,'D:/Program Files/Python - Projects/test/test.zip')
        btemp=True
    except Exception as e:
        print(e)
        time.sleep(60)
        pass

#extraer zip y renombrar archivo xls
archivo_zip=zipfile.ZipFile("D:/Program Files/Python - Projects/test/test.zip","r")
nameOriginal=archivo_zip.namelist()
tempSrt=''.join(nameOriginal)
archivo_zip.extractall(path="D:/Program Files/Python - Projects/test/")
ArchivoExcel="D:/Program Files/Python - Projects/test/"+tempSrt
NewName="D:/Program Files/Python - Projects/test/test.xls"
os.rename(ArchivoExcel,NewName)

input("Esperando que no se cierre webdriver: ")




