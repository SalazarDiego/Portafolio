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
from selenium.common.exceptions import *
import os
import zipfile
import wget
import time

driver = webdriver.Chrome()
driver.get('https://escritorio.acepta.pe/ext.php?r=https://escritorio.acepta.pe/appDinamicaClasses/index.php%3Fapp_dinamica=reportesNEW%26session_id=00776b390cd4bcea3d0154095d99343cd3586001433%26rutUsuario=11606%26rutCliente=20501057682%26aplicacion=DTE%26')
btemp=False

while btemp==False:
    try:
        descg=driver.find_element("xpath","//a[@id='1Descargargrilla_reportesNEW']").get_attribute('href')
        #driver.get(descg)
        wget.download(descg,'D:/Program Files/Python - Projects/test/test.zip')
        btemp=True
    except Exception as e:
        print(e)
        pass


archivo_zip=zipfile.ZipFile("D:/Program Files/Python - Projects/test/test.zip","r")
nameOriginal=archivo_zip.namelist()
tempSrt=''.join(nameOriginal)
archivo_zip.extractall(path="D:/Program Files/Python - Projects/test/")
ArchivoExcel="D:/Program Files/Python - Projects/test/"+tempSrt
NewName="D:/Program Files/Python - Projects/test/test.xls"
os.rename(ArchivoExcel,NewName)


input("Esperando que no se cierre webdriver: ")
