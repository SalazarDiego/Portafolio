from datetime import *
import pyodbc
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import wget
import os
import zipfile
import time
from email.mime.text import MIMEText
from email.header import Header
import smtplib
import sys

def descargaAcepta_():
    

    #df1 = pd.read_excel(r"\\191.1.64.110\VentasRMS\CruceAcepta_20230706.csv");
    df1 = pd.read_excel(r"D:\Program Files\Python - Projects\test\reporte_1081514.xls");
    df1.drop(["receptor",
                "razon_social_receptor",
                "moneda",
                "isc",
                "ivap",
                "operacionesexoneradas",
                "operacionesinafectas",
                "operacionesgratis",
                "descuentoglobal",
                "otroscargos",
                "otrostributos",
                "icbper",
                "resumen_boleta",
                "uri",
                "referencias",
                "mensaje_sii",
                "motivo_observacion",
            ],
            axis=1,
            inplace=True,
        )
    df1 = df1[df1['tipo'] != 20]
    df1['nro_cpe_'] = df1.apply(lambda row: '0'+str(row['tipo'])+'-' + row['nro_cpe'], axis=1)
        
    return df1


def descargaAcepta():
    usuario = "dsalazar@topitop.com.pe"
    contrasenia = "111691341"

    driver = webdriver.Chrome()
    driver.get("https://escritorio.acepta.pe/")
    driver.maximize_window()
    time.sleep(2)

    # usuario
    user = driver.find_element("xpath", '//*[@id="loginrut"]')
    user.send_keys(usuario)

    # password
    password = driver.find_element(
        "xpath",
        '//*[@id="container_login_form"]/div/div[2]/form/fieldset/div[2]/div/div[2]/div/input[1]',
    )
    password.send_keys(contrasenia)

    # iniciar sesion
    mBox = driver.find_element(
        "xpath",
        '//*[@id="container_login_form"]/div/div[2]/form/fieldset/div[2]/div/div[3]/input',
    ).click()
    #aceptar en el mensaje de alerta
    time.sleep(5)
    mBox = driver.find_element("xpath","//button[@id='btn_close_aviso_pantalla']").click()

    # emitidos
    mBox = driver.find_element("xpath", '//*[@id="badge_185"]').click()

    # busqueda avanzada
    mBox = driver.find_element("xpath", '//*[@id="form_params"]/div[5]/a').click()
    time.sleep(2)
    # tipofecha=Emision
    mBox = driver.find_element("xpath", '//*[@id="TIPO_FECHA"]/option[2]').click()

    # periodoDesde-- 7 dia menos
    fIni = date.today() - timedelta(days=7)
    fFin = date.today() - timedelta(days=1)

    fechaIni = driver.find_element("xpath", "//input[@name='FSTART']")
    fechaIni.send_keys(fIni.strftime("%m%d%Y"))
    #fechaIni.send_keys(fIni.strftime("%d%m%Y"))


    # periodoHasta--1 dia menos
    fechaFin = driver.find_element("xpath", "//input[@name='FEND']")
    fechaFin.send_keys(fFin.strftime("%m%d%Y"))
    #fechaFin.send_keys(fFin.strftime("%d%m%Y"))

    # buton Buscar
    buscar = driver.find_element("xpath", "//input[@name='buscar2']").click()
    time.sleep(8)

    # esperar que boton "exportar" este disponible
    # click boton exportar
    ex = False
    while ex == False:
        try:
            expor = driver.find_element(
                "xpath", "(//input[@name='Exportar'])[1]"
            ).click()
            ex = True
        except:
            pass

    # enter
    WebDriverWait(driver, 10).until(EC.alert_is_present())
    driver.switch_to.alert.accept()
    # ir a reportes
    reportes = driver.find_element("xpath", '//*[@id="badge_193"]').click()
    time.sleep(180)
    # descargar reportes
    btemp = False

    while btemp == False:
        try:
            descg = driver.find_element(
                "xpath", "//a[@id='1Descargargrilla_reportesNEW']"
            ).get_attribute("href")
            # driver.get(descg)
            fecha = date.today() - timedelta(days=1)
            fecha = fecha.strftime("%Y%m%d")
            ruta = r"\\191.1.64.110\VentasRMS\\CruceAcepta_" + fecha + ".zip"
            ruta2 = r"\\191.1.64.110\VentasRMS\\CruceAcepta_" + fecha + ".csv"
            ruta3 = r"\\191.1.64.110\VentasRMS\\CruceAcepta_" + fecha + ".xls"
            if os.path.exists(ruta):
                os.remove(ruta)
            if os.path.exists(ruta2):
                os.remove(ruta2)
            if os.path.exists(ruta3):
                os.remove(ruta3)    

            wget.download(
                descg, r"\\191.1.64.110\VentasRMS" + "\\CruceAcepta_" + fecha + ".zip"
            )
            btemp = True
        except Exception as e:
            print(e)
            time.sleep(60)
            pass

    # extraer zip y renombrar archivo xls
    driver.close()
    # path= r"\\191.1.64.110\VentasRMS"+"\\CruceAcepta_"+fecha+".xls"

    fecha = date.today() - timedelta(days=1)
    fecha = fecha.strftime("%Y%m%d")

    archivo_zip = zipfile.ZipFile(r"\\191.1.64.110\VentasRMS\\CruceAcepta_" + fecha + ".zip")
    
    nameOriginal = archivo_zip.namelist()
    tempStr = "".join(nameOriginal)
    archivo_zip.extractall(path=r"\\191.1.64.110\VentasRMS")
    if 'xls' in nameOriginal[0]:
        NewName = r"\\191.1.64.110\VentasRMS" + "\\CruceAcepta_" + fecha + ".xls"
    else:
        NewName = r"\\191.1.64.110\VentasRMS" + "\\CruceAcepta_" + fecha + ".csv"

    ArchivoExcel = r"\\191.1.64.110\VentasRMS" + "\\" + tempStr
    #NewName = r"\\191.1.64.110\VentasRMS" + "\\CruceAcepta_" + fecha + ".xls"
    #NewName = r"\\191.1.64.110\VentasRMS" + "\\CruceAcepta_" + fecha + ".csv"
    os.rename(ArchivoExcel, NewName)

    if 'xls' in nameOriginal[0]:
        df1=pd.read_excel(NewName)
    else:
        df1=pd.read_csv(NewName)
    #df1 = pd.read_csv(NewName)
    df1.drop(["receptor",
                "razon_social_receptor",
                "moneda",
                "isc",
                "ivap",
                "operacionesexoneradas",
                "operacionesinafectas",
                "operacionesgratis",
                "descuentoglobal",
                "otroscargos",
                "otrostributos",
                "icbper",
                "resumen_boleta",
                "uri",
                "referencias",
                "mensaje_sii",
                "motivo_observacion",
            ],
            axis=1,
            inplace=True,
        )
    df1 = df1[df1['tipo'] != 20]
    df1['nro_cpe_'] = df1.apply(lambda row: '0'+str(row['tipo'])+'-' + row['nro_cpe'], axis=1)
        
    return df1

def consultaRMS_BD():
    server = "191.1.64.111"
    database = "TFL"
    username = "retailuser"
    password = "retail"

    # cadena de conexion
    conn = pyodbc.connect(
        "DRIVER={SQL Server Native Client 11.0};SERVER=191.1.64.111;DATABASE=TFL;UID=retailuser;PWD=retail"
    )
    cursor = conn.cursor()

    # setear fecha ini/fin un dia anterior
    fIni = datetime.now() - timedelta(days=7)
    fIni = fIni.replace(hour=0, minute=0, second=0).strftime("%Y-%m-%d %H:%M:%S")
    fFin = datetime.now() - timedelta(days=1)
    fFin = fFin.replace(hour=23, minute=59, second=59).strftime("%Y-%m-%d %H:%M:%S")
    # query para ejecutar sp 2023-04-11 00:00:00 2023-04-11 23:59:59
    # query="EXEC Sp_Integration_StoreSummaryCustomerCompany_forDay '2023-04-11 00:00:00' , '2023-04-11 23:59:59' , 1"
    query = "EXEC Sp_Integration_StoreSummaryCustomerCompany_forDay_ @DateFrom = '{0}' ,@DateTo='{1}', @StoreNo={2}".format(
        fIni, fFin, 9999
    )
    # mostrar dataframe usando lib pandas
    dfRMS = pd.read_sql_query(query, conn)
    # print(df2.head(10))
    return dfRMS
    conn.close()
def consultaSAP_BD():
    server = "191.1.64.111"
    database = "BDINTEGRATION"
    username = "retailuser"
    password = "retail"

    # cadena de conexion
    conn = pyodbc.connect(
        "DRIVER={SQL Server Native Client 11.0};SERVER=191.1.64.111;DATABASE=BDINTEGRATION;UID=retailuser;PWD=retail"
    )
    cursor = conn.cursor()

    # setear fecha ini/fin un dia anterior
    fini = datetime.now() - timedelta(days=7)
    fini = fini.strftime("%Y%m%d")
    fFin = datetime.now() - timedelta(days=1)
    fFin = fFin.strftime("%Y%m%d")
    
    query = "EXEC sp_appConsultaSAP @fini = '{0}', @fFin = '{1}' ".format(fini,fFin)
    # mostrar dataframe usando lib pandas
    dfSAP = pd.read_sql_query(query, conn)
    return dfSAP

    conn.close()
def Cruce():


    df1 = consultaSAP_BD()
    df2 = consultaRMS_BD()
    df3 = descargaAcepta()

    #-----------------------------SAP - RMS-------------------------------------------------------------------
    # busca los nro_cpe de RMS que no estan en RMSFREF de SAP
    df_RmsDocNotInSAP = df2[~df2.nro_cpe.isin(df1.RMSFREF) & df2.nro_cpe.notna() & (df2.nro_cpe != "")]
    # Realizar la fusión (merge) de los dataframes en función del código (nro_cpe|rms y RMSFREF|sap)
    merged_df = pd.merge(df1, df2, left_on="RMSFREF", right_on="nro_cpe", suffixes=("_df1", "_df2"))
    # Realizar diferencia de importes rms - sap
        #merged_df["Diferencia"] = merged_df["PayTotal"] - merged_df["RMSFMONT"]
    # Seleccionar las columnas del df1 y la columna "Diferencia"
        #result_df = merged_df[["nro_cpe", "StoreCode", "StoreName", "emision", "Diferencia"]]
    # Filtrar resultados, donde la diferencia es mayor o igual 0.1
        #filtered_df = result_df[result_df["Diferencia"] >= 0.99]
    #resultadoFinalSAP = pd.concat([df_RmsDocNotInSAP, filtered_df])
    resultadoFinalSAP = df_RmsDocNotInSAP
    #----------------------------------------------------------------------------------------------------------

    #-----------------------------------------------Acepta - RMS -----------------------------------------------------

    #elimiar documentos q empiecen con FW y FX
    df4 = df3[~(df3['nro_cpe'].str.startswith('FX') | df3['nro_cpe'].str.startswith('FW') | df3['nro_cpe'].str.startswith('BX') | df3['nro_cpe'].str.startswith('BG'))] 
    
    df_rechazado = df4[df4["estado_acepta"] != "ACEPTADO"]
    df_AcepDocNotInRMS = df4[~df4.nro_cpe_.isin(df2.nro_cpe) & df4.nro_cpe_.notna() & (df4.nro_cpe_ != '')]
    resultadoFinalAcepta = pd.concat([df_rechazado,df_AcepDocNotInRMS])
    #----------------------------------------------------------------------------------------------------------

    #-----------------------------------------------RMS - ACEPTA ---------------------------------------------------------
    resultadoFinalRMS = df2[~df2.nro_cpe.isin(df3.nro_cpe_) & df2.nro_cpe.notna() & (df2.nro_cpe != '')]
    #-----------------------------------------------------------------------------------------------------------

    #----------------------------------------------SAP - ACEPTA--------------------------------------------
    #---Documentos que estan en sap pero no en acepta
    resultadoFinalSapAcep = df1[~df1.RMSFREF.isin(df3.nro_cpe_) & df1.RMSFREF.notna() & (df1.RMSFREF != '')]
    #------------------------------------------------------------------------------------------------------
    #-----------------ENVIAR CORREO (ALERTA)-----------------------

    html = """\
    <html>
    <head></head>
    <body>
        <h2 style='text-align: left; color : #5499C7;'>TRADING FASHION LINE S.A.</h2>
        <p>Fecha : {0}</p>
        <p> Rango de Cruce: 7 Dias</p>
        <h3>ACEPTA - RMS</h3>
        <p>Documentos que están en ACEPTA y no están en RMS y/o estado distinto.</p>
        {1}
        <br>
        <br>
        <h3>RMS - ACEPTA </h3>
        <p>Documentos que están en RMS y no están en ACEPTA y/o estado distinto.</p>
        {2} 
        <br>
        <h3>SAP - RMS </h3>
        <p>Documentos que están en RMS y no están en SAP y/o estado distinto.</p>
        {3} 
        <br>
        <h3>SAP - ACEPTA </h3>
        <p>Documentos que están en SAP y no están en ACEPTA y/o estado distinto.</p>
        {4} 
        <br>
        <p style='text-align: center; color: #999999;'>Nota: Este correo ha sigo generado automaticamente,no responder a este correo</p>
        <h2 style='text-align: center; color: #999999;'>© 2023 TI RETAIL TFL. All rights reserved.</h2>
    </body>
    </html>
    """.format((datetime.now()).strftime("%d-%m-%Y %H:%M"),resultadoFinalAcepta.to_html(index=False),resultadoFinalRMS.to_html(index=False),
               resultadoFinalSAP.to_html(index=False),resultadoFinalSapAcep.to_html(index=False))
               
              

    part1 = MIMEText(html, 'html')
    message= part1
    message['Subject'] = Header("Cruce RMS - Acepta -SAP",'utf-8')

    server=smtplib.SMTP('191.1.64.16',25)
    server.starttls()
    server.login('alertaprocesosti@topitop.com.pe','alertpro@.23*')
    destinatarios = ['dsalazar@topitop.com.pe','cbravo@topitop.com.pe','carenas@topitop.com.pe','rtello@topitop.com.pe','elezama@topitop.com.pe','pgarcia@topitop.com.pe','soporte_retail@topitop.com.pe']
    #destinatarios = ['dsalazar@topitop.com.pe','cbravo@topitop.com.pe']
    #destinatarios = ['dsalazar@topitop.com.pe']

    server.sendmail('alertaprocesosti@topitop.com.pe',destinatarios,message.as_string())

    server.quit()

Cruce()

sys.exit()
