import pandas as pd
import os
import time as t
import zipfile
from selenium import webdriver
from datetime import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import wget
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import locale
from decimal import Decimal
import sys

def descargaAcepta_():


    df1 = pd.read_csv(r"D:\Program Files\Python - Projects\test\reporte_1083529.csv")
    #df1 = pd.read_excel(r"D:\Program Files\Python - Projects\test\CruceAcepta_20230711.xls")
    #df1 = pd.read_excel(r"\\191.1.64.110\VentasRMS\CruceAcepta_20230711.xls")

    df1.drop(["receptor",
                "razon_social_receptor",
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
    
    df1['emision'] = pd.to_datetime(df1['emision'])
    
    #mes_actual = datetime.now().month
    
    df1 = df1[df1['emision'].dt.month == 6]
     
    return df1
    

def descargaAcepta():

    usuario = "dsalazar@topitop.com.pe"
    contrasenia = "111691341"

    driver = webdriver.Chrome()
    driver.get("https://escritorio.acepta.pe/")
    driver.maximize_window()
    t.sleep(2)

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
    t.sleep(5)
    mBox = driver.find_element("xpath","//button[@id='btn_close_aviso_pantalla']").click()

    # emitidos
    mBox = driver.find_element("xpath", '//*[@id="badge_185"]').click()

    # busqueda avanzada
    mBox = driver.find_element("xpath", '//*[@id="form_params"]/div[5]/a').click()
    t.sleep(2)
    # tipofecha=Emision
    mBox = driver.find_element("xpath", '//*[@id="TIPO_FECHA"]/option[2]').click()

    # FECHA INICIO
    fecha_actual = date.today()
    fIni = date(fecha_actual.year,fecha_actual.month,1)
    fFin = date.today() - timedelta(days=1)

    if ((fFin.day - fIni.day)==30 | (fFin.day - fIni.day)==29):
        fIni=date(fecha_actual.year,fecha_actual.month-1,1)


    fechaIni = driver.find_element("xpath", "//input[@name='FSTART']")
    fechaIni.send_keys(fIni.strftime("%m%d%Y"))
    #fechaIni.send_keys(fIni.strftime("%d%m%Y"))


    # periodoHasta--1 dia menos
    fechaFin = driver.find_element("xpath", "//input[@name='FEND']")
    fechaFin.send_keys(fFin.strftime("%m%d%Y"))
    #fechaFin.send_keys(fFin.strftime("%d%m%Y"))

    #Estado documento = Aceptado
    aceptado=driver.find_element("xpath", "(//option[@value='E_ACEPTADO'])[1]").click()

    # buton Buscar
    buscar = driver.find_element("xpath", "//input[@name='buscar2']").click()
    t.sleep(8)

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
    t.sleep(180)
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
            t.sleep(60)
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
    #df1['nro_cpe_'] = df1.apply(lambda row: '0'+str(row['tipo'])+'-' + row['nro_cpe'], axis=1)
    df1['emision'] = pd.to_datetime(df1['emision'])
    mes_actual = datetime.now().month
    df1 = df1[df1['emision'].dt.month == mes_actual]

    if ((fFin.day - fIni.day)==30 | (fFin.day - fIni.day)==29):
        df1 = df1[df1['emision'].dt.month == mes_actual-1]
        
    return df1

def reporte():

    df=descargaAcepta()
    try:
        tipo_mapping = {1: 'Facturas', 3: 'Boletas', 7: 'Nota de Credito', 8: 'Nota de Debito'}
        df['tipo'] = df['tipo'].replace(tipo_mapping)
        df.loc[df['tipo'] == 'Nota de Credito', 'monto_total'] *= -1
    except:pass

    df1 = df[df['moneda'] == 'USD' ]
    df = df[df['moneda'] == 'PEN' ]
    #print(df1)

    #----------------CREAR TABLA DE SOLES------------------------------------------------------
    #calcular monto total y cantidad
    tabla_pivote = pd.pivot_table(df, values='monto_total', index='tipo', aggfunc={'tipo': 'count','monto_total': 'sum' })
    tabla_pivote.columns = ['MontoTotal', 'Cantidad']

    # Calcular la suma total de las columnas
    total_sum = tabla_pivote.sum()

    # Crear una nueva fila con los totales
    total_row = pd.DataFrame({'MontoTotal': [total_sum['MontoTotal']], 'Cantidad': [total_sum['Cantidad']]}, index=['Total'])

    # Redondear y convertir 'Cantidad' a enteros
    tabla_pivote['Cantidad'] = tabla_pivote['Cantidad'].round().astype(int)
    total_row['Cantidad'] = total_row['Cantidad'].round().astype(int)
    
    # Concatenar la nueva fila al final de la tabla pivote
    tabla_pivote_with_total = pd.concat([tabla_pivote, total_row])

    #renombrar columnas
    tabla_pivote_with_total = tabla_pivote_with_total.reset_index()
    tabla_pivote_with_total = tabla_pivote_with_total.rename(columns={'index': 'Documentos'})
    tabla_pivote_with_total = tabla_pivote_with_total.rename(columns={'MontoTotal': 'MontoTotal(S/.)', 'Cantidad': 'CantidadTotal'})
    tabla_pivote_with_total['MontoTotal(S/.)'] = tabla_pivote_with_total['MontoTotal(S/.)'].apply(lambda x: '{:,.2f}'.format(x))
    tabla_pivote_with_total['CantidadTotal'] = tabla_pivote_with_total['CantidadTotal'].apply(lambda x: '{:,}'.format(x))
    
    #-------------------------------------CREAR TABLA DE DOLARES---------------------------------------------------
    tabla_pivotedol = pd.pivot_table(df1, values='monto_total', index='tipo', aggfunc={'tipo': 'count','monto_total': 'sum' })
    tabla_pivotedol.columns = ['MontoTotal', 'Cantidad']
    total_sumdol = tabla_pivotedol.sum()
    total_rowdol = pd.DataFrame({'MontoTotal': [total_sumdol['MontoTotal']], 'Cantidad': [total_sumdol['Cantidad']]}, index=['Total'])
    tabla_pivotedol['Cantidad'] = tabla_pivotedol['Cantidad'].round().astype(int)
    total_rowdol['Cantidad'] = total_rowdol['Cantidad'].round().astype(int)
    tabla_pivote_with_totaldol = pd.concat([tabla_pivotedol, total_rowdol])
    tabla_pivote_with_totaldol = tabla_pivote_with_totaldol.reset_index()
    tabla_pivote_with_totaldol = tabla_pivote_with_totaldol.rename(columns={'index': 'Documentos'})
    tabla_pivote_with_totaldol = tabla_pivote_with_totaldol.rename(columns={'MontoTotal': 'MontoTotal($)', 'Cantidad': 'CantidadTotal'})
    tabla_pivote_with_totaldol['MontoTotal($)'] = tabla_pivote_with_totaldol['MontoTotal($)'].apply(lambda x: '{:,.2f}'.format(x))
    tabla_pivote_with_totaldol['CantidadTotal'] = tabla_pivote_with_totaldol['CantidadTotal'].apply(lambda x: '{:,}'.format(x))

    #--------------REPORTE VENTAS POR DIA--------------------------------------------
    tabla = df.pivot_table(index='emision', values='monto_total', aggfunc='sum')
    tabla = tabla.sort_values('emision', ascending=True)

    tabla.index = pd.to_datetime(tabla.index)
    tabla.index = tabla.index.strftime('%d-%m-%y')

    suma_columna = tabla['monto_total'].sum()
    nueva_fila = pd.DataFrame({'monto_total': suma_columna}, index=['Total'])

    tabla_sorted = pd.concat([tabla,nueva_fila])
    #renombrar columnas
    tabla_sorted= tabla_sorted.reset_index()
    tabla_sorted = tabla_sorted.rename(columns={'index': 'F.Emision'})
    tabla_sorted = tabla_sorted.rename(columns={'monto_total': 'MontoTotal(S/.)'}) 
    tabla_sorted['MontoTotal(S/.)'] = tabla_sorted['MontoTotal(S/.)'].apply(lambda x: '{:,.2f}'.format(x))
    
    
    #-----------------------------------------------------------------------
    tabladol = df1.pivot_table(index='emision', values='monto_total', aggfunc='sum')
    
    tabladol = tabladol.sort_values('emision', ascending=True)

    tabladol.index = pd.to_datetime(tabladol.index)
    tabladol.index = tabladol.index.strftime('%d-%m-%y')

    suma_columnadol = tabladol['monto_total'].sum()
    nueva_filadol = pd.DataFrame({'monto_total': suma_columnadol}, index=['Total'])
    print(tabladol)
    tabla_sorteddol = pd.concat([tabladol,nueva_filadol])
    #renombrar columnas
    tabla_sorteddol= tabla_sorteddol.reset_index()
    tabla_sorteddol = tabla_sorteddol.rename(columns={'index': 'F.Emision'})
    tabla_sorteddol = tabla_sorteddol.rename(columns={'monto_total': 'MontoTotal($)'}) 
    tabla_sorteddol['MontoTotal($)'] = tabla_sorteddol['MontoTotal($)'].apply(lambda x: '{:,.2f}'.format(x))
    print(tabladol)
    #-----------------------------------------------------------------------


    html = """\
    <html>
    <head></head>
    <body>
        <h2 style='text-align: left; color : #5499C7;'>TRADING FASHION LINE S.A.</h2>
        <p>Fecha : {0}</p>
        <h3>REPORTE DE VENTAS</h3>
        <p> Documentos en SOLES</p>
        {1} 
        <br>
        <p>Documentos en DOLARES</p>
        {2}
        <br>
        <h3>VENTAS POR DIA </h3>
        {5}
        <br>
        <p style='text-align: center; color: #999999;'>Nota: Este correo ha sigo generado automaticamente,no responder a este correo</p>
        <h2 style='text-align: center; color: #999999;'>© 2023 TI RETAIL TFL. All rights reserved.</h2>
    </body>
    </html>
    """
    #.format((datetime.now()).strftime("%d-%m-%Y %H:%M"),tabla_pivote_with_total.to_html(index=False),tabla_sorted.to_html(index=False))
    current_datetime = datetime.now().strftime("%d-%m-%Y %H:%M")
    tabla_pivote_html = tabla_pivote_with_total.to_html(index=False).replace('<td>', '<td style="text-align:right;">')
    tabla_pivote_dol_html = tabla_pivote_with_totaldol.to_html(index=False).replace('<td>', '<td style="text-align:right;">')
    tabla_sorted_html = tabla_sorted.to_html(index=False).replace('<td>', '<td style="text-align:right;">')
    tabla_sorted_dol_html = tabla_sorteddol.to_html(index=False).replace('<td>', '<td style="text-align:right;">')

    tabla_final = tabla_sorted.merge(tabla_sorteddol, on='F.Emision', how='left')
    #tabla_final = pd.concat([tabladol,tabla], axis=1)
    tabla_final_html = tabla_final.to_html(index=False).replace('<td>', '<td style="text-align:right;">')

    formatted_html = html.format( current_datetime,tabla_pivote_html, tabla_pivote_dol_html,tabla_sorted_html,tabla_sorted_dol_html,tabla_final_html)
                

    part1 = MIMEText(formatted_html, 'html')
    message= part1
    message['Subject'] = Header("Reporte de ventas",'utf-8')

    server=smtplib.SMTP('191.1.64.16',25)
    server.starttls()
    server.login('alertaprocesosti@topitop.com.pe','alertpro@.23*')
    destinatarios = ['dsalazar@topitop.com.pe','cbravo@topitop.com.pe','carenas@topitop.com.pe','elezama@topitop.com.pe','pgarcia@topitop.com.pe','rramosc@topitop.com.pe']
    #destinatarios = ['dsalazar@topitop.com.pe','cbravo@topitop.com.pe']
    #destinatarios = ['dsalazar@topitop.com.pe']

    server.sendmail('alertaprocesosti@topitop.com.pe',destinatarios,message.as_string())

    server.quit()

reporte()

sys.exit()