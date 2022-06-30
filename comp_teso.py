from datetime import date
import pandas as pd
#from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
import win32com.client as win32
from datetime import datetime
import sys
import os
import random

def get_comp_teso():
    print("En ejecución... espere por favor")
    usuario = "sadmin"
    contraseña = "sadmin00"
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-download-notification')
    options.add_argument('--headless')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    s=Service('../chromedriver.exe')
    driver = webdriver.Chrome(service=s,options=options)
    driver.get("http://tesoreria.mecontuc.gov.ar/")
    wait = WebDriverWait(driver, 20)
    tabs = set(driver.window_handles)
    #Ingresar usuario
    try:
        input_cuit = wait.until(EC.element_to_be_clickable((By.ID, 'txtUser')))
        input_cuit.clear()
        input_cuit.send_keys(str(usuario))
        #Ingresar contraseña
        input_cuit = wait.until(EC.element_to_be_clickable((By.ID, 'txtPass')))
        input_cuit.clear()
        input_cuit.send_keys(str(contraseña))
        input_cuit.send_keys(Keys.ENTER)
        wait.until(EC.number_of_windows_to_be(len(tabs)+1))
        newWindow = (set(driver.window_handles) - tabs).pop()
        driver.switch_to.window(newWindow)
        tabla_mail = wait.until(EC.element_to_be_clickable((By.ID, 'GridView1')))
        html = tabla_mail.get_attribute('outerHTML')
        driver.quit()
        df = pd.read_html(html)[0]
        df["Descargar"] = ("http://tesoreria.mecontuc.gov.ar/Comprobantes/"+df["Organismo"]+
            "/"+df["Tipo"]+"_"+df["Nombre Archivo"]).replace(" ","%20",regex=True)
        #f['Cta Destino'] = df['Cta Destino'].round(0)
        df['Cta Destino'] = df['Cta Destino'].fillna(0)
        df['Cta Destino'] = df['Cta Destino'].apply(int)
        df['Cta Destino'] = df['Cta Destino'].apply(str)
        lista_importe = []
        for i in df['Importe']:
            if type(i) is str:
                r = float(i.replace(".",""))/100
            else:
                r= float(i)
            lista_importe.append(r)
        df['Importe'] = pd.Series(lista_importe)

        #df['Importe'] = pd.to_numeric(df['Importe'],errors='ignore')
        df['Entrega'] = pd.to_datetime(df['Entrega'],errors='ignore',dayfirst=True)
        df['Comp.'] = pd.to_datetime(df['Comp.'],errors='ignore',dayfirst=True)
        n_reg = len(df)
        print(f"Se obtuvieron {n_reg} registros de la página")
    except TimeoutException:
        error = "Hubo un problema con la página de la Tesoreria, intente nuevamente."
        print(error)
        enviar_mail_error(error)
        sys.exit(0)
    try:
        archivo = "comprobantes.xlsx"
        writer = pd.ExcelWriter(archivo,
                datetime_format='dd/mm/yyyy',
                date_format='dd/mm/yyyy')
        df.to_excel(writer,index=False,sheet_name='Sheet1')
        writer.save() 
        
        print("Archivo comprobantes.xlsx creado exitosamente")
        if len(df)>0:
            return 1, archivo
    except Exception as e:
        if e.startwith("[Errno 13] Permission denied"):
            error = "El archivo comprobantes.xlsx está abierto, creando otro de resplado"
            aleatorio = str(random.randit(1,99))
            archivo = f"comprobantes_{aleatorio}.xlsx"
            writer = pd.ExcelWriter(archivo,
                datetime_format='dd/mm/yyyy',
                date_format='dd/mm/yyyy')
            df.to_excel(writer,index=False,sheet_name='Sheet1')
            writer.save() 
            return 1, archivo


def enviar_mail(comprobantes_ok,archivo):
    if comprobantes_ok==1:
        try:
            Outlook = win32.Dispatch('Outlook.application')
            mail = Outlook.CreateItem(0)
            #lista de emails a enviar en el txt
            with open("lista_mails.txt","r") as mails:
                lista_mails = mails.readlines()
                mails.close()

            mailTo=";".join(lista_mails)
            mail.To = mailTo
            hoy = datetime.today().strftime('%d-%m-%Y')
            mail.Subject = f'Comprobantes pagina - {hoy}'
            mail.Body = f'Comprobantes tesoreria día {hoy}'
            cwd = os.getcwd()
            attachment = os.path.join(cwd,archivo)
            mail.Attachments.Add(attachment)
            mail.Send()
            print("Email de comprobantes enviado exitosamente!")
        except:
            print("Error al enviar el mail de los comprobantes, comprobar acceso a red y reintentar")

def enviar_mail_error(error):
    Outlook = win32.Dispatch('Outlook.application')
    mail = Outlook.CreateItem(0)
    mailTo="cpulido@mecontuc.gov.ar;cveliz@mecontuc.gov.ar;fcarreno@mecontuc.gov.ar"
    mail.To = mailTo
    hoy = datetime.today().strftime('%d-%m-%Y')
    mail.Subject = f'Error Comprobantes pagina - {hoy}'
    mail.Body = f'Hubo un error al enviar los Comprobantes tesoreria día {hoy}\n{error}'
    mail.Send()
    print("Se envió un mail informando el error")

comprobantes_ok,archivo = get_comp_teso()# si retorna 1, no hay error y se envia el mail
enviar_mail(comprobantes_ok,archivo)