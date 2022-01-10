from datetime import date
import pandas as pd
#from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import win32com.client as win32
from datetime import datetime
import sys
import os

def get_comp_teso():
    usuario = ""
    contraseña = ""
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
    n_reg = len(df)
    print(f"Se obtuvieron {n_reg} registros de la página")
    try:
        df.to_excel("comprobantes.xlsx",index=False)
        print("Archivo comprobantes.xlsx creado exitosamente")
        return 1
    except Exception as e:
        if e.startwith("[Errno 13] Permission denied"):
            print("El archivo comprobantes.xlsx está abierto, cierrelo y corra de nuevo el programa")
            sys.exit(0)

def enviar_mail(comprobantes_ok):
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
            attachment = os.path.join(cwd,"comprobantes.xlsx")
            mail.Attachments.Add(attachment)
            mail.Send()
            print("Email de comprobantes enviado exitosamente!")
        except:
            print("Error al enviar el mail de los comprobantes, comprobar acceso a red y reintentar")

comprobantes_ok = get_comp_teso()# si retorna 1, no hay error y se envia el 
enviar_mail(comprobantes_ok)
