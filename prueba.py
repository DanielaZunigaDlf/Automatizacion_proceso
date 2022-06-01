from email import message
from openpyxl import load_workbook
from selenium import webdriver
import smtplib
from email.message import EmailMessage
import os

class ControlBrowser:
    def __init__(self, driver_path="chromedriver.exe"):
        """Crea una clase para controlar el navegador.
        - driver_path: Ruta de la ubicación del driver de chrome
        """
        self.driver_path = driver_path   
        self.driver = None    #es None porque lo creo despues

    def send_data(self, xpath, value):          
        """
        Escribe sobre un input según el xpath indicado.
        - xpath: xpath del input donde se escribirá
        - value: valor que se enviará al input
        return None
        """
        element = self.driver.find_element_by_xpath(xpath)
        element.send_keys(value)


    def open_browser(self, url):
        self.driver = webdriver.Chrome(self.driver_path)   
        self.driver.maximize_window() #maximiza la ventana
        self.driver.get(url)  #direccion abierta


class Gmail:
    def __init__(self, user, password):
        self.server = smtplib.SMTP_SSL(host="smtp.gmail.com", port=465) #smpt es el protocolo para poder enviar correos, en host va el servidor de gmail y en port el puerto en el cual trabaja
        self.server.login(user, password)
    
    def send_mail():



fileName = r"C:\Users\Usuario\Downloads\Base Seguimiento Observ Auditoría al_30042021.xlsx"

wb = load_workbook(fileName, data_only=True)  # data_only se usa para que cargue los valores de la hoja y no las formulas

sheet = wb.active  #obtengo la hoja /sheet = hoja


for row in sheet.rows:   #row = fila
    if row[0].value is None:
        continue

    proceso, observacion, riesgo, severidad, plan, fecha, responsable, area, correo, estado = row[:-1]   
    print(proceso.value)


    if estado.value == "Regularizado":   # con .value obtengo el valor de la celda porque estado es un objeto celda
        browser = ControlBrowser() # Creo el objeto de la clase ControBrowser
        browser.open_browser("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG") #Abre el navegador en la url que indico
        # driver = webdriver.Chrome("chromedriver.exe")   #abrir el navegador
        # driver.maximize_window() #maximiza la ventana
        # driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")  #direccion abierta
        
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[3]/div/select", proceso.value)
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[4]/div/input", riesgo.value)
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[5]/div/select", severidad.value)
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[6]/div/input", responsable.value)
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[7]/div/input", fecha.value.strftime("%d-%m-%Y"))
        browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[8]/div/textarea", observacion.value)
    

        # select_proceso = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[3]/div/select")   
        # select_proceso.send_keys(proceso.value)    # .value obtiene el valor de la celda proceso, .send_keys escribe el valor que obtenga de .value
        # input_riesgo = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[4]/div/input")
        # input_riesgo.send_keys(riesgo.value)  
        # select_severidad = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[5]/div/select")  
        # select_severidad.send_keys(severidad.value)
        # input_responsable = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[6]/div/input")  
        # input_responsable.send_keys(responsable.value)
        # input_fecha = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[7]/div/input")
        # input_fecha.send_keys(fecha.value.strftime("%d-%m-%Y"))  # .strftime me escribe la fecha como string (porque llega como un objeto datetime) en el formato que le pido
        # input_observacion = driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[1]/div/div[8]/div/textarea")
        # input_observacion.send_keys(observacion.value)
        button_enviar_formulario = browser.driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[2]/div/div/button")

        button_enviar_formulario.click()
        browser.driver.close() #cerrar el navegador
        

    elif estado.value == "Atrasado":
        # server = smtplib.SMTP_SSL(host="smtp.gmail.com", port=465)     #smpt es el protocolo para poder enviar correos, en host va el servidor de gmail y en port el puerto en el cual trabaja
        # server.login("correo.prueba2136@gmail.com", password_mail)  #conexion al mail
        password_mail = os.environ["PASSWORD_MAIL"]
        gmail = Gmail("correo.prueba2136@gmail.com", password_mail)

        message = EmailMessage()
        message['Subject'] = "Proceso por regularizar"
        message['From'] = "correo.prueba2136@gmail.com" 
        message['To'] = correo.value

        content = f"""Hola {responsable.value}, adjunto datos para regularizar el proceso:
        Proceso: {proceso.value}
        Estado: {estado.value}
        Observación: {observacion.value}
        Fecha: {fecha.value.strftime("%d-%m-%Y")}

        Saludos. """
        message.set_content(content)   # agrega el contenido de arriba al mensaje

        server.send_message(message)  #envia el correo
        server.quit()
        break


if __name__ == "__main__":
    pass
