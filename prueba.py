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

    def send_data(self, xpath:str, value:str|int):          
        """
        Escribe sobre un input según el xpath indicado.
        - xpath: xpath del input donde se escribirá
        - value: valor que se enviará al input
        return None
        """
        element = self.driver.find_element_by_xpath(xpath)
        element.send_keys(value)


    def open_browser(self, url:str):
        self.driver = webdriver.Chrome(self.driver_path)   
        self.driver.maximize_window() #maximiza la ventana
        self.driver.get(url)  #direccion abierta


class Gmail:
    def __init__(self, user:str, password:str):
        """
        Función de inicio de la clase Gmail
        - user : Correo de la cuenta gmail desde donde enviare correos
        - password : Clave del correo gmail
        """
        self.server = smtplib.SMTP_SSL(host="smtp.gmail.com", port=465) #smpt es el protocolo para poder enviar correos, en host va el servidor de gmail y en port el puerto en el cual trabaja
        self.server.login(user, password)   
        self.user = user    #aqui guardo el correo que le de
    
    def send_mail(self, content:str, subject:str, mail_dest:str):
        """
        Función que envia un correo
        - content : Es el contenido del mail que se enviara
        - subject : Es el asunto del correo
        - mail_dest : Es el correo al que se le enviara este mensaje 
        """
        message = EmailMessage()
        message['Subject'] = subject   #Asunto del correo
        message['From'] = self.user   #aqui ocupo el correo que ingreso en la clase
        message['To'] = mail_dest  #correo destinatario
        message.set_content(content)   # agrega el contenido al mensaje
        self.server.send_message(message)  #envia el correo
    
    def disconect(self):
        self.server.quit() #cierra la sesión con el servidor



if __name__ == "__main__":
    fileName = r"C:\Users\Usuario\Downloads\Base Seguimiento Observ Auditoría al_30042021.xlsx"  #ruta del excel con el que estoy trabajando

    wb = load_workbook(fileName, data_only=True)  # load_workbook lee el libro excel, data_only se usa para que cargue los valores de la hoja y no las formulas

    sheet = wb.active  #obtengo la hoja activa del excel    /sheet = hoja

    for row in sheet.rows:   #recorro cada fila en la hoja.  row = fila
        if row[0].value is None:   #valido que la primera celda de la fila tenga un dato, si es None continuo con el siguiente
            continue

        proceso, observacion, riesgo, severidad, plan, fecha, responsable, area, correo, estado = row[:-1]    #guardo en su respectiva variable cada celda como un objeto, ignorando la ultima celda (row[:-1])

        if estado.value == "Regularizado":   # con .value obtengo el valor de la celda porque estado es un objeto celda
            browser = ControlBrowser() # Creo el objeto de la clase ControBrowser
            browser.open_browser("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG") #Abre el navegador en la url que indico
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[3]/div/select", proceso.value)
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[4]/div/input", riesgo.value)
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[5]/div/select", severidad.value)
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[6]/div/input", responsable.value)
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[7]/div/input", fecha.value.strftime("%d-%m-%Y"))
            browser.send_data("/html/body/div/div/div/form/div/div[1]/div/div[8]/div/textarea", observacion.value)
            
            button_enviar_formulario = browser.driver.find_element_by_xpath("/html/body/div/div/div/form/div/div[2]/div/div/button")  #busca el boton donde luego hare click para enviar el formulario
            button_enviar_formulario.click()  #hace click en el boton
            browser.driver.close() #cerrar el navegador   

        elif estado.value == "Atrasado":
            password_mail = os.environ["PASSWORD_MAIL"]
            gmail = Gmail("correo.prueba2136@gmail.com", password_mail)  #creo el objeto gmail a partir de la clase Gmail que tiene las funciones para trabajar 
            content = f"""Hola {responsable.value}, adjunto datos para regularizar el proceso:
            Proceso: {proceso.value}   
            Estado: {estado.value}
            Observación: {observacion.value}
            Fecha: {fecha.value.strftime("%d-%m-%Y")}

            Saludos. """  #.value me obtiene el valor de la celda 
            gmail.send_mail(content, "Proceso por regularizar", correo.value) #envia el correo
            
            break   #esta de PRUEBA para no mandar todos los mails de la planilla!!!!
