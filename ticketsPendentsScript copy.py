from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
import smtplib
from email.message import EmailMessage
import pandas as pd
import datetime



# REQUISITOS:
# Instalar ChromeDriver
# instalar selenium webdriver pip install selenium
# pip install pandas openpyxl



# Advertencias:
# En el PATH debe estar el ChromeDriver correspondiente a la versión del Chrome que estás utilizando
# El Python debe de estar en el PATH para que el script pueda ejecutarlo
# Este script utiliza el navegador Chrome por defecto
# No debe de existir ningun fichero Jira Exportar Excel CSV (campos filtrados) en la carpeta de descargas
# El script puede dejar de funcionar si los desarolladores de JIRA cambian el front-end
# Tener version de python 3.13 o superior
# Version de ChromeDriver debe ser la versión 113 o superior



#Automatizacion de hacer login y descargar el CSV#

####################################################################################################################################

####Parametros de configuracion#####

# Usuario y contraseña de JIRA

usuario = "xxxx"
contra = "xxx"

# Ruta de descargas y prefijo para el nombre del archivo

ruta_descargas = r"C:\Users\auxinfor\Downloads"
prefijo = "Jira Exportar Excel CSV (campos filtrados)"
tipo_mime = "text/csv"


#############outlook##############
remitente = "xxxx@outlook.com"
contrasena = "dbkczdycgtyfxkrx"

destinatario = "ajtorhm@gmail.com"
servidor_smtp = "smtp.office365.com"
puerto_smtp = 587


####################################################################################################################################




# Iniciar el navegador chrome (se podria cambiar de navegador)
driver = webdriver.Chrome()
#driver = webdriver.Edge()

# Ir a la página de login
driver.get("https://id.atlassian.com/login?continue=https%3A%2F%2Fid.atlassian.com%2Fjoin%2Fuser-access%3Fresource%3Dari%253Acloud%253Ajira%253A%253Asite%252F090ef366-082f-4102-8997-0c8372386905%26continue%3Dhttps%253A%252F%252Fdecalesp.atlassian.net%252F&application=jira")
wait = WebDriverWait(driver, 100)



# Esperar que aparezca el campo de usuario
campo_usuario = wait.until(EC.presence_of_element_located((By.NAME, "username")))
campo_usuario.send_keys(usuario)

# Continuar al siguiente paso (como hace el sitio)
campo_usuario.send_keys(Keys.RETURN)


time.sleep(2)
# Esperar que cargue el campo de contraseña
campo_contrasena = wait.until(EC.element_to_be_clickable((By.NAME, "password")))
campo_contrasena.send_keys(contra)
campo_contrasena.send_keys(Keys.RETURN)

# Esperar que cargue la lista de proyectos
wait.until(EC.url_contains("jira/your-work"))

time.sleep(2)
# Clic en el boton marcado
marcado =  wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Marcado']")))
marcado.click()

time.sleep(1)
#Clic en el botón de filtro pendent
filtro =  wait.until(EC.element_to_be_clickable((By.XPATH, "//a[.//h4[text()='Tickets Pendent']]")))
filtro.click()


time.sleep(3)
#Clic en el botón de exportar
exportar =  wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Exportar']]")))
exportar.click()


time.sleep(4)
#Clic en el botón de exportar Excel
exportar =  wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Exportar Excel CSV (campos filtrados)']]")))
exportar.click()


max_intentos = 5
intento = 0

while intento < max_intentos:
    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//h2[.//span[text()='Exportación terminada']]"))
        )
        print("✅ Exportación exitosa")
        break
    except:
        print(f"Reintentando exportación... (Intento {intento + 1})")

        # Cancelar exportación
        try:
            driver.find_element(By.XPATH, "//button[.//span[text()='Cancelar la exportación']]").click()
        except: pass
        time.sleep(1)

        # Hacer clic en los botones para exportar
        try:
            driver.find_element(By.XPATH, "//button[.//span[text()='Exportar']]").click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//button[.//span[text()='Exportar Excel CSV (campos filtrados)']]").click()
        except: pass
        time.sleep(2)

        intento += 1

else:
    print("❌ No se logró exportar después de varios intentos.")

driver.quit()


#time.sleep(3)
#driver.quit()




####################################################################################################################################

hora_actual = time.strftime("%H:%M:%S")
fecha = time.strftime("%d/%m/%Y")
now = datetime.datetime.now()
mes = now.strftime("%B")
any = now.strftime("%Y")


patron = os.path.join(ruta_descargas, f"{prefijo}*.csv")
archivos = glob.glob(patron)

if len(archivos) == 1:
    ruta_archivo = archivos[0]
    nombre_archivo = os.path.basename(ruta_archivo)
    print("Archivo encontrado:", ruta_archivo)

    # Leer el CSV y eliminar comas del contenido
    df = pd.read_csv(ruta_archivo, dtype=str)  # Leemos todo como texto para evitar problemas
    df = df.applymap(lambda x: x.replace(",", " ") if isinstance(x, str) else x)

    # Guardar como Excel
    ruta_excel = ruta_archivo.replace(".csv", ".xlsx")
    df.to_excel(ruta_excel, index=False)
    print("Archivo convertido a Excel:", ruta_excel)

    # Crear el mensaje de correo
    msg = EmailMessage()
    msg['Subject'] = f'Tickets Pendents Resum de {mes} {any}' 
    msg['From'] = remitente
    msg['To'] = destinatario
    msg.set_content(f"Archiu extret de JIRA: {fecha} - {hora_actual} (Formato Excel)")

    # Adjuntar el archivo Excel
    with open(ruta_excel, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=os.path.basename(ruta_excel)
        )

    # Enviar el correo
    try:
        with smtplib.SMTP(servidor_smtp, puerto_smtp) as smtp:
            smtp.starttls()
            smtp.login(remitente, contrasena)
            smtp.send_message(msg)
            print("Correo enviado con éxito.")
    except Exception as e:
        print("Error al enviar el correo:", e)

    # Borrar archivos
    os.remove(ruta_archivo)
    os.remove(ruta_excel)
    print("Archivos eliminados.")
else:
    print("No se encontró exactamente un archivo con ese prefijo.")

print("Ejecución finalizada.")

