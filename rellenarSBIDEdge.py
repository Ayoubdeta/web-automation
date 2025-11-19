from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

# 🔹 RUTA A TU msedgedriver.exe
edge_driver_path = "C:\\Herramientas\\edgedriver\\msedgedriver.exe"
service = Service(edge_driver_path)

# Iniciar Edge
driver = webdriver.Edge(service=service)
wait = WebDriverWait(driver, 10)

usuario = "rthe9fas"
contrasena = "ElPrat08820*"

driver.get("https://www.empresaiformacio.org/sBid")
time.sleep(1)

# Entrar al iframe de login
wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "jspContainer")))

# Aceptar cookies
try:
    accept_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Acceptar totes')]")))
    accept_btn.click()
    print("Cookies aceptadas")
except Exception as e:
    print("No se pudo aceptar cookies:", e)

# Login
wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(usuario)
wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(contrasena)
wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='pageBody']/div[1]/div[2]/form/div[3]/div/button"))).click()

# Salir del iframe de login
driver.switch_to.default_content()
time.sleep(2)

#  Función recursiva para buscar y hacer clic en enlaces dentro de iframes
def buscar_enlace(driver, wait, enlace_xpath, profundidad=0):
    indent = "  " * profundidad
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for i, frame in enumerate(iframes):
        try:
            driver.switch_to.frame(frame)
            elems = driver.find_elements(By.XPATH, enlace_xpath)
            if elems:
                elems[0].click()
                print(f"{indent}Enlace clicado en iframe nivel {profundidad}")
                return True
            else:
                encontrado = buscar_enlace(driver, wait, enlace_xpath, profundidad + 1)
                if encontrado:
                    return True
        except Exception as e:
            print(f"{indent} Error en iframe nivel {profundidad}: {e}")
        finally:
            driver.switch_to.parent_frame()
    return False

enlace_xpath = "//a[contains(@href, 'ActivitatFct') and contains(., 'Activitat diària del dossier')]"
exito = buscar_enlace(driver, wait, enlace_xpath)
if not exito:
    print(" No se pudo hacer clic en el enlace.")
    driver.quit()
    exit()

time.sleep(3)

#  Función recursiva para buscar selects en iframes anidados
def buscar_y_cambiar_selects(driver, profundidad=0):
    indent = "  " * profundidad
    selects = driver.find_elements(By.TAG_NAME, "select")
    if selects:
        print(f"{indent} {len(selects)} selects encontrados en profundidad {profundidad}")
        for i, sel_elem in enumerate(selects):
            try:
                sel = Select(sel_elem)
                if i == len(selects) - 3:
                    sel.select_by_visible_text("30min")  # Último select
                    print(f"{indent} Select {i+1}: seleccionado '30min'")
                else:
                    sel.select_by_visible_text("15min")
                    print(f"{indent} Select {i+1}: seleccionado '15min'")
            except Exception as e:
                print(f"{indent} Error seleccionando en select {i+1}: {e}")
        return True
    else:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"{indent} Buscando selects en {len(iframes)} iframe(s) a profundidad {profundidad}")
        for i, frame in enumerate(iframes):
            try:
                driver.switch_to.frame(frame)
                encontrado = buscar_y_cambiar_selects(driver, profundidad + 1)
                if encontrado:
                    return True
            except Exception as e:
                print(f"{indent} Error accediendo iframe: {e}")
            finally:
                driver.switch_to.parent_frame()
    return False

exito_select = buscar_y_cambiar_selects(driver)
if not exito_select:
    print(" No se encontró ningún select para cambiar.")

# Función recursiva para buscar y clicar botón 'Emmagatzemar'
def buscar_y_clicar_boton_emmagatzemar(driver, profundidad=0):
    indent = "  " * profundidad
    try:
        botones = driver.find_elements(By.XPATH, 
            "//span[@class='btn btn-info' and @title='Emmagatzemar' and contains(@onclick, 'save()')]")
        if botones:
            botones[0].click()
            print(f"{indent} Botón 'Emmagatzemar' clicado con éxito")
            return True
    except Exception as e:
        print(f"{indent} Error al intentar clicar el botón: {e}")

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"{indent} Buscando botón 'Emmagatzemar' en {len(iframes)} iframe(s) a profundidad {profundidad}")
    for frame in iframes:
        try:
            driver.switch_to.frame(frame)
            encontrado = buscar_y_clicar_boton_emmagatzemar(driver, profundidad + 1)
            if encontrado:
                return True
        except Exception as e:
            print(f"{indent}Error accediendo a iframe: {e}")
        finally:
            driver.switch_to.parent_frame()
    return False

exito_boton = buscar_y_clicar_boton_emmagatzemar(driver)
if not exito_boton:
    print("No se encontró el botón 'Emmagatzemar'.")

print("programa finalizado.")
time.sleep(2)
driver.quit()
