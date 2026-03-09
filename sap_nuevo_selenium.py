#IMPORTS
'''
=========================================================================================
# sap_facturar_hitos_new.py
# Fecha de creacion: 09/03/2026
# Correo: santiago.perezalbarran@telefonica.com
# Script para cambiar la fecha real de un hito de un proyecto en SAP
# Proceso que se realiza:
# 1. Iniciar sesión en SAP
# 2. Navegar a la transacción ZHITOS (https://fm21global.tg.telefonica/fiori?sap-client=550&sap-language=ES#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize&sap-touch=0)
# 3. Buscar el proyecto y el hito
# 4. Cambiar la fecha real del hito (a traves de un click en el campo de fecha)
# 5. Guardar los cambios (otro click)
# 6. Cerrar sesión en SAP
=========================================================================================
'''

#pasos
#en la pagina de sap (https://fm21global.tg.telefonica/fiori?sap-client=550&sap-language=ES#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize&sap-touch=0)
# Escrives  en : (id="M0:46:::1:34-r") el nombre/numero del proyecto
# En id="M0:50::btn[8]-r" haces click para ejecutar 
# Pasaras a una nueva PESTAÑA donde se muestran los hitos del proyecto
# en id id="grid#C102#1,4" esta el hito en esa tabla
#en esa tabla se tendra que leer los hitos y selecionar lo que se quieran (entan en el power bi/excell)
# una vez encima del hito se dan 3 flechas hacia la derecha de tal manera que estaremos encima de marcar el hito
# se le dara a clik (es a la de hitos)
# si hay mas hitos se puede regresar a la columna hitos ir hacia abajo o hacia arriba (como si fueran flechas) encotrar dichos hitos y repetir del proceso.
# Se hara clik en d="M0:48::btn[25]-r" es el boton de editar hito
# una vez dado a editar nos iremos a la columna id="grid#C134#0,7#cp1" y se ira hacia bajo dando clik a cada hito selecionado anteriormente
#de esta manera se cambiara la fecha real del hito
# se ira al boton id="M0:50::btn[11]-r" que es el de guardar los cambios 
# y se le dara a clik
# IMPORTANTE: Asegurarse de que el WebDriver y el navegador estén actualizados y sean compatibles.
# Requiere instalar las siguientes librerías:
# pip install selenium webdriver-manager
# pip install selenium-stealth
# pip install selenium-wire
# pip install undetected-chromedriver
# pip install pyautogui


# =====================
# CARGAR VARIABLES .ENV
# =====================
from dotenv import load_dotenv
import os
load_dotenv()

sap_user = os.getenv("FM21_USER2")
sap_pass = os.getenv("FM21_PASS2")
# Asegúrate de que las variables de entorno FM21_USER2 y FM21_PASS2 estén definidas en tu archivo .env
# =====================
# IMPORTS SELENIUM
# =====================
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options


def main():
    # Configuración del navegador
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    
    # Inicializar el WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        # Abrir la página de SAP
        driver.get("https://fm21global.tg.telefonica/fiori?sap-client=550&sap-language=ES#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize&sap-touch=0")
        
        # ==========================
        # LOGIN AUTOMÁTICO EN SAP
        # ==========================
        # ⚠️ Debes sustituir estos IDs por los reales del formulario de SAP Fiori
        USER_ID = "USERNAME_FIELD_ID"
        PASS_ID = "PASSWORD_FIELD_ID"
        LOGIN_BTN_ID = "LOGIN_BUTTON_ID"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, USER_ID))
        )

        user_input = driver.find_element(By.ID, USER_ID)
        pass_input = driver.find_element(By.ID, PASS_ID)
        login_button = driver.find_element(By.ID, LOGIN_BTN_ID)

        user_input.send_keys(sap_user)
        pass_input.send_keys(sap_pass)
        login_button.click()

        # Esperar tras el login
        time.sleep(5)

        # ==========================
        # CONTINÚA TU SCRIPT NORMAL
        # ==========================

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "M0:46:::1:34-r"))
        )
        
        project_input = driver.find_element(By.ID, "M0:46:::1:34-r")
        project_input.send_keys("Nombre o numero del proyecto")  # Reemplazar
        
        execute_button = driver.find_element(By.ID, "M0:50::btn[8]-r")
        execute_button.click()
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "sapMListTbl"))
        )

        milestone_code = "Codigo del hito"  # Reemplazar
        milestone_row = driver.find_element(
            By.XPATH, f"//tr[td[contains(text(), '{milestone_code}')]]"
        )
        
        select_checkbox = milestone_row.find_element(
            By.XPATH, ".//td[1]//input[@type='checkbox']"
        )
        select_checkbox.click()
        
        modify_button = driver.find_element(By.ID, "M0:50::btn[9]-r")
        modify_button.click()

    except Exception as e:
        print("ERROR:", e)

    finally:
        time.sleep(5)
        driver.quit()


if __name__ == "__main__":
    main()