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
# 5. Guardar los cambios(otro click)
# 6. Cerrar sesión en SAP
=========================================================================================
'''
#pasos
#en la pagina de sap (https://fm21global.tg.telefonica/fiori?sap-client=550&sap-language=ES#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize&sap-touch=0)
# Escrives  en : (id="M0:46:::1:34-r") el nombre/numero del proyecto
# En id="M0:50::btn[8]-r" haces click para ejecutar 
# Pasaras a una nueva PESTAÑA donde se muestran los hitos del proyecto
# En la pestaña de hitos, buscas el hito/hitos que quieres modificar
#apareceran en una tabla todos los hitos del proyecto
# En la tabla de hitos, buscas el hito que quieres modificar (se buscan por codigo/nombre del hito)
#En la columna "x" de la tabla se clica el hito que quieres modificar (se pueden varios en este proceso)
#le clicas (una vez echo el paso anterior) a Modificar
#una vez dentro solo apareceran los hitos modificables
# aparecera un campo llamado modificar fecha real , le haces clik y automaticamente se modifican todas las fechas reales de los hitos seleccionados
# le haces click a Guardar
# IMPORTANTE: Asegurarse de que el WebDriver y el navegador estén actualizados y sean compatibles.
# Requiere instalar las siguientes librerías:
# pip install selenium webdriver-manager
# pip install selenium-stealth
# pip install selenium-wire
# pip install undetected-chromedriver
xcx
# pip install pyautogui

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
        
        # Esperar a que la página cargue completamente
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "M0:46:::1:34-r")))
        
        # Ingresar el nombre/numero del proyecto
        project_input = driver.find_element(By.ID, "M0:46:::1:34-r")
        project_input.send_keys("Nombre o numero del proyecto")  # Reemplazar con el valor real
        
        # Hacer click en el botón de ejecutar
        execute_button = driver.find_element(By.ID, "M0:50::btn[8]-r")
        execute_button.click()
        
        # Esperar a que se abra la nueva pestaña con los hitos del proyecto
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "sapMListTbl")))
        
        # Buscar el hito/hitos que quieres modificar (reemplazar con el criterio de búsqueda real)
        milestone_code = "Codigo del hito"  # Reemplazar con el código real del hito
        milestone_row = driver.find_element(By.XPATH, f"//tr[td[contains(text(), '{milestone_code}')]]")
        
        # Seleccionar el hito haciendo click en la columna "x"
        select_checkbox = milestone_row.find_element(By.XPATH, ".//td[1]//input[@type='checkbox']")
        select_checkbox.click()
        
        # Hacer click en Modificar
        modify_button = driver.find_element(By.ID, "M0:50::btn[9]-r")  # Reemplazar con el ID correcto del botón Modificar
        modify_button.click()
