# sap_facturar_hitos_new.py
# Fecha de creación: 09/03/2026
# Autor: santiago.perezalbarran@telefonica.com
# Descripción: Automatiza el cambio de "fecha real" de hitos de un proyecto en SAP Fiori (ZHITOS)

import os
import time
from datetime import datetime

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ====== Config ======
Fiori_URL = ("https://fm21global.tg.telefonica/fiori"
             "?sap-client=550&sap-language=ES#ZOBJ_Z_GESTION_HITOS_0001-display"
             "?sap-ie=edge&sap-theme=sap_belize&sap-touch=0")
TIMEOUT = 30  # segundos

# ====== Utilidades ======
def wait(driver, by, selector, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )

def wait_clickable(driver, by, selector, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, selector))
    )

def safe_click(element):
    element.location_once_scrolled_into_view
    element.click()

def switch_to_last_tab(driver):
    driver.switch_to.window(driver.window_handles[-1])

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

# ====== Core ======
def init_driver(headless=False):
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    if headless:
        options.add_argument("--headless=new")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def login_fiori(driver, user, password):
    log("Abriendo Fiori…")
    driver.get(Fiori_URL)

    # TODO: Ajusta estos selectores al formulario de login real de tu Fiori/IdP SSO.
    # Si tenéis SSO, esta sección puede no ser necesaria (y el portal te logará directamente).
    try:
        log("Esperando campos de login… (si hay SSO, saltará)")
        user_input = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "USERNAME_FIELD_ID"))
        )
        pass_input = driver.find_element(By.ID, "PASSWORD_FIELD_ID")
        login_btn = driver.find_element(By.ID, "LOGIN_BUTTON_ID")

        user_input.clear(); user_input.send_keys(user)
        pass_input.clear(); pass_input.send_keys(password)
        safe_click(login_btn)
        time.sleep(3)
        log("Login enviado.")
    except Exception:
        log("No se mostró login clásico (posible SSO o login ya hecho). Continuando…")

def buscar_proyecto(driver, proyecto):
    log(f"Buscando proyecto: {proyecto}")
    # Preferible localizar por etiqueta/placeholder/texto; evita IDs volátiles.
    # TODO: Ajusta al input del filtro "Proyecto".
    # Ejemplo por aria-label:
    try:
        project_input = wait(driver, By.XPATH, "//input[@aria-label='Proyecto' or @title='Proyecto']")
    except Exception:
        # Fallback por ID conocido (menos estable)
        project_input = wait(driver, By.ID, "M0:46:::1:34-r")

    project_input.clear()
    project_input.send_keys(proyecto)

    # Botón Ejecutar/Buscar
    # TODO: Ajusta al botón real por texto o aria-label
    try:
        ejecutar_btn = wait_clickable(driver, By.XPATH, "//button[.//span[normalize-space(text())='Ejecutar'] or @aria-label='Ejecutar']")
    except Exception:
        ejecutar_btn = wait_clickable(driver, By.ID, "M0:50::btn[8]-r")

    safe_click(ejecutar_btn)

    # Muchas apps abren detalle en nueva pestaña/ventana
    time.sleep(2)
    if len(driver.window_handles) > 1:
        switch_to_last_tab(driver)

    # Espera a la tabla de resultados/hitos
    # TODO: Ajusta a la tabla concreta (Grid Table de UI5)
    wait(driver, By.XPATH, "//*[contains(@class,'sapUiTable') or contains(@class,'sapMListTbl')]")
    log("Resultados del proyecto cargados.")

def seleccionar_hitos(driver, codigos_hito):
    """
    Marca hitos en la tabla principal según su código visible.
    """
    log(f"Seleccionando {len(codigos_hito)} hitos…")
    for codigo in codigos_hito:
        # Busca la fila que contiene el código del hito
        # TODO: Ajustar columna/estructura. Se usa búsqueda genérica por texto en una fila.
        xpath_fila = f"//tr[.//*[contains(normalize-space(text()), '{codigo}')]]"
        try:
            fila = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_fila)))
            fila.location_once_scrolled_into_view
            time.sleep(0.2)

            # Si hay checkbox de selección:
            try:
                checkbox = fila.find_element(By.XPATH, ".//input[@type='checkbox' or contains(@class,'sapMCb')]")
                safe_click(checkbox)
            except Exception:
                # Como alternativa, click en la fila
                safe_click(fila)

            log(f"✔ Hito {codigo} seleccionado")
        except Exception:
            log(f"⚠ No encontré el hito {codigo}. Revisa el texto visible/columna.")

def abrir_editar_hitos(driver):
    """
    Pulsa el botón 'Editar hito' (o equivalente).
    """
    # TODO: Ajusta el botón por texto/aria-label
    try:
        editar_btn = wait_clickable(driver, By.XPATH, "//button[.//span[normalize-space(text())='Editar hito'] or @aria-label='Editar hito']")
    except Exception:
        editar_btn = wait_clickable(driver, By.ID, "M0:48::btn[25]-r")

    safe_click(editar_btn)
    log("Entrando en modo edición de hitos…")
    time.sleep(1)

def cambiar_fecha_real(driver, codigo_hito, nueva_fecha_iso="2026-03-09"):
    """
    Cambia la fecha real de UN hito ya en modo edición.
    - Busca el hito en la tabla editable y escribe la fecha.
    """
    # TODO: Ajustar la tabla editable y la columna de fecha real
    # Muchos grids UI5 usan celdas input. Intentamos localizar por encabezado de columna.
    try:
        # Encuentra índice de la columna 'Fecha real' si existe en encabezados
        # (Si no, accede por ID fijo 'grid#C134#0,7#cp1' como fallback).
        celda_fecha = None

        # Primera, ubica la fila del hito dentro del grid de edición:
        xpath_fila_edit = f"//tr[.//*[contains(normalize-space(text()), '{codigo_hito}')]]"
        fila_edit = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_fila_edit)))
        fila_edit.location_once_scrolled_into_view
        time.sleep(0.2)

        # Intenta localizar input de fecha dentro de la fila
        posibles_inputs = fila_edit.find_elements(By.XPATH, ".//input[contains(@class,'sapMInputBaseInner') and (@type='text' or @type='date')]")
        if not posibles_inputs:
            # Fallback: ID directo conocido (menos estable)
            try:
                celda_fecha = driver.find_element(By.ID, "grid#C134#0,7#cp1")
            except Exception:
                pass
        else:
            celda_fecha = posibles_inputs[-1]  # Suele estar al final si es columna derecha

        if celda_fecha is None:
            raise Exception("No encontré el input de fecha para el hito.")

        # Escribir fecha (formato depende del control; si es texto, escribe con formato local)
        # Intenta limpiar y escribir:
        celda_fecha.click()
        celda_fecha.send_keys(Keys.CONTROL, "a")
        celda_fecha.send_keys(Keys.DELETE)
        celda_fecha.send_keys(nueva_fecha_iso)  # Ej: 2026-03-09
        celda_fecha.send_keys(Keys.TAB)
        log(f"   ↳ Fecha real de {codigo_hito} → {nueva_fecha_iso}")

    except Exception as e:
        log(f"⚠ Error cambiando fecha de {codigo_hito}: {e}")

def guardar_cambios(driver):
    # TODO: Ajusta por texto/aria-label
    try:
        guardar_btn = wait_clickable(driver, By.XPATH, "//button[.//span[normalize-space(text())='Guardar'] or @aria-label='Guardar']")
    except Exception:
        guardar_btn = wait_clickable(driver, By.ID, "M0:50::btn[11]-r")
    safe_click(guardar_btn)
    log("Guardando cambios…")
    # Espera notificación/Toast de éxito si existe
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'sapMMessageToast')]")))
        log("✔ Cambios guardados (toast detectado).")
    except Exception:
        log("No se detectó toast; verifica visualmente o logs de app.")

def logout_y_cerrar(driver):
    try:
        # Si hay botón de usuario y logout en Fiori:
        # TODO: ajusta selectores si quieres hacer logout explícito
        pass
    finally:
        time.sleep(1)
        driver.quit()

# ====== Lectura de Excel/CSV (opcional) ======
def cargar_hitos_desde_excel(ruta_excel):
    import pandas as pd
    df = pd.read_excel(ruta_excel, engine="openpyxl")
    # Esperados: columnas 'proyecto', 'codigo_hito', 'nueva_fecha' (YYYY-MM-DD)
    df['nueva_fecha'] = pd.to_datetime(df['nueva_fecha']).dt.strftime('%Y-%m-%d')
    return df

# ====== Main ======
def main():
    load_dotenv()
    user = os.getenv("FM21_USER2")
    pwd = os.getenv("FM21_PASS2")
    if not user or not pwd:
        raise RuntimeError("Faltan variables FM21_USER2 / FM21_PASS2 en .env")

    driver = init_driver(headless=False)

    try:
        login_fiori(driver, user, pwd)

        # === Caso 1: un proyecto y varios hitos en código ===
        proyecto = "NÚMERO_O_NOMBRE_PROYECTO"  # TODO
        codigos_hito = ["HITO001", "HITO002"]  # TODO
        fecha_para_todos = "2026-03-09"        # TODO (ISO)

        buscar_proyecto(driver, proyecto)
        seleccionar_hitos(driver, codigos_hito)
        abrir_editar_hitos(driver)
        for ch in codigos_hito:
            cambiar_fecha_real(driver, ch, fecha_para_todos)
        guardar_cambios(driver)

        # === Caso 2 (opcional): múltiples proyectos/hitos desde Excel ===
        # df = cargar_hitos_desde_excel("hitos.xlsx")
        # for proyecto, chunk in df.groupby('proyecto'):
        #     buscar_proyecto(driver, proyecto)
        #     codigos = chunk['codigo_hito'].tolist()
        #     seleccionar_hitos(driver, codigos)
        #     abrir_editar_hitos(driver)
        #     for _, row in chunk.iterrows():
        #         cambiar_fecha_real(driver, row['codigo_hito'], row['nueva_fecha'])
        #     guardar_cambios(driver)

    except Exception as e:
        log(f"ERROR: {e}")
    finally:
        logout_y_cerrar(driver)

if __name__ == "__main__":
    main()