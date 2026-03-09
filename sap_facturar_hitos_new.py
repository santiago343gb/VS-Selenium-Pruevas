# =========================================================================================
# sap_facturar_hitos_new.py
# Fecha de creación: 09/03/2026
# Autor: santiago.perezalbarran@telefonica.com
#
# Script para cambiar la fecha real de hitos en SAP FIORI (ZHITOS)
#
# FUNCIONALIDAD:
# - Lee un EXCEL con las columnas:
#       proyecto | codigo_hito | fecha_real | editar | comentario
# - Para cada proyecto del Excel:
#       1. Entra en SAP ZHITOS
#       2. Busca el proyecto
#       3. Selecciona SOLO los hitos del Excel
#       4. Entra a Editar
#       5. Cambia SOLO la fecha de esos hitos
#       6. Guarda
#
# NOTAS IMPORTANTES:
# - Si en la fila del Excel 'fecha_real' va vacía → se usa la FECHA_GLOBAL
# - Si existe columna 'editar' → solo procesa los que tienen SI / 1 / TRUE / X
# - Este script usa SOLO time.sleep() para tiempos
# =========================================================================================

import os
import time
import pandas as pd
from datetime import datetime

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


# =============================
# CONFIGURACIÓN GLOBAL
# =============================
EXCEL_PATH = r"C:\Users\bt00092\Downloads\hitos.xlsx"   # Excel de entrada
FECHA_GLOBAL = "2026-03-31" # Fecha por defecto si no hay en Excel
USAR_COLUMNA_EDITAR = True  # True = respeta columna EDITAR


# =============================
# DRIVER
# =============================
def init_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# =============================
# LOGIN FIORI
# =============================
def login_fiori(driver, user, password):
    URL = ("https://fm21global.tg.telefonica/fiori"
           "?sap-client=550&sap-language=ES"
           "#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize")

    driver.get(URL)
    time.sleep(6)

    try:
        # IDs REALES SI NO HAY SSO
        user_input = driver.find_element(By.ID, "USERNAME_FIELD_ID")
        pass_input = driver.find_element(By.ID, "PASSWORD_FIELD_ID")
        login_btn = driver.find_element(By.ID, "LOGIN_BUTTON_ID")

        user_input.send_keys(user)
        pass_input.send_keys(password)
        login_btn.click()

        time.sleep(8)

    except:
        print("SSO detectado o login no requerido.")
        time.sleep(3)


# =============================
# BÚSQUEDA DEL PROYECTO
# =============================
def buscar_proyecto(driver, proyecto):
    time.sleep(5)

    try:
        campo = driver.find_element(By.XPATH, "//input[@aria-label='Proyecto']")
    except:
        campo = driver.find_element(By.ID, "M0:46:::1:34-r")

    campo.clear()
    campo.send_keys(proyecto)
    time.sleep(1)

    try:
        btn = driver.find_element(By.XPATH, "//button[.//span[text()='Ejecutar']]")
    except:
        btn = driver.find_element(By.ID, "M0:50::btn[8]-r")

    btn.click()
    time.sleep(8)

    # Cambio pestaña si la app abre nueva
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(4)


# =============================
# SELECCIONAR HITO EN LISTA
# =============================
def seleccionar_hito(driver, codigo_hito):
    time.sleep(3)
    try:
        fila = driver.find_element(By.XPATH, f"//tr[td[contains(text(), '{codigo_hito}')]]")
        fila.location_once_scrolled_into_view
        time.sleep(1)

        chk = fila.find_element(By.XPATH, ".//input[@type='checkbox']")
        chk.click()
        time.sleep(1)

        print(f"✔ Hito seleccionado: {codigo_hito}")
        return True

    except Exception as e:
        print(f"⚠ No se pudo seleccionar el hito {codigo_hito}: {e}")
        return False


# =============================
# ENTRAR EN MODO EDICIÓN
# =============================
def abrir_editar(driver):
    try:
        btn = driver.find_element(By.XPATH, "//button[.//span[text()='Editar hito']]")
    except:
        btn = driver.find_element(By.ID, "M0:48::btn[25]-r")

    btn.click()
    time.sleep(6)


# =============================
# CAMBIAR FECHA DE UN HITO
# =============================
def cambiar_fecha_real_hito(driver, codigo_hito, fecha):
    time.sleep(3)
    try:
        fila = driver.find_element(By.XPATH, f"//tr[td[contains(text(), '{codigo_hito}')]]")
        fila.location_once_scrolled_into_view
        time.sleep(1)

        campo_fecha = fila.find_element(By.XPATH, ".//input[contains(@class,'sapMInputBaseInner')]")

        campo_fecha.click()
        time.sleep(0.3)
        campo_fecha.send_keys(Keys.CONTROL, "a")
        campo_fecha.send_keys(Keys.DELETE)
        campo_fecha.send_keys(fecha)
        campo_fecha.send_keys(Keys.TAB)
        time.sleep(0.5)

        print(f"   ↳ Fecha aplicada a hito {codigo_hito}: {fecha}")
        return True

    except Exception as e:
        print(f"⚠ No se pudo cambiar la fecha del hito {codigo_hito}: {e}")
        return False


# =============================
# GUARDAR
# =============================
def guardar(driver):
    try:
        btn = driver.find_element(By.XPATH, "//button[.//span[text()='Guardar']]")
    except:
        btn = driver.find_element(By.ID, "M0:50::btn[11]-r")

    btn.click()
    time.sleep(6)
    print("✔ Cambios guardados.")


# =============================
# LECTURA DEL EXCEL
# =============================
def cargar_excel(path, usar_editar=True):
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [c.lower().strip() for c in df.columns]

    if "proyecto" not in df.columns or "codigo_hito" not in df.columns:
        raise Exception("El Excel DEBE incluir columnas: proyecto, codigo_hito")

    # Normalizar fecha_real
    if "fecha_real" in df.columns:
        df["fecha_real"] = pd.to_datetime(df["fecha_real"], errors="coerce").dt.strftime("%Y-%m-%d")
    else:
        df["fecha_real"] = None

    # Filtrar por columna EDITAR si existe
    if usar_editar and "editar" in df.columns:
        mask = df["editar"].astype(str).str.upper().isin(["SI", "1", "X", "TRUE"])
        df = df[mask].copy()

    if df.empty:
        raise Exception("No hay filas que procesar después del filtrado.")

    return df


# =============================
# MAIN
# =============================
def main():

    load_dotenv()
    USER = os.getenv("FM21_USER2")
    PASS = os.getenv("FM21_PASS2")

    driver = init_driver()

    try:
        login_fiori(driver, USER, PASS)

        df = cargar_excel(EXCEL_PATH, usar_editar=USAR_COLUMNA_EDITAR)

        for proyecto, grupo in df.groupby("proyecto"):

            print(f"\n========== PROYECTO {proyecto} ==========")
            buscar_proyecto(driver, proyecto)

            seleccionados = []

            for _, row in grupo.iterrows():

                codigo = str(row["codigo_hito"])
                fecha = row["fecha_real"] if pd.notna(row["fecha_real"]) else FECHA_GLOBAL

                if seleccionar_hito(driver, codigo):
                    seleccionados.append((codigo, fecha))

            if not seleccionados:
                print("⚠ No hitos seleccionados. Siguiente proyecto.")
                continue

            abrir_editar(driver)

            for codigo, fecha in seleccionados:
                cambiar_fecha_real_hito(driver, codigo, fecha)

            guardar(driver)

    except Exception as e:
        print("ERROR:", e)

    finally:
        time.sleep(3)
        driver.quit()


if __name__ == "__main__":
    main()
