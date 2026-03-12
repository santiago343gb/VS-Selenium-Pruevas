###########################################################################################
# sap_facturar_hitos_new.py — MULTIPROYECTO + SESIÓN LIMPIA + BUSCADOR ROBUSTO + NOK/OK POR HITO
###########################################################################################

import os, re, time
import pandas as pd
from dotenv import load_dotenv
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==============================
# CONFIG
# ==============================
EXCEL_PATH = r"C:\Users\bt00092\Downloads\tabla_facturar.xlsx"
RESULTADO_PATH = r"C:\Users\bt00092\Downloads\resultado_hitos.xlsx"
CHROME_DRIVER_PATH = r"C:\Python Project\drivers\chromedriver.exe"

FAST_WAIT = 15
SLEEP_SHORT = 0.35
MAX_REINTENTOS = 3
RETRASO_ENTRE_REINTENTOS = 1.2

# Escaneo por PAGE DOWN solo como fallback (muchos hitos / render tardío)
MAX_PAGEDOWN_PASOS = 240       # ~ 240 pantallas; ajusta si necesitas más
SLEEP_PAGEDOWN = 0.18          # espera corta entre saltos

# ==============================
# UTILIDADES
# ==============================
def ensure_env():
    load_dotenv()
    u = os.getenv("FM21_USER2")
    p = os.getenv("FM21_PASS2")
    if not u or not p:
        raise Exception("Faltan credenciales")
    print(f"Usuario cargado: {u}")
    return u, p

def wait_no_busy(driver):
    try:
        WebDriverWait(driver, FAST_WAIT).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".sapUiBlockLayer, .sapUiLocalBusyIndicator")
            )
        )
    except:
        pass

def safe_type(driver, el, txt):
    try:
        el.clear()
        el.send_keys(txt)
    except:
        driver.execute_script("arguments[0].value='';", el)
        time.sleep(0.1)
        el.send_keys(txt)

def iniciar_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-software-rasterizer")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--remote-debugging-pipe")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    return webdriver.Chrome(service=Service(CHROME_DRIVER_PATH), options=opts)

# ==============================
# LOGIN
# ==============================
def login(driver, user, pwd):
    URL = (
        "https://fm21global.tg.telefonica/fiori"
        "?sap-client=550&sap-language=ES"
        "#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize"
    )
    driver.get(URL)
    time.sleep(1.2)
    driver.find_element(By.CSS_SELECTOR,"input[placeholder='Usuario']").send_keys(user)
    driver.find_element(By.CSS_SELECTOR,"input[placeholder='Clave de acceso']").send_keys(pwd)
    driver.find_element(By.XPATH,"//button[contains(text(),'Acceder')]").click()
    time.sleep(1.2)
    print("✔ Login OK")

# ==============================
# EJECUTAR PROYECTO
# ==============================
def ejecutar_proyecto(driver, proyecto):
    wait = WebDriverWait(driver, FAST_WAIT)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.XPATH,"//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]")
    ))

    campo = wait.until(EC.presence_of_element_located(
        (By.XPATH,"//input[@title='Definición del proyecto']")
    ))
    safe_type(driver, campo, proyecto)
    time.sleep(SLEEP_SHORT)

    try:
        sug = WebDriverWait(driver,3).until(
            EC.element_to_be_clickable((By.XPATH,"//ul/li[1]"))
        )
        sug.click()
    except:
        try:
            campo.send_keys(Keys.ENTER)
        except:
            pass

    print("Buscando EJECUTAR…")
    try:
        btn = WebDriverWait(driver,2).until(
            EC.element_to_be_clickable((By.XPATH,"//*[normalize-space()='Ejecutar']/ancestor::button"))
        )
        driver.execute_script("arguments[0].click();", btn)
    except:
        ActionChains(driver).send_keys(Keys.F8).perform()

    driver.switch_to.default_content()
    wait_no_busy(driver)
    print("✔ Proyecto ejecutado correctamente")

# ==============================
# BUSCADOR ROBUSTO DE CELDA DEL HITO
# ==============================
def _buscar_celda_hito(driver, hito: str):
    """
    Intenta localizar la celda del hito con distintos patrones sin scroll.
    Devuelve WebElement o None.
    """
    patrones = [
        f"//span[starts-with(@id,'grid#') and contains(@id,'#if') and contains(normalize-space(.), '{hito}')]",
        f"//td[starts-with(@id,'grid#') and contains(@id,'#it') and contains(normalize-space(.), '{hito}')]",
        f"//*[self::span or self::td or self::a][contains(normalize-space(.), '{hito}')]",
    ]
    for xp in patrones:
        els = driver.find_elements(By.XPATH, xp)
        if els:
            return els[0]
    return None

def _focus_tabla(driver):
    """Intenta enfocar la primera fila de la tabla para que PAGE_DOWN tenga efecto."""
    filas = driver.find_elements(By.XPATH, "//tr[starts-with(@id,'grid#')]")
    if not filas:
        filas = driver.find_elements(By.XPATH, "//tr[contains(@id,'grid')]")
    if filas:
        try:
            filas[0].click()
            return True
        except:
            pass
    return False

# ==============================
# SELECCIÓN DE HITOS (devuelve set de hitos seleccionados)
# ==============================
def seleccionar_hitos(driver, lista_hitos):
    wait = WebDriverWait(driver, FAST_WAIT)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.ID,"application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ))
    print("✔ Dentro de WebGUI (búsqueda directa + fallback PageDown)")
    time.sleep(0.8)

    acciones = ActionChains(driver)
    seleccionados = set()

    # Enfocar la tabla una vez (por si necesitamos PageDown)
    _focus_tabla(driver)

    for hito in [str(h).strip() for h in lista_hitos]:
        print(f"\n🔍 Hito a seleccionar: {hito}")

        celda = _buscar_celda_hito(driver, hito)
        if not celda:
            # Fallback: escaneo por PAGE DOWN con la tabla enfocada
            for _ in range(MAX_PAGEDOWN_PASOS):
                acciones.send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(SLEEP_PAGEDOWN)
                celda = _buscar_celda_hito(driver, hito)
                if celda:
                    break

        if not celda:
            print(f"❌ No encontrado (selección): {hito}")
            continue

        try:
            fila = celda.find_element(By.XPATH, "./ancestor::tr[1]")
            chk = fila.find_element(By.XPATH, ".//span[contains(@id,'#cb')]")
            driver.execute_script("arguments[0].click();", chk)
            seleccionados.add(hito)
            print(f"✔ Seleccionado: {hito}")
        except Exception as e:
            print(f"❌ No se pudo marcar checkbox del hito {hito}: {e}")

    driver.switch_to.default_content()
    print("✔ Selección completada")
    return seleccionados

# ==============================
# FRD (devuelve set de hitos con FRD marcado)
# ==============================
def marcar_fecha_real_dia(driver, lista_hitos):
    wait = WebDriverWait(driver, FAST_WAIT)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.ID,"application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ))
    print("✔ FRD (búsqueda directa + fallback PageDown)")
    time.sleep(0.8)

    acciones = ActionChains(driver)
    marcados = set()

    _focus_tabla(driver)

    for hito in [str(h).strip() for h in lista_hitos]:
        print(f"\n🔍 FRD para: {hito}")

        celda = _buscar_celda_hito(driver, hito)
        if not celda:
            for _ in range(MAX_PAGEDOWN_PASOS):
                acciones.send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(SLEEP_PAGEDOWN)
                celda = _buscar_celda_hito(driver, hito)
                if celda:
                    break

        if not celda:
            print(f"❌ No encontrado (FRD): {hito}")
            continue

        try:
            fila = celda.find_element(By.XPATH, "./ancestor::tr[1]")
            # FRD: suele ser un #cb distinto al de selección inicial (otra columna de checks)
            # Si sólo hay un #cb por fila, SAP marcará el que esté disponible
            cb = fila.find_element(By.XPATH, ".//span[contains(@id,'#cb')]")
            driver.execute_script("arguments[0].click();", cb)
            marcados.add(hito)
            print(f"✔ FRD marcado: {hito}")
        except Exception as e:
            print(f"❌ No se pudo marcar FRD de {hito}: {e}")

    driver.switch_to.default_content()
    print("✔ FRD finalizado")
    return marcados

# ==============================
# MODIFICACIÓN HITOS
# ==============================
def pulsar_modificacion_hitos(driver):
    driver.switch_to.default_content()
    wait_no_busy(driver)
    print("Pulsando Modificación Hitos…")
    for xp in [
        "//span[contains(text(),'Modificación Hitos')]/ancestor::div",
        "//div[contains(@id,'btn[25]')]"
    ]:
        try:
            btn = WebDriverWait(driver,3).until(EC.element_to_be_clickable((By.XPATH,xp)))
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(1)
            wait_no_busy(driver)
            print("✔ Modificación Hitos abierta")
            return
        except:
            pass
    ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.F1).key_up(Keys.CONTROL).perform()
    print("✔ Modificación por Ctrl+F1")
    time.sleep(1)
    wait_no_busy(driver)

# ==============================
# GRABAR
# ==============================
def pulsar_grabar(driver):
    driver.switch_to.default_content()
    wait_no_busy(driver)
    print("Pulsando GRABAR…")
    for xp in [
        "//span[contains(text(),'Grabar')]/ancestor::div",
        "//div[contains(@id,'btn[11]')]"
    ]:
        try:
            btn = WebDriverWait(driver,3).until(
                EC.element_to_be_clickable((By.XPATH,xp))
            )
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(2)
            wait_no_busy(driver)
            print("✔ Grabado OK")
            return True
        except:
            pass
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
        time.sleep(2)
        wait_no_busy(driver)
        print("✔ Grabado por Ctrl+S")
        return True
    except:
        print("❌ No se pudo pulsar Grabar")
        return False

# ==============================
# EXCEL
# ==============================
def inicializar_excel_resultado(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws.append(["Proyecto", "Hito", "Estado"])
    wb.save(path)

def escribir_resultado(path, proyecto, hito, estado):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    ws.append([proyecto, hito, estado])
    wb.save(path)

# ==============================
# CARGA EXCEL
# ==============================
def cargar_excel():
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = df.columns.str.lower().str.replace(" ", "").str.replace(".", "")
    colp = next(c for c in df.columns if "proyecto" in c or "pep" in c)
    colh = next(c for c in df.columns if "hito" in c)
    df["proyecto"] = df[colp].astype(str).str.strip()
    df["codigo_hito"] = df[colh].astype(str).str.replace(".0","",regex=False).str.strip()
    return df[["proyecto","codigo_hito"]]

# ==============================
# MAIN
# ==============================
def main():
    user, pwd = ensure_env()
    df = cargar_excel()
    inicializar_excel_resultado(RESULTADO_PATH)

    for proyecto, grupo in df.groupby("proyecto"):

        print("\n====================================")
        print("Procesando proyecto:", proyecto)
        print("====================================")

        # Estados por hito (default NOK hasta demostrar lo contrario)
        hitos = [str(h).strip() for h in grupo["codigo_hito"].tolist()]
        estado_por_hito = {h: "NOK" for h in hitos}

        for intento in range(1, MAX_REINTENTOS + 1):
            driver = None
            try:
                driver = iniciar_driver()
                login(driver, user, pwd)
                ejecutar_proyecto(driver, proyecto)

                # 1) Seleccionar hitos → devuelve cuáles se pudieron marcar en la selección
                seleccionados = seleccionar_hitos(driver, hitos)

                # 2) Modificación Hitos
                pulsar_modificacion_hitos(driver)

                # 3) Marcar FRD → devuelve cuáles se pudieron marcar FRD
                frd_marcados = marcar_fecha_real_dia(driver, hitos)

                # 4) Grabar
                grabado_ok = pulsar_grabar(driver)

                # 5) Consolidar estado por hito:
                #    OK sólo si (seleccionado y FRD marcado y grabado_ok)
                for h in hitos:
                    if (h in seleccionados) and (h in frd_marcados) and grabado_ok:
                        estado_por_hito[h] = "OK"
                    else:
                        estado_por_hito[h] = "NOK"

                print(f"✔ Proyecto {proyecto} completado en intento {intento}")
                break  # salimos del bucle de reintentos

            except Exception as e:
                print(f"❌ Intento {intento}/{MAX_REINTENTOS} falló → {e}")
                # Si falla el intento, dejamos los hitos en el estado que tengan (por ahora NOK)
            finally:
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
                time.sleep(RETRASO_ENTRE_REINTENTOS)

        # 6) Escribir resultado por hito
        for h in hitos:
            escribir_resultado(RESULTADO_PATH, proyecto, h, estado_por_hito[h])

    print("\n✔ PROCESO COMPLETO ✔")

if __name__ == "__main__":
    main()
