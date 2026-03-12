###########################################################################################
# sap_facturar_hitos_new.py — MULTIPROYECTO SEGURO + REINTENTOS + ZOOM
###########################################################################################

import os
import re
import time
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

# Reintentos por proyecto (para neutralizar cierres de Chrome / sesiones inválidas)
MAX_REINTENTOS = 3
RETRASO_ENTRE_REINTENTOS = 1.2  # seg

# --- ZOOM ---
UMBRAL_ZOOM_HITOS = 50       # si hay más de 50 hitos, aplica zoom
FORZAR_ZOOM_SIEMPRE = False  # si True, aplica zoom siempre
ZOOM_PORCENTAJE = 10         # porcentaje de zoom (CSS) cuando aplica
ZOOM_CTRL_MINUS_PULSACIONES = 8  # veces Ctrl+'-' si CSS no aplica

# ==============================
# UTILIDADES
# ==============================
def ensure_env():
    load_dotenv()
    u = os.getenv("FM21_USER2")
    p = os.getenv("FM21_PASS2")
    if not u or not p:
        raise Exception("Faltan credenciales (.env FM21_USER2 / FM21_PASS2)")
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
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
            arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
        """, el, txt)

# ==============================
# ZOOM helpers
# ==============================
def _apply_zoom_css_in_iframe(driver, iframe_locator, percent: int):
    WebDriverWait(driver, FAST_WAIT).until(
        EC.frame_to_be_available_and_switch_to_it(iframe_locator)
    )
    ok = False
    try:
        driver.execute_script("document.body.style.zoom=arguments[0];", f"{percent}%")
        time.sleep(0.2)
        ok = True
        print(f"✔ Zoom CSS {percent}% aplicado en iframe")
    except Exception as e:
        print("⚠ No se pudo aplicar zoom CSS en iframe:", e)
    finally:
        driver.switch_to.default_content()
    return ok

def _apply_zoom_ctrl_minus(driver, veces=8):
    try:
        for _ in range(veces):
            try:
                ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.SUBTRACT).key_up(Keys.CONTROL).perform()
            except:
                ActionChains(driver).key_down(Keys.CONTROL).send_keys('-').key_up(Keys.CONTROL).perform()
            time.sleep(0.08)
        print(f"✔ Zoom navegador reducido con Ctrl+'-' x{veces}")
    except Exception as e:
        print("⚠ No se pudo emular Ctrl+'-':", e)

def aplicar_zoom_tabla_hitos(driver, percent=10, usar_css=True):
    iframe_locator = (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ok = False
    if usar_css:
        ok = _apply_zoom_css_in_iframe(driver, iframe_locator, percent)
    if not ok:
        _apply_zoom_ctrl_minus(driver, veces=ZOOM_CTRL_MINUS_PULSACIONES)
    time.sleep(0.2)
    wait_no_busy(driver)

# ==============================
# DRIVER NUEVO POR PROYECTO
# ==============================
def iniciar_driver():
    opts = Options()
    opts.add_argument("--start-maximized")

    # Estabilidad con Chrome 145+
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-software-rasterizer")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--remote-debugging-pipe")

    # (Opcional) reducir interferencias de Chrome
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
# EJECUTAR PROYECTO (robusto con F8)
# ==============================
def ejecutar_proyecto(driver, proyecto):
    wait = WebDriverWait(driver, FAST_WAIT)

    # 1) Entrar al iframe principal de la app
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.XPATH,"//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]")
    ))

    # 2) Campo de proyecto
    campo = wait.until(EC.presence_of_element_located(
        (By.XPATH,"//input[@title='Definición del proyecto']")
    ))
    safe_type(driver, campo, proyecto)
    time.sleep(SLEEP_SHORT)

    # 3) Sugerencia o Enter
    try:
        sug = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH,"//ul[(contains(@id,'suggest') or @role='listbox')]/li[1]"))
        )
        sug.click()
    except:
        try:
            campo.send_keys(Keys.ENTER)
        except:
            pass

    print("Buscando EJECUTAR…")

    # 4) Click botón Ejecutar (varias estrategias)
    clicked = False
    for xp in [
        "//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::button",
        "//button[.//bdi[normalize-space()='Ejecutar']]",
        "//button[.//*[normalize-space()='Ejecutar']]",
        "//*[@id='M0:50::btn[8]-cnt']/ancestor::button | //*[@id='M0:50::btn[8]-cnt']"
    ]:
        try:
            btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
            driver.execute_script("arguments[0].click();", btn)
            clicked = True
            print("✔ Ejecutar pulsado")
            break
        except:
            pass

    # 5) Fallback universal SAP: F8 (dentro del iframe)
    if not clicked:
        try:
            ActionChains(driver).send_keys(Keys.F8).perform()
            clicked = True
            print("✔ Ejecutar por F8 (iframe)")
        except:
            pass

    # 6) Salir del iframe y espera
    driver.switch_to.default_content()
    wait_no_busy(driver)

    # 7) Si aún no se consiguió, mando F8 también en root (algunas UIs lo escuchan)
    if not clicked:
        try:
            ActionChains(driver).send_keys(Keys.F8).perform()
            print("✔ Ejecutar por F8 (root)")
        except:
            print("❌ No se pudo pulsar Ejecutar (ni con F8)")

    print("✔ Proyecto ejecutado correctamente")

# ==============================
# SELECCIONAR HITOS
# ==============================
def seleccionar_hitos(driver, lista_hitos):
    wait = WebDriverWait(driver, FAST_WAIT)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.ID,"application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ))
    print("✔ Dentro iframe WebGUI")

    time.sleep(1.0)

    header_variants = ["Nº de hito", "Número de hito", "Nº hito"]

    def detectar_col():
        for t in header_variants:
            xp = f"//span[starts-with(@id,'grid#') and contains(@id,'#0,') and normalize-space()='{t}']"
            h = driver.find_elements(By.XPATH,xp)
            if h:
                hid = h[0].get_attribute("id")
                m = re.search(r"#0,(\d+)#cp", hid)
                if m:
                    return m.group(1)
        return "4"

    col = detectar_col()

    pendientes = set(str(h).strip() for h in lista_hitos)

    for h in list(pendientes):
        xp = f"//span[starts-with(@id,'grid#') and contains(@id,',{col}#if') and normalize-space()='{h}']"
        celda = driver.find_elements(By.XPATH,xp)
        if not celda:
            continue

        fila = celda[0].find_element(By.XPATH,"./ancestor::tr[1]")
        chk = fila.find_element(By.XPATH,".//span[contains(@id,'#1,1#cb')]")
        driver.execute_script("arguments[0].click();", chk)
        pendientes.remove(h)
        print("✔ Seleccionado:", h)

    driver.switch_to.default_content()

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
            btn = WebDriverWait(driver,3).until(
                EC.element_to_be_clickable((By.XPATH,xp))
            )
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(1)
            wait_no_busy(driver)
            print("✔ Modificación Hitos abierta")
            return
        except:
            pass

    # Fallback atajo
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.F1).key_up(Keys.CONTROL).perform()
        time.sleep(1)
        wait_no_busy(driver)
        print("✔ Modificación Hitos por Ctrl+F1")
    except:
        print("❌ No se pudo abrir Modificación Hitos")

# ==============================
# MARCAR FRD
# ==============================
def marcar_fecha_real_dia(driver, lista_hitos):
    wait = WebDriverWait(driver, FAST_WAIT)
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.ID,"application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ))
    time.sleep(0.6)

    for h in lista_hitos:
        xp = f"//span[normalize-space()='{h}']/ancestor::tr[1]"
        filas = driver.find_elements(By.XPATH, xp)
        if not filas:
            continue

        fila = filas[0]
        cb = fila.find_element(By.XPATH, ".//span[contains(@id,'#cb')]")
        driver.execute_script("arguments[0].click();", cb)
        print("✔ FRD marcado:", h)

    driver.switch_to.default_content()

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
            return
        except:
            pass

    # Fallback Ctrl+S
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
        time.sleep(2)
        wait_no_busy(driver)
        print("✔ Grabado por Ctrl+S")
    except:
        print("❌ No se pudo pulsar Grabar")

# ==============================
# EXCEL RESULTADO
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
    df["codigo_hito"] = df[colh].astype(str).str.replace(".0","", regex=False).str.strip()

    return df[["proyecto","codigo_hito"]]

# ==============================
# MAIN — CICLO LIMPIO POR PROYECTO + REINTENTOS + ZOOM
# ==============================
def main():
    user, pwd = ensure_env()
    df = cargar_excel()
    inicializar_excel_resultado(RESULTADO_PATH)

    for proyecto, grupo in df.groupby("proyecto"):
        print("\n=====================================")
        print("Procesando proyecto:", proyecto)
        print("=====================================")

        estado_final = "NOK"
        # Reintentos por proyecto
        for intento in range(1, MAX_REINTENTOS + 1):
            driver = None
            try:
                driver = iniciar_driver()
                login(driver, user, pwd)

                ejecutar_proyecto(driver, proyecto)

                # Zoom en selección si procede
                if FORZAR_ZOOM_SIEMPRE or len(grupo) > UMBRAL_ZOOM_HITOS:
                    print(f"Aplicando zoom {ZOOM_PORCENTAJE}% en selección de hitos…")
                    aplicar_zoom_tabla_hitos(driver, percent=ZOOM_PORCENTAJE, usar_css=True)

                seleccionar_hitos(driver, grupo["codigo_hito"].tolist())

                pulsar_modificacion_hitos(driver)

                # Zoom en modificación si procede
                if FORZAR_ZOOM_SIEMPRE or len(grupo) > UMBRAL_ZOOM_HITOS:
                    print(f"Aplicando zoom {ZOOM_PORCENTAJE}% en Modificación Hitos…")
                    aplicar_zoom_tabla_hitos(driver, percent=ZOOM_PORCENTAJE, usar_css=True)

                marcar_fecha_real_dia(driver, grupo["codigo_hito"].tolist())

                pulsar_grabar(driver)

                estado_final = "OK"
                print(f"✔ Proyecto {proyecto} completado en intento {intento}")
                break  # salir del bucle de reintentos

            except Exception as e:
                print(f"❌ Intento {intento}/{MAX_REINTENTOS} para {proyecto} falló → {e}")
                estado_final = "NOK"
                # Cierra y reintenta con una sesión completamente nueva
            finally:
                try:
                    if driver:
                        driver.quit()
                except:
                    pass
                time.sleep(RETRASO_ENTRE_REINTENTOS)

        # Registrar resultados de todos los hitos del proyecto
        for hito in grupo["codigo_hito"].tolist():
            escribir_resultado(RESULTADO_PATH, proyecto, hito, estado_final)

    print("\n✔ PROCESO COMPLETO ✔")

if __name__ == "__main__":
    main()