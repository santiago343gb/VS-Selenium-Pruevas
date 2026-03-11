###########################################################################################
# sap_facturar_hitos_new.py — FINAL COMPLETO Y FUNCIONAL
###########################################################################################

import os
import re
import time
import pandas as pd
from dotenv import load_dotenv

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
CHROME_DRIVER_PATH = r"C:\Python Project\drivers\chromedriver.exe"

FAST_WAIT = 15
SLEEP_SHORT = 0.35

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
                (By.CSS_SELECTOR, ".sapUiBlockLayer,.sapUiLocalBusyIndicator")
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

def dump_iframe_html(driver, path="iframe_dump.html"):
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print("Dump iframe:", path)

def iniciar_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
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
# PANTALLA 2 — Proyecto + Ejecutar
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
    try:
        btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH,"//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::button")
        ))
        driver.execute_script("arguments[0].click();", btn)
        print("✔ Ejecutar pulsado")
    except:
        node = driver.find_element(By.XPATH,"//*[@id='M0:50::btn[8]-cnt']")
        driver.execute_script("arguments[0].click();", node)
        print("✔ Ejecutar fallback")

    driver.switch_to.default_content()
    wait_no_busy(driver)
    print("✔ Proyecto ejecutado correctamente")

# ==============================
# PANTALLA 3 — Selección de hitos (WebGUI)
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
    print("Columna Nº hito detectada:", col)

    pendientes = set(str(h).strip() for h in lista_hitos)

    def intentar(h):
        xp = f"//span[starts-with(@id,'grid#') and contains(@id,',{col}#if') and normalize-space()='{h}']"
        celda = driver.find_elements(By.XPATH,xp)
        if not celda:
            return False
        fila = celda[0].find_element(By.XPATH,"./ancestor::tr[1]")
        chk = fila.find_element(By.XPATH,".//span[contains(@id,'#1,1#cb')]")
        driver.execute_script("arguments[0].click();", chk)
        print("✔ Seleccionado:", h)
        return True

    for h in list(pendientes):
        if intentar(h):
            pendientes.remove(h)

    if pendientes:
        dump_iframe_html(driver)
        print("⚠ No encontrados:", pendientes)
    else:
        print("✔ Todos los hitos seleccionados")

    driver.switch_to.default_content()
# ======================================================================================
# UTILIDADES GRID (Cabeceras)
# ======================================================================================
def _detectar_grid_y_columna_por_titulo(driver, variantes_titulo):
    for titulo in variantes_titulo:
        xp = (
            "//span[starts-with(@id,'grid#') and contains(@id,'#0,') "
            f"and normalize-space()='{titulo}']"
        )
        spans = driver.find_elements(By.XPATH, xp)
        if spans:
            hid = spans[0].get_attribute("id") or ""
            m = re.search(r"grid#([^#]+)#0,(\d+)", hid)
            if m:
                return m.group(1), int(m.group(2)), spans[0]

    for titulo in variantes_titulo:
        xp = (
            "//th[starts-with(@id,'grid#') and contains(@id,'#0,') "
            f"and .//*[normalize-space()='{titulo}']]"
        )
        ths = driver.find_elements(By.XPATH, xp)
        if ths:
            hid = ths[0].get_attribute("id") or ""
            m = re.search(r"grid#([^#]+)#0,(\d+)", hid)
            if m:
                return m.group(1), int(m.group(2)), ths[0]

    return None, None, None

# ======================================================================================
# MARCAR FECHA REAL DÍA (con fallback C142 / 7)
# ======================================================================================
def marcar_fecha_real_dia(driver, lista_hitos):
    wait = WebDriverWait(driver, FAST_WAIT)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
    ))
    time.sleep(0.6)

    grid_hito, col_hito, _ = _detectar_grid_y_columna_por_titulo(
        driver, ["Número de hito", "Nº de hito", "Nº hito"]
    )
    grid_fecha, col_fecha, _ = _detectar_grid_y_columna_por_titulo(
        driver, ["Fecha Real Día", "Fecha Real Dia", "Fecha Real DÃ­a"]
    )

    if grid_fecha is None:
        print("ℹ Usando fallback grid_id C142")
        grid_fecha = "C142"
    if col_fecha is None:
        print("ℹ Usando fallback col_fecha 7")
        col_fecha = 7

    if col_hito is None:
        raise Exception("No pude detectar columna del hito.")

    grid_id = grid_fecha

    pendientes = set(str(h).strip() for h in lista_hitos)

    def marcar_en_fila(fila):
        xps = [
            f".//span[contains(@id,'grid#{grid_id}') and contains(@id,',{col_fecha}#cb')]",
            f".//span[contains(@id,',{col_fecha}#cb')]",
        ]
        for xp in xps:
            cands = fila.find_elements(By.XPATH,xp)
            if cands:
                cb = cands[0]
                try:
                    driver.execute_script("arguments[0].click();", cb)
                    return True
                except:
                    cb.send_keys(Keys.SPACE)
                    return True
        return False

    for h in list(pendientes):
        xp = (
            f"//span[contains(@id,'grid#{grid_id}') and contains(@id,',{col_hito}#if') "
            f"and normalize-space()='{h}']"
        )
        celdas = driver.find_elements(By.XPATH,xp)
        if not celdas:
            continue
        fila = celdas[0].find_element(By.XPATH,"./ancestor::tr[1]")

        if marcar_en_fila(fila):
            print(f"✔ Marcado FRD → {h}")
            pendientes.remove(h)

    driver.switch_to.default_content()

    if pendientes:
        print("⚠ No marqué FRD en:", pendientes)
    else:
        print("✔ FRD marcada a todos los hitos")
        # ======================================================================================
# BOTÓN — “Modificación Hitos”
# ======================================================================================
def pulsar_modificacion_hitos(driver):
    driver.switch_to.default_content()
    wait_no_busy(driver)
    time.sleep(0.3)

    print("Pulsando Modificación Hitos…")

    # A) Click por ID dinámico y caption
    xpaths = [
        "//div[starts-with(@id,'M') and contains(@id,'btn[25]') and contains(@class,'lsButton')]",
        "//span[contains(@id,'btn[25]-caption') and contains(normalize-space(),'Modificación Hitos')]/ancestor::div",
    ]
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH,xp))
            )
            driver.execute_script("arguments[0].click();", el)
            print(f"✔ Pulsado → {xp}")
            time.sleep(1); wait_no_busy(driver); return
        except:
            pass

    # B) Hotkey Ctrl + F1
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.F1).key_up(Keys.CONTROL).perform()
        print("✔ Ejecutado Ctrl+F1")
        time.sleep(1)
        wait_no_busy(driver)
        return
    except:
        pass

    # C) Menú → Modificación Hitos
    try:
        btn_menu = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH,"//bdi[normalize-space()='Menú']/ancestor::button"))
        )
        driver.execute_script("arguments[0].click();", btn_menu)
        time.sleep(0.5)

        item = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH,"//*[@id='CtxMnu0']//tr[.//span[contains(text(),'Modificación Hitos')]]"))
        )
        driver.execute_script("arguments[0].click();", item)
        print("✔ Pulsado vía menú")
        time.sleep(1)
        wait_no_busy(driver)
        return
    except:
        pass

    raise Exception("❌ No pude pulsar Modificación Hitos")
# ==============================
# EXCEL
# ==============================
def cargar_excel():
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = df.columns.str.lower().str.replace(" ", "").str.replace(".", "")
    colp = next(c for c in df.columns if "pep" in c or "proyecto" in c)
    colh = next(c for c in df.columns if "hito" in c)
    df["proyecto"] = df[colp].astype(str).str.strip()
    df["codigo_hito"] = df[colh].astype(str).str.replace(".0","").str.strip()
    return df[["proyecto","codigo_hito"]]

# ==============================
# MAIN
# ==============================
def main():
    user, pwd = ensure_env()
    driver = iniciar_driver()

    try:
        login(driver, user, pwd)
        df = cargar_excel()
        print(df)

        for proyecto, grupo in df.groupby("proyecto"):
            ejecutar_proyecto(driver, proyecto)
            seleccionar_hitos(driver, grupo["codigo_hito"].tolist())

            pulsar_modificacion_hitos(driver)
            print("✔ Modificación Hitos abierta")

            marcar_fecha_real_dia(driver, grupo["codigo_hito"].tolist())

    except Exception as e:
        print("ERROR:", e)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
   