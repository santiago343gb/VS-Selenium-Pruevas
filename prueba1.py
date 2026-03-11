###########################################################################################
# sap_facturar_hitos_new.py — FINAL (iframe + popup + Ejecutar con UI5 firePress)
###########################################################################################

import os
import time
import pandas as pd
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# CONFIG
# ==========================================
EXCEL_PATH = r"C:\Users\bt00092\Downloads\tabla_facturar.xlsx"
CHROME_DRIVER_PATH = r"C:\Python Project\drivers\chromedriver.exe"

# Zoom fallback si Selenium “no ve” el botón
ZOOM_FALLBACK_ENABLED = True
ZOOM_FACTORS = [0.9, 0.8]

# ==========================================
# UTILIDADES
# ==========================================
def ensure_env():
    load_dotenv()
    u = os.getenv("FM21_USER2")
    p = os.getenv("FM21_PASS2")
    if not u or not p:
        raise Exception("❌ Faltan FM21_USER2 y/o FM21_PASS2 en .env")
    print(f"🔐 Credenciales cargadas → {u}, len_pass={len(p)}")
    return u, p

def safe_type(driver, element, text):
    """Escribe aunque send_keys falle (inyecta por JS si es necesario)."""
    try:
        element.clear()
        element.send_keys(text)
    except:
        driver.execute_script("arguments[0].value='';", element)
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
            arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
        """, element, text)

def set_zoom_fallback(driver, factor: float):
    """Reduce el zoom para forzar visibilidad (varias técnicas)."""
    try:
        w = driver.execute_script("return window.innerWidth;")
        h = driver.execute_script("return window.innerHeight;")
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {"width": int(w), "height": int(h), "deviceScaleFactor": 1, "mobile": False, "scale": factor}
        )
        print(f"🔎 Zoom CDP applied: scale={factor}")
        return
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Emulation.setPageScaleFactor", {"pageScaleFactor": factor})
        print(f"🔎 Zoom Emulation.setPageScaleFactor applied: {factor}")
        return
    except Exception:
        pass
    try:
        driver.execute_script(f"document.body.style.zoom='{int(factor*100)}%';")
        print(f"🔎 Zoom CSS aplicado: {int(factor*100)}%")
    except:
        print("⚠ No se pudo aplicar zoom; continúo sin zoom.")

def ui5_fire_press(driver, control_id: str) -> bool:
    """
    Intenta invocar sap.ui.getCore().byId(control_id).firePress() en el contexto de la página.
    Devuelve True si pudo invocarse, False en caso contrario.
    """
    script = """
        try {
            if (window.sap && sap.ui && sap.ui.getCore) {
                var c = sap.ui.getCore().byId(arguments[0]);
                if (c && c.firePress) { c.firePress(); return true; }
            }
            return false;
        } catch(e) { return false; }
    """
    try:
        return bool(driver.execute_script(script, control_id))
    except Exception:
        return False

def derive_ui5_control_id(dom_id: str) -> str:
    """
    A partir de un id DOM tipo 'M0:50::btn[8]-cnt' o '...-bdi', deriva el control-id base 'M0:50::btn[8]'.
    """
    if not dom_id:
        return ""
    # quita sufijos típicos de subnodos
    for suf in ("-cnt", "-bdi", "-inner", "-img", "-icon"):
        if dom_id.endswith(suf):
            return dom_id[: -len(suf)]
    return dom_id

# ==========================================
# DRIVER
# ==========================================
def iniciar_driver():
    options = Options()
    options.add_argument("--start-maximized")
    service = Service(CHROME_DRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

# ==========================================
# LOGIN SAP
# ==========================================
def login(driver, user, password):
    url = (
        "https://fm21global.tg.telefonica/fiori"
        "?sap-client=550&sap-language=ES"
        "#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize"
    )
    print("🌐 Abriendo SAP…")
    driver.get(url)
    time.sleep(2)

    driver.find_element(By.CSS_SELECTOR, "input[placeholder='Usuario']").send_keys(user)
    driver.find_element(By.CSS_SELECTOR, "input[placeholder='Clave de acceso']").send_keys(password)
    driver.find_element(By.XPATH, "//button[contains(text(),'Acceder')]").click()

    print("✔ Login completado.")
    time.sleep(4)

# ==========================================
# POPUP SUGERENCIAS
# ==========================================
def cerrar_popup_sugerencias(driver, campo, wait: WebDriverWait):
    """
    Cierra o selecciona el popup de sugerencias mediante varias estrategias.
    """
    # 1) UL/LI listbox (lo de tus capturas)
    try:
        item = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//ul[(contains(@id,'suggestion') or @role='listbox')]/li[1]")
        ))
        item.click()
        print("✔ Popup: seleccionada 1ª opción (UL/LI).")
        return
    except Exception:
        pass
    # 2) Tabla
    try:
        row = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "(//table[contains(@id,'__table') or contains(@class,'sapMListTbl')]/tbody/tr)[1]")
        ))
        row.click()
        print("✔ Popup: seleccionada 1ª fila (TABLA).")
        return
    except Exception:
        pass
    # 3) ENTER + TAB
    try:
        campo.send_keys(Keys.ENTER); time.sleep(0.3)
        campo.send_keys(Keys.TAB);   time.sleep(0.3)
        print("✔ Popup: cerrado con ENTER+TAB.")
        return
    except Exception:
        pass
    # 4) ESC
    try:
        campo.send_keys(Keys.ESCAPE); time.sleep(0.3)
        print("✔ Popup: cerrado con ESC.")
        return
    except Exception:
        pass
    # 5) click fuera
    try:
        driver.execute_script("document.body.click();"); time.sleep(0.3)
        print("✔ Popup: cerrado clicando fuera (body).")
        return
    except Exception:
        pass

    raise Exception("❌ No pude cerrar/seleccionar el popup de sugerencias.")

# ==========================================
# CLICK EJECUTAR (dentro del IFRAME) — UI5
# ==========================================
def click_ejecutar(driver, wait: WebDriverWait):
    """
    Pulsa EJECUTAR dentro del iframe empleando:
      1) localizar por texto y derivar control-id base -> UI5 firePress()
      2) localizar por id DOM con sufijo (-cnt, -bdi, etc.) -> derivar control-id -> firePress()
      3) click JS sobre el DOM node
      4) zoom fallback
      5) F8 (algunas apps lo asocian a ejecutar)
    """
    # 1) Por texto 'Ejecutar'
    dom_node = None
    try:
        # bdi o span con 'Ejecutar' y sube al contenedor (button/div) más cercano
        dom_node = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::*[self::button or self::div][1]")
        ))
        dom_id = dom_node.get_attribute("id") or ""
        base_id = derive_ui5_control_id(dom_id)
        print(f"🔎 Ejecutar detectado por texto → dom_id='{dom_id}', base_id='{base_id}'")
        if base_id and ui5_fire_press(driver, base_id):
            print("✔ UI5 firePress() → Ejecutar (por texto).")
            return
        # si firePress no funcionó, intenta click JS sobre el contenedor
        driver.execute_script("arguments[0].click();", dom_node)
        print("✔ Ejecutar pulsado por JS (por texto).")
        return
    except Exception:
        pass

    # 2) Por id DOM de patrón SAP (como 'M0:50::btn[8]-cnt')
    candidates_xpath = [
        # el exacto que acabas de ver en DevTools:
        "//*[@id='M0:50::btn[8]-cnt']",
        # otros posibles ids con -cnt/-bdi:
        "//*[starts-with(@id,'M0:50::btn') and (contains(@id,'-cnt') or contains(@id,'-bdi') or contains(@id,'-inner'))]",
        # el propio button base si existe
        "//*[@id='M0:50::btn[8]']",
        "//*[@id='M0:50::btn[0]']",
    ]
    for xp in candidates_xpath:
        try:
            dom_node = driver.find_element(By.XPATH, xp)
            dom_id = dom_node.get_attribute("id") or ""
            base_id = derive_ui5_control_id(dom_id)
            print(f"🔎 Ejecutar detectado por XPath '{xp}' → dom_id='{dom_id}', base_id='{base_id}'")
            if base_id and ui5_fire_press(driver, base_id):
                print("✔ UI5 firePress() → Ejecutar (por id).")
                return
            driver.execute_script("arguments[0].click();", dom_node)
            print("✔ Ejecutar pulsado por JS (por id).")
            return
        except Exception:
            continue

    # 3) Zoom fallback y reintento por texto
    if ZOOM_FALLBACK_ENABLED:
        print("🔎 Activando ZOOM fallback para hacer visible el botón EJECUTAR…")
        for factor in ZOOM_FACTORS:
            set_zoom_fallback(driver, factor); time.sleep(0.4)
            try:
                dom_node = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::*[self::button or self::div][1]")
                ))
                dom_id = dom_node.get_attribute("id") or ""
                base_id = derive_ui5_control_id(dom_id)
                if base_id and ui5_fire_press(driver, base_id):
                    print(f"✔ UI5 firePress() tras zoom {int(factor*100)}%.")
                    return
                driver.execute_script("arguments[0].click();", dom_node)
                print(f"✔ Ejecutar pulsado por JS tras zoom {int(factor*100)}%.")
                return
            except Exception:
                continue

    # 4) Último recurso: F8 (algunas apps lo usan como ejecutar)
    try:
        print("⌨ Enviando F8 como último recurso…")
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.F8)
        time.sleep(1.5)
        print("✔ F8 enviado.")
        return
    except Exception:
        pass

    # 5) Fallo
    driver.save_screenshot("ERROR_ejecutar.png")
    raise Exception("❌ No pude pulsar EJECUTAR (ver ERROR_ejecutar.png).")

# ==========================================
# BUSCAR PROYECTO
# ==========================================
def buscar_proyecto(driver, proyecto):
    wait = WebDriverWait(driver, 30)
    print(f"\n🔎 Insertando PROYECTO: {proyecto}")

    # 1) Entrar al iframe ZHITOS
    iframe = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]")
    ))
    driver.switch_to.frame(iframe); time.sleep(0.8)

    # 2) Input del proyecto
    campo = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@title='Definición del proyecto']")
    ))
    safe_type(driver, campo, proyecto); time.sleep(0.6)

    # 3) Cerrar/seleccionar popup
    try:
        cerrar_popup_sugerencias(driver, campo, wait)
    except Exception as e:
        driver.save_screenshot("ERROR_popup.png")
        raise

    # 4) Pulsar EJECUTAR (dentro del mismo iframe)
    click_ejecutar(driver, wait)

    time.sleep(2.5)
    print("✔ Proyecto ejecutado.")

# ==========================================
# EDITAR HITOS
# ==========================================
def abrir_editar(driver):
    wait = WebDriverWait(driver, 20)
    try:
        btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//bdi[normalize-space()='Editar hito']/ancestor::button")
        ))
        # Click por JS para evitar overlays
        driver.execute_script("arguments[0].click();", btn)
        print("🖱 'Editar hito' pulsado.")
    except Exception:
        raise Exception("❌ No encontré 'Editar hito'.")
    time.sleep(2)

# ==========================================
# SELECCIONAR HITO
# ==========================================
def seleccionar_hito(driver, hito):
    try:
        fila = driver.find_element(By.XPATH, f"//tr[td[contains(text(),'{hito}')]]")
        chk  = fila.find_element(By.XPATH, ".//input[@type='checkbox']")
        driver.execute_script("arguments[0].click();", chk)
        print(f"✔ Hito seleccionado: {hito}")
    except Exception as e:
        print(f"⚠ Hito no encontrado {hito}: {e}")

# ==========================================
# GUARDAR (OFF)
# ==========================================
def guardar(driver):
    print("🛑 Guardado desactivado por seguridad.")

# ==========================================
# CARGAR EXCEL
# ==========================================
def cargar_excel():
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = (df.columns.str.lower().str.replace(" ", "").str.replace(".", ""))
    col_pep  = [c for c in df.columns if "pep" in c][0]
    col_hito = [c for c in df.columns if "hito" in c][0]

    df["proyecto"] = df[col_pep].astype(str).str.strip()

    def norm(v):
        if pd.isna(v): return ""
        s = str(v).strip()
        if s.endswith(".0"): s = s[:-2]
        return s

    df["codigo_hito"] = df[col_hito].apply(norm)
    df = df[(df["proyecto"]!="") & (df["codigo_hito"]!="")]
    return df[["proyecto", "codigo_hito"]]

# ==========================================
# MAIN
# ==========================================
def main():
    user, pwd = ensure_env()
    driver = iniciar_driver()

    try:
        login(driver, user, pwd)

        df = cargar_excel()
        print("\n📄 Excel cargado:")
        print(df)

        for proyecto, grupo in df.groupby("proyecto"):
            buscar_proyecto(driver, proyecto)   # (iframe + popup + ejecutar con UI5)
            abrir_editar(driver)                # (mismo iframe)
            for _, row in grupo.iterrows():
                seleccionar_hito(driver, row["codigo_hito"])
            guardar(driver)

    except Exception as e:
        print("❌ ERROR:", e)

    finally:
        driver.quit()

if __name__ == "__main__":
    main()