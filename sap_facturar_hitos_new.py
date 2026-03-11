###########################################################################################
# sap_facturar_hitos_new.py — FINAL (iframe + popup + ejecutar + zoom fallback)
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

# Si Selenium “no ve” botones o están fuera de vista, habilita zoom fallback:
ZOOM_FALLBACK_ENABLED = True
ZOOM_FACTORS = [0.9, 0.8]  # intentos de zoom (90%, 80%)

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
    """Intenta reducir zoom de la página (CDP o CSS) para que el botón quede visible."""
    try:
        # 1) CDP Page.setDeviceMetricsOverride (emula scaling del viewport)
        width = driver.execute_script("return window.innerWidth;")
        height = driver.execute_script("return window.innerHeight;")
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {"width": int(width), "height": int(height), "deviceScaleFactor": 1, "mobile": False, "scale": factor}
        )
        print(f"🔎 Zoom CDP aplicado (scale={factor}).")
        return
    except Exception:
        pass
    try:
        # 2) CDP Emulation.setPageScaleFactor
        driver.execute_cdp_cmd("Emulation.setPageScaleFactor", {"pageScaleFactor": factor})
        print(f"🔎 Zoom Emulation.setPageScaleFactor aplicado ({factor}).")
        return
    except Exception:
        pass
    try:
        # 3) CSS zoom
        driver.execute_script(f"document.body.style.zoom='{int(factor*100)}%';")
        print(f"🔎 Zoom CSS aplicado ({int(factor*100)}%).")
    except Exception:
        print("⚠ No se pudo aplicar zoom; continuando sin zoom…")

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
# POPUP HELPERS (SAP Fiori)
# ==========================================
def cerrar_popup_sugerencias(driver, campo, wait: WebDriverWait):
    """
    Intenta cerrar/seleccionar el popup de sugerencias:
      1) seleccionar primer elemento de UL/LI (role=listbox / suggestion-list)
      2) seleccionar primera fila de una tabla (fallback)
      3) ENTER + TAB
      4) ESC
      5) click fuera (body)
    Lanza excepción si nada funciona.
    """
    # 1) UL/LI suggestion list (patrón de tus capturas)
    try:
        primer_item = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//ul[(contains(@id,'suggestion') or @role='listbox')]/li[1]")
            )
        )
        primer_item.click()
        print("✔ Popup: seleccionada primera opción de la lista (UL/LI).")
        return
    except Exception:
        pass

    # 2) Tabla de sugerencias
    try:
        primera_fila_tabla = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "(//table[contains(@id,'__table') or contains(@class,'sapMListTbl')]/tbody/tr)[1]")
            )
        )
        primera_fila_tabla.click()
        print("✔ Popup: seleccionada primera fila de TABLA.")
        return
    except Exception:
        pass

    # 3) ENTER + TAB
    try:
        campo.send_keys(Keys.ENTER)
        time.sleep(0.5)
        campo.send_keys(Keys.TAB)
        print("✔ Popup: cerrado con ENTER + TAB.")
        return
    except Exception:
        pass

    # 4) ESC
    try:
        campo.send_keys(Keys.ESCAPE)
        print("✔ Popup: cerrado con ESC.")
        return
    except Exception:
        pass

    # 5) click fuera (body del iframe)
    try:
        driver.execute_script("document.body.click();")
        print("✔ Popup: cerrado clicando fuera (body).")
        return
    except Exception:
        pass

    # Si llegamos aquí, no cerró:
    raise Exception("❌ No pude cerrar/seleccionar el popup de sugerencias.")

# ==========================================
# CLICK EJECUTAR (en el MISMO IFRAME)
# ==========================================
def click_ejecutar(driver, wait: WebDriverWait):
    """
    Clica EJECUTAR dentro del iframe:
      A) localizar por texto (bdi) y click JS
      B) si no, localizar por IDs SAP (M0:50::btn[0]) y click JS
      C) si falla visibilidad, aplicar zoom fallback y reintentar
    """
    # A) Por texto
    ejecutar_btn = None
    try:
        ejecutar_btn = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//bdi[normalize-space()='Ejecutar']/ancestor::button[1]")
            )
        )
        # click por JS (evita scroll y overlays)
        driver.execute_script("arguments[0].click();", ejecutar_btn)
        print("✔ EJECUTAR pulsado (bdi + JS).")
        return
    except Exception:
        pass

    # B) Por IDs SAP dinámicos
    candidatos = [
        "M0:50::btn[0]", "M0:50::btn[0]-r", "M0:50::btn[1]",
        "M0:50:::0:btn[0]"
    ]
    for pid in candidatos:
        try:
            ejecutar_btn = driver.find_element(By.ID, pid)
            driver.execute_script("arguments[0].click();", ejecutar_btn)
            print(f"✔ EJECUTAR pulsado por ID SAP: {pid}")
            return
        except Exception:
            continue

    # C) Zoom fallback y reintento por texto
    if ZOOM_FALLBACK_ENABLED:
        print("🔎 Activando ZOOM fallback para hacer visible el botón EJECUTAR…")
        for factor in ZOOM_FACTORS:
            set_zoom_fallback(driver, factor)
            time.sleep(0.5)
            try:
                ejecutar_btn = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//bdi[normalize-space()='Ejecutar']/ancestor::button[1]")
                    )
                )
                driver.execute_script("arguments[0].click();", ejecutar_btn)
                print(f"✔ EJECUTAR pulsado tras zoom ({int(factor*100)}%).")
                return
            except Exception:
                continue

    # Si nada funcionó:
    driver.save_screenshot("ERROR_ejecutar.png")
    raise Exception("❌ No pude pulsar EJECUTAR (ver ERROR_ejecutar.png).")

# ==========================================
# BUSCAR PROYECTO — flujo completo
# ==========================================
def buscar_proyecto(driver, proyecto):
    wait = WebDriverWait(driver, 30)
    print(f"\n🔎 Insertando PROYECTO: {proyecto}")

    # 1) Entrar al iframe (varios patrones)
    try:
        iframe = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]")
        ))
    except Exception:
        # Alternativo por si cambia el id en otra build
        iframe = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//iframe[contains(@id,'application-') and contains(@id,'display') and contains(@id,'iframe')]")
        ))
    driver.switch_to.frame(iframe)
    time.sleep(1)

    # 2) Input proyecto
    campo = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@title='Definición del proyecto']")
    ))
    safe_type(driver, campo, proyecto)
    time.sleep(0.8)

    # 3) Cerrar/seleccionar popup
    try:
        cerrar_popup_sugerencias(driver, campo, wait)
    except Exception as e:
        driver.save_screenshot("ERROR_popup.png")
        raise

    # 4) Pulsar EJECUTAR (mantenemos contexto dentro del iframe)
    click_ejecutar(driver, wait)

    time.sleep(3)
    print("✔ Proyecto ejecutado correctamente.")

# ==========================================
# EDITAR HITOS
# ==========================================
def abrir_editar(driver):
    wait = WebDriverWait(driver, 20)
    try:
        btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//bdi[normalize-space()='Editar hito']/ancestor::button")
            )
        )
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
            buscar_proyecto(driver, proyecto)   # (dentro iframe + popup + ejecutar)

            abrir_editar(driver)                # (dentro del mismo iframe)

            for _, row in grupo.iterrows():
                seleccionar_hito(driver, row["codigo_hito"])

            guardar(driver)

    except Exception as e:
        print("❌ ERROR:", e)

    finally:
        driver.quit()

if __name__ == "__main__":
    main()