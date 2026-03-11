###########################################################################################
# sap_facturar_hitos_new.py — FINAL (WebGUI + selección por Nº de hito + botón robusto)
###########################################################################################

import os
import re
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

FAST_WAIT = 15
SLEEP_SHORT = 0.35
ZOOM_LEVELS = [0.9, 0.8, 0.7]   # NO usar scroll; sólo zoom


# ==========================================
# UTILIDADES
# ==========================================
def ensure_env():
    load_dotenv()
    u = os.getenv("FM21_USER2")
    p = os.getenv("FM21_PASS2")
    if not u or not p:
        raise Exception("❌ Faltan credenciales FM21_USER2/FM21_PASS2 en .env")
    print(f"🔐 Usuario cargado: {u} (len={len(p)})")
    return u, p


def safe_type(driver, el, txt):
    """Escritura compatible con entornos SAP (fallback por JS)."""
    try:
        el.clear()
        el.send_keys(txt)
    except:
        driver.execute_script("arguments[0].value='';", el)
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
        """, el, txt)


def ui5_firepress(driver, cid: str) -> bool:
    """Invoca firePress() en un control UI5 por su control-id base."""
    script = """
    try{
        var c = sap.ui.getCore().byId(arguments[0]);
        if(c && c.firePress){ c.firePress(); return true; }
        return false;
    }catch(e){ return false; }
    """
    try:
        return bool(driver.execute_script(script, cid))
    except:
        return False


def derive_id(dom_id: str) -> str:
    """M0:50::btn[8]-cnt → M0:50::btn[8]."""
    if not dom_id:
        return ""
    for suf in ("-cnt", "-bdi", "-inner", "-img", "-icon"):
        if dom_id.endswith(suf):
            return dom_id[:-len(suf)]
    return dom_id


def wait_no_busy(driver):
    """Espera a que desaparezca BusyIndicator/overlays."""
    try:
        WebDriverWait(driver, FAST_WAIT).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".sapUiBlockLayer, .sapUiLocalBusyIndicator")
            )
        )
    except:
        pass


def set_zoom(driver, factor: float):
    """Reduce zoom de página (intenta CDP; fallback a CSS zoom)."""
    try:
        w = driver.execute_script("return window.innerWidth;")
        h = driver.execute_script("return window.innerHeight;")
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {"width": int(w), "height": int(h), "deviceScaleFactor": 1,
             "mobile": False, "scale": factor}
        )
        print(f"🔎 Zoom CDP aplicado: {factor}")
        return
    except:
        pass

    try:
        driver.execute_cdp_cmd(
            "Emulation.setPageScaleFactor",
            {"pageScaleFactor": factor}
        )
        print(f"🔎 Zoom PageScaleFactor aplicado: {factor}")
        return
    except:
        pass

    try:
        driver.execute_script(f"document.body.style.zoom='{int(factor*100)}%';")
        print(f"🔎 Zoom CSS aplicado: {int(factor*100)}%")
    except:
        print("⚠ No se pudo aplicar zoom.")


def dump_iframe_html(driver, path="iframe_dump.html"):
    """Guarda el HTML que ve Selenium dentro del iframe."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"🧪 Dump HTML guardado en {path}")
    except:
        print("⚠ No se pudo guardar dump del iframe.")


# ==========================================
# DRIVER
# ==========================================
def iniciar_driver():
    opt = Options()
    opt.add_argument("--start-maximized")
    return webdriver.Chrome(service=Service(CHROME_DRIVER_PATH), options=opt)


# ==========================================
# LOGIN
# ==========================================
def login(driver, user, pwd):
    URL = (
        "https://fm21global.tg.telefonica/fiori"
        "?sap-client=550&sap-language=ES"
        "#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize"
    )
    driver.get(URL)
    time.sleep(1.2)

    driver.find_element(By.CSS_SELECTOR, "input[placeholder='Usuario']").send_keys(user)
    driver.find_element(By.CSS_SELECTOR, "input[placeholder='Clave de acceso']").send_keys(pwd)
    driver.find_element(By.XPATH, "//button[contains(text(),'Acceder')]").click()
    time.sleep(1.2)

    print("✔ Login OK")


# ==========================================
# PÁGINA 2 — Proyecto + Ejecutar (RESTORED 100%)
# ==========================================
def ejecutar_proyecto(driver, proyecto):
    wait = WebDriverWait(driver, FAST_WAIT)

    # 1) Entrar al iframe del buscador (campo + ejecutar)
    wait.until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, "//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]")
        )
    )

    # 2) Campo proyecto
    campo = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[@title='Definición del proyecto']")
        )
    )
    safe_type(driver, campo, proyecto)
    time.sleep(SLEEP_SHORT)

    # 3) Cerrar sugerencias
    try:
        item = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//ul[(contains(@id,'suggest') or @role='listbox')]/li[1]")
            )
        )
        item.click()
    except:
        try:
            row = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "(//table[contains(@id,'__table') or contains(@class,'sapMListTbl')]/tbody/tr)[1]")
                )
            )
            row.click()
        except:
            try:
                campo.send_keys(Keys.ENTER); time.sleep(SLEEP_SHORT)
                campo.send_keys(Keys.TAB);   time.sleep(SLEEP_SHORT)
            except:
                try:
                    campo.send_keys(Keys.ESCAPE)
                except:
                    driver.execute_script("document.body.click();")

    # 4) Pulsar EJECUTAR (texto + lista de IDs original)
    print("🔎 Buscando EJECUTAR…")
    try:
        node = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::*[self::button][1]")
            )
        )
        dom_id = node.get_attribute("id") or ""
        base = derive_id(dom_id)

        if base and ui5_firepress(driver, base):
            print("✔ Ejecutar con firePress")
        else:
            driver.execute_script("arguments[0].click();", node)
            print(f"✔ Ejecutar JS → {dom_id}")

    except:
        print("⚠ Fallback EJECUTAR con IDs dinámicos…")

        candidates = [
            "//*[@id='M0:50::btn[8]-cnt']",
            "//*[starts-with(@id,'M0:50::btn') and (contains(@id,'-cnt') or contains(@id,'-bdi') or contains(@id,'-inner'))]",
            "//*[@id='M0:50::btn[8]']",
            "//*[@id='M0:50::btn[0]']",
        ]
        clicked = False

        for xp in candidates:
            try:
                node = driver.find_element(By.XPATH, xp)
                dom_id = node.get_attribute("id") or ""
                base = derive_id(dom_id)

                if base and ui5_firepress(driver, base):
                    print(f"✔ Ejecutar con firePress → {base}")
                else:
                    driver.execute_script("arguments[0].click();", node)
                    print(f"✔ Ejecutar JS → {dom_id}")

                clicked = True
                break
            except:
                pass

        if not clicked:
            raise Exception("❌ No pude pulsar EJECUTAR (ni por texto ni por ID)")

    driver.switch_to.default_content()
    wait_no_busy(driver)
    print("✔ Proyecto ejecutado, pasamos a la tabla.")


# ======================================================================================
# PÁGINA 3 — SELECCIONAR HITOS (SAP GUI for HTML / ITS WebGUI)
# ======================================================================================
def seleccionar_hitos(driver, lista_hitos):
    """
    Estrategia WebGUI (no SAPUI5):
      - Detecta índice de columna de “Nº de hito” en el header (grid#...#0,<col>#cp*).
      - Busca cada valor en data (grid#...#1,<col>#if == texto del hito).
      - Sube a <tr> y marca el checkbox grid#...#1,1#cb por JS.
      - Sin scroll; reintenta con zoom 90/80/70.
      - Si falla: dump + screenshots.
    """
    wait = WebDriverWait(driver, FAST_WAIT)

    # Entrar al iframe de la tabla (pantalla 3)
    wait.until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
        )
    )
    print("✔ Dentro del iframe de la tabla (WebGUI).")

    # Espera a que el grid esté pintado
    time.sleep(0.8)

    # --- 1) Detectar índice de columna del "Nº de hito" dinámicamente ---
    # En tu dump, el header tiene spans tipo: grid#C102#0,4#cp2 => texto "Nº de hito"
    # Tomamos varias variantes de texto de cabecera para robustez.
    header_text_variants = [
        "Nº de hito",
        "Número de hito",
        "Nº hito",
    ]

    def _detectar_indice_columna_hito() -> str:
        # Busca spans header que contengan el texto de la cabecera.
        for t in header_text_variants:
            # El header puede estar en #cp1, #cp2, ... y siempre con #0,<COL> en el id.
            xp = (
                "//span[starts-with(@id,'grid#') and contains(@id,'#0,') "
                f"and normalize-space()='{t}']"
            )
            headers = driver.find_elements(By.XPATH, xp)
            if headers:
                hid = headers[0].get_attribute("id") or ""
                # Ej: grid#C102#0,4#cp2  -> extraer ",4#"
                m = re.search(r"#0,(\d+)#cp", hid)
                if m:
                    return m.group(1)
        # Fallback seguro: en tu dump es 4
        return "4"

    col_hito = _detectar_indice_columna_hito()
    print(f"ℹ Índice de columna 'Nº de hito' detectado: {col_hito}")

    pendientes = set(str(h).strip() for h in lista_hitos if str(h).strip())

    def intentar(h: str) -> bool:
        # 2) Buscar el span de datos EXACTO con ese hito en la columna detectada
        xp_valor = (
            f"//span[starts-with(@id,'grid#') and contains(@id,',{col_hito}#if') "
            f"and normalize-space()='{h}']"
        )
        celdas = driver.find_elements(By.XPATH, xp_valor)
        if not celdas:
            return False

        celda = celdas[0]

        # 3) Subir a la fila real
        try:
            fila = celda.find_element(By.XPATH, "./ancestor::tr[1]")
        except:
            return False

        # 4) Checkbox en col 1 => id termina en "#1,1#cb"
        try:
            chk = fila.find_element(By.XPATH, ".//span[contains(@id,'#1,1#cb')]")
        except:
            return False

        # 5) Click JS
        try:
            driver.execute_script("arguments[0].click();", chk)
            print(f"✔ Hito seleccionado: {h}")
            return True
        except:
            # Fallback por tecla espacio, por si acaso
            try:
                chk.send_keys(Keys.SPACE)
                print(f"✔ Hito seleccionado (SPACE): {h}")
                return True
            except:
                return False

    # --- 1er intento ---
    for h in list(pendientes):
        if intentar(h):
            pendientes.remove(h)

    # --- Reintentos con ZOOM (sin scroll) ---
    if pendientes:
        driver.switch_to.default_content()

        for z in ZOOM_LEVELS:
            set_zoom(driver, z)
            time.sleep(0.6)

            wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
                )
            )
            time.sleep(0.6)

            for h in list(pendientes):
                if intentar(h):
                    pendientes.remove(h)

            driver.switch_to.default_content()
            if not pendientes:
                break

    # --- Diagnóstico si quedan pendientes ---
    if pendientes:
        try:
            wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
                )
            )
            dump_iframe_html(driver, "iframe_dump.html")
            for h in pendientes:
                driver.save_screenshot(f"ERROR_hito_{h}.png")
                print(f"🧪 Captura guardada: ERROR_hito_{h}.png")
        except:
            pass

        driver.switch_to.default_content()
        print(f"⚠ No encontrados: {sorted(pendientes)}")
    else:
        print("✔ Todos los hitos seleccionados (WebGUI).")


# ======================================================================================
# BOTÓN — “Modificación Hitos” (primero fuera; si no, dentro de WebGUI)
# ======================================================================================
def pulsar_modificacion_hitos(driver):
    """
    Intenta pulsar el botón "Modificación Hitos".
    1) Primero FUERA del iframe (tu caso inicial con <bdi>).
    2) Si no lo encuentra, lo intenta DENTRO del iframe WebGUI (estructura real).
    """
    wait = WebDriverWait(driver, FAST_WAIT)

    # 1) Intento fuera de iframe (como pediste inicialmente)
    driver.switch_to.default_content()
    wait_no_busy(driver)

    try:
        btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//bdi[contains(normalize-space(),'Modificación Hitos')]/ancestor::button")
            )
        )
        driver.execute_script("arguments[0].click();", btn)
        print("🖱 Pulsado → Modificación Hitos (fuera iframe)")
        wait_no_busy(driver)
        return
    except:
        print("ℹ Botón no visible fuera del iframe; probamos dentro (WebGUI).")

    # 2) Intento dentro del iframe WebGUI (según DOM real)
    # En el dump: existe un botón con id "M0:48::btn[25]" y caption "Modificación Hitos".
    # Probaremos varios selectores compatibles con WebGUI.  [1](https://telefonicacorp-my.sharepoint.com/personal/santiago_perezalbarran_telefonica_com/_layouts/15/Doc.aspx?sourcedoc=%7BE652A0DB-077E-4EE0-BBF4-C4AC775DCD0B%7D&file=cpipasar.docx&action=default&mobileredirect=true)
    try:
        wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.ID, "application-ZOBJ_Z_GESTION_HITOS_0001-display-iframe")
            )
        )

        # a) Por id directo (si está)
        try:
            node = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "M0:48::btn[25]"))
            )
            driver.execute_script("arguments[0].click();", node)
            print("🖱 Pulsado → Modificación Hitos (WebGUI por id)")
            driver.switch_to.default_content()
            wait_no_busy(driver)
            return
        except:
            pass

        # b) Por caption dentro de un botón/div WebGUI
        candidates = [
            # estructura típica: div[@ct='B'] con span -caption que contiene texto
            "//div[@ct='B'][.//span[contains(@id,'-caption') and contains(normalize-space(),'Modificación Hitos')]]",
            # cualquier nodo con caption que contenga el texto
            "//*[contains(@id,'-caption') and contains(normalize-space(),'Modificación Hitos')]/ancestor::*[@ct='B' or @role='button' or self::div][1]",
        ]
        clicked = False
        for xp in candidates:
            try:
                node = driver.find_element(By.XPATH, xp)
                driver.execute_script("arguments[0].click();", node)
                print("🖱 Pulsado → Modificación Hitos (WebGUI por caption)")
                clicked = True
                break
            except:
                continue

        driver.switch_to.default_content()
        wait_no_busy(driver)

        if not clicked:
            raise Exception("No se localizó el botón en WebGUI.")

    except Exception as e:
        driver.switch_to.default_content()
        raise Exception(f"❌ No pude pulsar 'Modificación Hitos': {e}")


# ==========================================
# CARGA EXCEL
# ==========================================
def cargar_excel():
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = df.columns.str.lower().str.replace(" ", "").str.replace(".", "")
    col_proj = next(c for c in df.columns if "pep" in c or "proyecto" in c)
    col_hito = next(c for c in df.columns if "hito" in c)

    df["proyecto"] = df[col_proj].astype(str).str.strip()
    df["codigo_hito"] = df[col_hito].astype(str).str.replace(".0", "").str.strip()

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

            # PANTALLA 2
            ejecutar_proyecto(driver, proyecto)

            # PANTALLA 3 (selección WebGUI)
            seleccionar_hitos(driver, grupo["codigo_hito"].tolist())

            # Botón superior
            pulsar_modificacion_hitos(driver)

            print(f"➡ Siguiente pantalla tras 'Modificación Hitos' para proyecto {proyecto}")

    except Exception as e:
        print("❌ ERROR:", e)

    finally:
        driver.quit()


if __name__ == "__main__":
    main()