# -*- coding: utf-8 -*-
# ==========================================================
# SAP WebGUI — BLOQUE 1/10
# Imports, CONFIG, utils base
# ==========================================================

import os
import re
import sys
import time
import json
import traceback
import pandas as pd
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# CONFIG
EXCEL_PATH = r"C:\Users\bt00092\Downloads\tabla_facturar.xlsx"
RESULTADO_PATH = r"C:\Users\bt00092\Downloads\resultado_hitos.xlsx"
CHROME_DRIVER_PATH = r"C:\Python Project\drivers\chromedriver.exe"

FAST_WAIT = 15
LONG_WAIT = 45
SLEEP_SHORT = 0.35
SLEEP_MED = 0.75

# Credenciales
def ensure_env():
    load_dotenv()
    u = os.getenv("FM21_USER2")
    p = os.getenv("FM21_PASS2")
    if not u or not p:
        raise RuntimeError("Faltan credenciales (.env FM21_USER2 / FM21_PASS2)")
    print("✔ Credenciales cargadas desde .env")
    return u, p

# Esperas globales SAP busy
def wait_no_busy(driver, timeout=FAST_WAIT):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR,
                 ".sapUiBlockLayer, .sapUiLocalBusyIndicator, div[id*='busyIndicator']")
            )
        )
    except:
        print("WARN: busy timeout")

# Escritura segura en inputs
def safe_type(driver, el, txt):
    try:
        el.clear()
        el.send_keys(txt)
        return
    except:
        pass
    driver.execute_script("arguments[0].value='';", el)
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input',{bubbles:true}));
        arguments[0].dispatchEvent(new Event('change',{bubbles:true}));
    """, el, txt)

# Dump HTML
def dump_html(driver, path="dump.html"):
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print("HTML volcado en:", path)

# Driver
def iniciar_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--lang=es-ES")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(CHROME_DRIVER_PATH), options=opts)
    driver.set_page_load_timeout(120)
    print("✔ Chrome iniciado")
    return driver

# Click robusto
def click_robusto(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.1)
    except:
        pass
    try:
        driver.execute_script("arguments[0].click();", el)
        return True
    except:
        pass
    try:
        ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
        return True
    except:
        pass
    try:
        driver.execute_script("""
            const e=arguments[0];
            function fire(t){e.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true}))}
            fire('mouseover');fire('mousedown');fire('mouseup');fire('click');
        """, el)
        return True
    except:
        print("ERROR click_robusto")
        return False
    # ==========================================================
# BLOQUE 2/10 — Zoom + Scroll helpers (SAP WebGUI)
# ==========================================================

def set_zoom_global(driver, percent=10):
    """
    Fuerza zoom global en la página SAP.
    Usa CDP y como fallback CSS zoom.
    """
    ok = False
    try:
        driver.execute_cdp_cmd(
            "Emulation.setPageScaleFactor",
            {"pageScaleFactor": percent / 100}
        )
        ok = True
    except:
        pass

    try:
        driver.execute_script(
            "document.body.style.zoom = arguments[0];",
            f"{percent}%"
        )
        ok = True
    except:
        pass

    if ok:
        print(f"✔ Zoom global aplicado: {percent}%")
    else:
        print("WARN: No se pudo aplicar zoom global")
    return ok


def restore_zoom(driver):
    """
    Restaura zoom global a 100%.
    """
    ok = False
    try:
        driver.execute_cdp_cmd(
            "Emulation.setPageScaleFactor",
            {"pageScaleFactor": 1}
        )
        ok = True
    except:
        pass

    try:
        driver.execute_script("document.body.style.zoom='100%';")
        ok = True
    except:
        pass

    if ok:
        print("✔ Zoom restaurado a 100%")
    else:
        print("WARN: no fue posible restaurar zoom")


def _find_scrollable_ancestor(driver, el):
    """
    Devuelve el primer ancestro que tenga scroll real.
    """
    return driver.execute_script("""
        function isScrollable(n){
            if(!n) return false;
            const oy = getComputedStyle(n).overflowY;
            return (oy === 'auto' || oy === 'scroll') && n.scrollHeight > n.clientHeight;
        }
        let c = arguments[0];
        while (c) {
            if (isScrollable(c)) return c;
            c = c.parentElement;
        }
        return null;
    """, el)


def _scroll_until_condition(driver, container, xp, step=600, max_steps=200):
    """
    Desplaza el contenedor hasta que aparece un elemento que cumpla el xpath.
    Devuelve esa celda o None.
    """
    try:
        els = driver.find_elements(By.XPATH, xp)
        if els:
            return els[0]
    except:
        pass

    if container is None:
        try:
            container = driver.execute_script("return document.scrollingElement")
        except:
            container = None

    try:
        driver.execute_script("arguments[0].scrollTop = 0;", container)
    except:
        pass

    last = -1

    for i in range(max_steps):
        els = driver.find_elements(By.XPATH, xp)
        if els:
            return els[0]

        try:
            top = driver.execute_script("return arguments[0].scrollTop;", container)
            max_top = driver.execute_script(
                "return arguments[0].scrollHeight - arguments[0].clientHeight;",
                container,
            )
        except:
            break

        if top == last or top >= max_top:
            break

        last = top

        try:
            driver.execute_script(
                "arguments[0].scrollTop += arguments[1];",
                container,
                step,
            )
        except:
            try:
                driver.execute_script("window.scrollBy(0, arguments[0]);", step)
            except:
                pass

        time.sleep(0.05)

    return None


def _buscar_celda_hito_con_scroll(driver, grid_id, col_hito, valor_hito):
    """
    Busca una celda de hito por scroll incremental.
    Necesita el grid_id y el número de columna detectado.
    """
    xp_val = (
        f"//span[contains(@id,'grid#{grid_id}') and "
        f"contains(@id,',{col_hito}#if') and normalize-space()='{valor_hito}']"
    )

    # Identificar header para localizar contenedor scrollable
    header_xp = f"//span[starts-with(@id,'grid#{grid_id}#0,')]"
    headers = driver.find_elements(By.XPATH, header_xp)
    container = None

    if headers:
        try:
            container = _find_scrollable_ancestor(driver, headers[0])
        except:
            container = None

    return _scroll_until_condition(driver, container, xp_val, step=700, max_steps=250)
# ==========================================================
# BLOQUE 3/10 — Login + ejecutar_proyecto + iframe helper
# ==========================================================

APP_URL = (
    "https://fm21global.tg.telefonica/fiori"
    "?sap-client=550&sap-language=ES"
    "#ZOBJ_Z_GESTION_HITOS_0001-display?sap-ie=edge&sap-theme=sap_belize"
)

def login(driver, user, pwd):
    """
    Login inicial a SAP (una sola vez por driver).
    """
    print("Navegando a login…")

    driver.switch_to.default_content()
    driver.get(APP_URL)
    time.sleep(1)

    usr = WebDriverWait(driver, LONG_WAIT).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[placeholder='Usuario'],input[name='j_username']")
        )
    )
    pwd_in = WebDriverWait(driver, FAST_WAIT).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[placeholder='Clave de acceso'],input[type='password']")
        )
    )

    safe_type(driver, usr, user)
    safe_type(driver, pwd_in, pwd)

    btn = WebDriverWait(driver, FAST_WAIT).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Acceder') or contains(., 'Iniciar')]")
        )
    )
    click_robusto(driver, btn)

    wait_no_busy(driver, LONG_WAIT)
    time.sleep(1)

    print("✔ Login OK")


def esperar_iframe_app(driver, retries=3):
    """
    Localiza el iframe principal de SAP WebGUI.
    """
    xp = "//iframe[contains(@id,'application-ZOBJ_Z_GESTION_HITOS')]"

    for i in range(retries):
        try:
            driver.switch_to.default_content()
            wait_no_busy(driver)

            WebDriverWait(driver, FAST_WAIT).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, xp))
            )
            print("✔ Iframe app OK")
            return

        except Exception:
            print(f"⚠ No se detecta iframe (intento {i+1}/{retries})…")
            time.sleep(1)

    raise RuntimeError("No se encontró el iframe de la app.")


def ejecutar_proyecto(driver, proyecto, user, pwd):
    """
    Ejecuta el proyecto:
      - Asume que ya se hizo login en el driver actual
      - Entra al iframe
      - Escribe el proyecto
      - Pulsa 'Ejecutar' o usa F8 como fallback
    """

    print(f"→ Ejecutando proyecto {proyecto}…")

    # Entrar al iframe de la app
    esperar_iframe_app(driver)

    wait = WebDriverWait(driver, LONG_WAIT)

    # Campo de proyecto
    campo = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[@title='Definición del proyecto']")
        )
    )
    safe_type(driver, campo, proyecto)
    time.sleep(SLEEP_SHORT)

    # Sugerencia
    try:
        sug = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//ul[(contains(@id,'suggest'))]/li[1]")
            )
        )
        click_robusto(driver, sug)
    except:
        try:
            campo.send_keys(Keys.ENTER)
        except:
            pass

    # Botón Ejecutar
    xps = [
        "//*[self::bdi or self::span][normalize-space()='Ejecutar']/ancestor::button",
        "//button[.//bdi[text()='Ejecutar']]"
    ]

    clicked = False

    for xp in xps:
        try:
            b = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            if click_robusto(driver, b):
                clicked = True
                print("✔ Ejecutar pulsado")
                break
        except:
            pass

    # Fallback: F8
    if not clicked:
        try:
            ActionChains(driver).send_keys(Keys.F8).perform()
            clicked = True
            print("✔ Ejecutar por F8 (fallback)")
        except:
            pass

    if not clicked:
        raise RuntimeError("No se pudo pulsar EL BOTÓN EJECUTAR")

    # Salir del iframe y esperar carga
    driver.switch_to.default_content()
    wait_no_busy(driver, LONG_WAIT)
    time.sleep(SLEEP_MED)

    print("✔ Proyecto ejecutado correctamente")
    