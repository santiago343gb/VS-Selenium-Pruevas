#IMPORTS
'''
=========================================================================================
# sap_facturar_hitos_old.py
# Fecha de creacion: 06/05/2025
# Correo: miguelangel.diazmagister@telefonica.com
# Script para cambiar la fecha real de un hito de un proyecto en SAP
# v23
=========================================================================================
'''
import os,sys,oracledb, re,subprocess,json, codecs, time,shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import ProcessPoolExecutor, as_completed
from contextlib import contextmanager
from datetime import datetime
from dotenv import load_dotenv
path_CURRENT=os.path.dirname(os.path.realpath(__file__))
sys.path.append(path_CURRENT+'./../../')
from utilities.data import paths
from utilities.master import exportDF, totalTime, configureLogson
#=====================================================================
#=====================================================================
#=====================================================================
os.environ['GRPC_VERBOSITY'] = 'ERROR'
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
os.environ['ABSL_LOG_LEVEL'] = '3'
# sys.stderr = open(os.devnull, 'w')
start=datetime.now()
load_dotenv()
username = os.getenv('FM21_USER9')
password = os.getenv('FM21_PASS9')
#=======================================================================================
#path_descargas
path_local_download = 'C:/Users/BT00092/Downloads/'
# path_local_download = 'C:/Users/t152430/Downloads/'
# path_local_download = 'C:/Users/Administrator/Downloads/'

FLAG_PRUEBA = False 

name_excel='facturacion_extraordinaria_sotis_20260225.xlsx'
#path_origen
path = paths['onedrive-database']+'DB_facturacion/Facturacion automatica/' + name_excel
#path_result
path_result = paths['onedrive-database']+'DB_facturacion/Facturacion automatica/Resultado/RESULT - ' + name_excel
#=======================================================================================
def timexHito(start, hitos_len):
    if hitos_len == 0: return 0
    tiempo_total = (datetime.now() - start).total_seconds()
    tiempo_por_hito = tiempo_total / hitos_len
    return tiempo_por_hito
# Cargar Excel y preparar datos
df = pd.read_excel(path)
df['resultado'] = ''
lista_proy = df['PROYECTO'].unique().tolist()
# URL base
url_web = 'https://fm21global.tg.telefonica/fiori#Project-displayDetails?sap-ui-tech-hint=GUI'

def iniciar_driver():
    svc = Service(path_local_download + 'chromedriver-win32/chromedriver.exe')
    chrome_options = webdriver.ChromeOptions()
    if not FLAG_PRUEBA: chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=svc, options=chrome_options)
    return driver

# Función de login, por cada instancia
def login(driver):
    driver.get(url_web)
    wait = WebDriverWait(driver, 30)
    wait.until(EC.presence_of_element_located((By.ID, 'USERNAME_FIELD-inner'))).send_keys(username)
    driver.find_element(By.ID, 'USERNAME_FIELD-inner').send_keys(Keys.TAB)
    wait.until(EC.presence_of_element_located((By.ID, 'PASSWORD_FIELD-inner'))).send_keys(password)
    driver.find_element(By.ID, 'PASSWORD_FIELD-inner').send_keys(Keys.ENTER)
    WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(1))  # Espera a que esté listo

# Facturación
def facturarListaHitos(proy, hitos):
    resultado_bloque = 'NOK'
    try:
        driver = iniciar_driver()
        driver.get(url_web)
        login(driver)
        wait = WebDriverWait(driver, 30)
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.frame('application-Project-displayDetails-iframe')
        wait.until(EC.presence_of_element_located((By.ID, 'M0:46:::4:22'))).send_keys(proy)
        driver.find_element(By.ID, 'M0:46:::11:22').send_keys(hitos[0])
        time.sleep(3)
        driver.find_element(By.ID, 'M0:50::btn[0]').click()
        time.sleep(5)
        driver.find_element(By.ID,'M0:48::btn[13]-r').click()
        time.sleep(3)
        if len(driver.find_elements(By.XPATH,'//*[contains(text(), "No se han bloqueado todos los objeto")]'))>0:
            raise ValueError('NOK-Bloqueo')
        if len(driver.find_elements(By.XPATH,'//*[contains(text(), "No se ha encontrado")]'))>0:
            raise ValueError('NOK-No se ha encontrado el hito')
        else:
            flag_first = True
            for hit in hitos:
                # if hit in[7475,7480, 7508]: raise ValueError('NOK-Prueba')
                time.sleep(2)
                try:
                    if flag_first:
                        driver.find_element(By.ID,'C127-hiddenOpener').click()
                        time.sleep(2)
                        driver.find_element(By.ID,'C127_btn14-BtnMenu').click()
                        flag_first = False
                    time.sleep(2)
                    driver.find_element(By.XPATH,'//*[contains(text(), '+str(hit)+')]').click()
                except Exception as e:
                    # print(e)
                    driver.find_element(By.ID,'C127-hiddenOpener').click()
                    time.sleep(2)
                    driver.find_element(By.ID,'C127_btn11-BtnMenu').click()
                    time.sleep(2)
                    driver.find_element(By.ID,'M1:46:::0:17').send_keys(str(hit))
                    time.sleep(1)
                    driver.find_element(By.ID,'M1:46:::2:17').click()
                    driver.find_element(By.ID,'M1:50::btn[0]').click()
                time.sleep(1)
                #Campo Fecha real
                driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').click()
                fecha_real = df['FECHA_CAMBIAR'].loc[(df['PROYECTO'] == proy)&(df['CODIGO_HITO'] == hit)].iloc[0]
                if pd.isna(fecha_real) or (fecha_real.date() > datetime.today().date()): fecha = datetime.today().strftime('%d.%m.%Y')
                else: fecha = fecha_real.date().strftime('%d.%m.%Y')
                # print(fecha)
                driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').clear()
                driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').send_keys(fecha)
                time.sleep(1)
                if driver.find_elements(By.XPATH, '//*[contains(text(), "Ajuste la fecha de la etapa con el elemento")]'):
                    raise ValueError('NOK - Ajuste la fecha de la etapa con el elemento PEP') 

                if driver.find_elements(By.XPATH, '//*[contains(text(), "Las fechas reales se encuentran en el futuro")]'):
                    driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').click()
                    fecha = datetime.today().strftime('%d.%m.%Y')
                    driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').clear()
                    driver.find_element(By.ID,'M0:46:1:2:1:2B256::12:17').send_keys(fecha)
            
            time.sleep(1)
            # Boton grabar
            driver.find_element(By.ID,'M0:50::btn[11]-r').click()
            time.sleep(3 if len(hitos) < 5 else 8)
            if driver.find_elements(By.XPATH, '//*[contains(text(), "Proyecto NO es multicliente")]'):
                    raise ValueError('NOK - Proyecto NO es multicliente, el cliente contable deber ser el de proyecto')  
            if driver.find_elements(By.XPATH, '//*[contains(text(), "No se han modificado los datos")]'):
                resultado_bloque = 'OK - No se han modificado los datos'
            else:
                resultado_bloque = 'OK'
            driver.find_element(By.ID,'M1:50::btn[0]').click()
            time.sleep(2)
    except Exception as e:
        # print(e)
        error_str = str(e)
        if "Message:" in error_str:
            message_only = error_str.split("Message:")[1].split("Stacktrace:")[0].strip()
            print(message_only)
        elif 'NOK' in str(e): 
            print(str(e))
            resultado_bloque = str(e)
        else: resultado_bloque = 'NOK'
    finally:
        driver.quit()
    return resultado_bloque

# -------------------------------
# FUNCIÓN DE PROCESO INDIVIDUAL
# -------------------------------
def procesar_proyecto(proyecto, df_proyecto):
    mylogs = configureLogson(__name__,paths['logs_online']+'ma/'+os.path.basename(__file__).rsplit('.', 1)[0]+".log")
    resultados_hito = []
    try:
        hitos = df_proyecto['CODIGO_HITO'].unique().tolist()
        mylogs.info(f"Facturando -> [{proyecto}]:{hitos}")
        procesamiento_bloques = 15
        for i in range(0, len(hitos), procesamiento_bloques):
            bloque = hitos[i:i+procesamiento_bloques]
            mylogs.info(f"Facturando bloque -> [{proyecto}]:{bloque}")
            resultado_bloque = facturarListaHitos(proyecto, bloque)
            if resultado_bloque == 'NOK':
                mylogs.info(f"Reintentando bloque -> [{proyecto}]:{bloque}")
                resultado_bloque = facturarListaHitos(proyecto, bloque)
            if resultado_bloque == 'NOK': resultado_bloque = 'NOK'
            for hito in bloque:
                resultados_hito.append((proyecto, hito, resultado_bloque))
    except Exception as e:
        for hito in df_proyecto['CODIGO_HITO'].unique():
            resultados_hito.append((proyecto, hito, f"NOK - Error: {str(e)}"))
    return (proyecto, resultados_hito)
# -------------------------------
# EJECUCIÓN PARALELA
# -------------------------------
if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    mylogs = configureLogson(__name__,paths['logs_online']+'ma/'+os.path.basename(__file__).rsplit('.', 1)[0]+".log")
    mylogs.info("======================================================")
    mylogs.info("[START-TIME]["+str(start)+"]")
    resultados = []
    if not FLAG_PRUEBA: max_concurrentes = 10
    else: max_concurrentes = 1
    contador_proyecto = 0
    with ProcessPoolExecutor(max_workers=max_concurrentes) as executor:
        futures = [executor.submit(procesar_proyecto, proy, df[df['PROYECTO'] == proy].copy()) for proy in lista_proy]
        for future in as_completed(futures):
            proyecto, resultado = future.result()
            for proyecto, codigo_hito, resultado in resultado:
                df.loc[(df['PROYECTO'] == proyecto) & (df['CODIGO_HITO'] == codigo_hito),'resultado'] = resultado
            contador_proyecto = contador_proyecto+1
            mylogs.info(f"{contador_proyecto}/{len(lista_proy)} : [{proyecto}] -> Resultados por bloques: {df['resultado'].loc[(df['PROYECTO'] == proyecto)].unique().tolist()}")
    len_total = str(len(df))
    mylogs.info(f"Resultados OK: {len(df.loc[(df['resultado'].isin(['OK','OK - No se han modificado los datos']))])}/{len_total}")
    mylogs.info(f"Resultados NOK: {len(df.loc[(df['resultado'].str.contains('NOK'))])}/{len_total}")    # Exportar resultados
    with pd.ExcelWriter(path_result, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='facturar')
        workbook = writer.book
        worksheet = writer.sheets['facturar']
        for col_idx, col in enumerate(df.columns):
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                excel_col = col_idx + 1
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=excel_col)
                    cell.number_format = 'DD/MM/YYYY'
    mylogs.info('Fichero exportado->: '+ path_result)
    mylogs.info(f'[Tiempo por Hito][{"{:.3f}".format(timexHito(start, len(df)))} seconds]')
    mylogs.info("[END-TIME]["+str(datetime.now())+"]")
    mylogs.info('Total runtime: '+ str(totalTime(start, datetime.now())[0])+ ' minutes '+ str("{:.3f}".format(totalTime(start, datetime.now())[1]))+ ' seconds')
    mylogs.info("======================================================")