import urllib.request
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.edge.options import Options
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from undetected_chromedriver import ChromeOptions
import undetected_chromedriver.v2 as uc
import os
import sys
import glob
import time
import traceback
import variables as var

from fake_useragent import UserAgent
from selenium.webdriver.support.ui import Select




def get_driver():
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'

    #options = webdriver.ChromeOptions()
    # specify headless mode
    #options.add_argument('headless')
    # specify the desired user agent
    #options.add_argument(f'user-agent={user_agent}')

    options = uc.ChromeOptions()
    os.chdir('SIB')

    prefs = {'download.default_directory' : os.getcwd() }
    options.add_experimental_option('prefs',prefs)
    #options.headless=True
    #options.add_argument('--headless')
    #options.add_argument(f'user-agent={user_agent}')
    
    driver = uc.Chrome(options=options)
    driver.set_window_position(0, 0)
    driver.set_window_size(400, 700)
    return driver

def get_headless_driver(download_dir):
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    options = Options()
    options.add_argument("headless")
    os.chdir(download_dir)
    prefs = {'download.default_directory' : os.getcwd() }
    options.add_experimental_option('prefs',prefs)
    driver = webdriver.Chrome(executable_path="../chromedriver.exe",options = options)
    driver.set_window_position(0, 0)
    driver.set_window_size(1500, 1500)
    return driver

def get_non_headless_driver(download_dir):
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    options = Options()
    #options.add_argument("headless")
    os.chdir(download_dir)
    prefs = {'download.default_directory' : os.getcwd() }
    options.add_experimental_option('prefs',prefs)

    driver = webdriver.Chrome(executable_path="../chromedriver.exe",options = options)
    driver.set_window_position(0, 0)
    driver.set_window_size(1500, 1500)
    return driver


#################################################################
#                              SIB                              #
#################################################################

### Perfil Financiero de todas las instituciones a una fecha ###
#https://www.sib.gob.gt/ConsultaDinamica/?cons=383
def sib_1(driver,site):
    
    driver.get(site)
    time.sleep(6)

    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='v-button v-widget']//span[@class='v-button-caption']")))

    driver.find_element(By.XPATH, "//div[@class='v-button v-widget']//span[@class='v-button-caption']").click()

    #driver.execute_script("arguments[0].click();", boton)

    WebDriverWait(driver,1000).until(EC.element_to_be_clickable((By.XPATH,"//img[@src='https://www.sib.gob.gt/ConsultaDinamica/VAADIN/themes/consultadinamica/img/excel_new.jpg']")))

    driver.find_element("xpath","//img[@src='https://www.sib.gob.gt/ConsultaDinamica/VAADIN/themes/consultadinamica/img/excel_new.jpg']").click()

    time.sleep(5)
    return True



#################################################################
#                            BANGUAT                            #
#################################################################
def banguat_descarga_excel(driver,site): #Op 0
    driver.get(site)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#content > div > article > div > div.field.field--name-body.field--type-text-with-summary.field--label-hidden.field--item > p:nth-child(3) > a")))
    driver.find_element(By.XPATH, '//*[@id="content"]/div/article/div/div[2]/p[2]/a').click()
    return True

def banguat_descarga_img(driver,site):
    #body > div > img
    driver.get(site)
    grafica = driver.find_element(By.XPATH, '/html/body/div/img')
    src = grafica.get_attribute('src')
    urllib.request.urlretrieve(src, "Imagenes/Balanza Comercial - Comportamiento Semanal.png")
    return True

def banguat_descarga_html(driver,site,i):
    driver.get(site)
    div_ano = driver.find_element(By.XPATH, "/html/body/div[4]/span/b")
    ano = div_ano.text 
    import pandas as pd
    try:
        table = pd.read_html(site)[0]
        writer = pd.ExcelWriter("Exportaciones e importaciones.xlsx", engine='xlsxwriter')
        table.to_excel(writer, sheet_name=ano.replace("AÑO ",""))
        writer.save()
    except Exception:
        traceback.print_exc()
        return False
    return True

#################################################################
#                             BVNSA                             #
#################################################################

########### Comportamiento Tasa Overnight ###########

def bvnsa_comportamiento_tasa_overnight(driver,site):
    driver.get(site)

    boton1 = driver.find_element("id","boton1")

    driver.execute_script("arguments[0].click();", boton1)
    isExist = os.path.exists("./comportamiento_tasa_overnight")

    if not isExist:
        os.makedirs("./comportamiento_tasa_overnight")

    grafica  = driver.find_element("id","grafica")
    src = grafica.get_attribute('src')
    urllib.request.urlretrieve(src, "./comportamiento_tasa_overnight/grafica.png")

    grafica2 = boton1 = driver.find_element("id","grafica2")
    src = grafica2.get_attribute('src')
    urllib.request.urlretrieve(src, "./comportamiento_tasa_overnight/grafica2.png")

    grafica3 = boton1 = driver.find_element("id","grafica3")
    src = grafica3.get_attribute('src')
    urllib.request.urlretrieve(src, "./comportamiento_tasa_overnight/grafica3.png")
    return True
    
def bvnsa_reportos(name,driver,site):
    driver.get(site)
    selector_fecha = var.bvnsa["reportos"]["id"]
    fecha = driver.find_element(By.ID, selector_fecha)
    fecha.click()
    fecha.clear()
    fecha.send_keys(var.bvnsa["reportos"]["fecha"])
    waitFor(By.ID,'boton1',driver)
    click(By.ID,'boton1',driver)

    waitFor(By.CSS_SELECTOR,'#tabla_publicos_wrapper > div.dt-buttons.btn-group.float-right > a.btn.btn-default.buttons-excel.buttons-html5',driver)
    click(By.CSS_SELECTOR,'#tabla_publicos_wrapper > div.dt-buttons.btn-group.float-right > a.btn.btn-default.buttons-excel.buttons-html5',driver)

    waitFor(By.CSS_SELECTOR,'#tabla_privados_wrapper > div.dt-buttons.btn-group.float-right > a.btn.btn-default.buttons-excel.buttons-html5',driver)
    click(By.CSS_SELECTOR,'#tabla_privados_wrapper > div.dt-buttons.btn-group.float-right > a.btn.btn-default.buttons-excel.buttons-html5',driver)

    return True

def bvnsa_resumen_reportos(driver,site):
    driver.get(site)
    fecha = driver.find_element(By.ID, var.bvnsa["resumen_reportos"]["idDel"])
    fecha.click()
    fecha.clear()
    fecha.send_keys(var.bvnsa["resumen_reportos"]["fechaDel"])
    fecha = driver.find_element(By.ID, var.bvnsa["resumen_reportos"]["idAl"])
    fecha.click()
    fecha.clear()
    fecha.send_keys(var.bvnsa["resumen_reportos"]["fechaAl"])
    waitFor(By.ID,'boton1',driver)
    click(By.ID,'boton1',driver)

    waitFor(By.XPATH,'//*[@id="tabla_datos_wrapper"]/div[1]/a[2]',driver)
    click(By.XPATH,'//*[@id="tabla_datos_wrapper"]/div[1]/a[2]',driver)
    return True

def bvnsa_curva_de_rendimiento(driver,site):
    driver.get(site)
    fecha = driver.find_element(By.ID, var.bvnsa["curva_de_rendimiento"]["idFecha"])
    fecha.click()
    fecha.clear()
    fecha.send_keys(var.bvnsa["curva_de_rendimiento"]["fecha"])
    select = Select(driver.find_element(By.ID, 'monedaInforme'))
    if var.bvnsa["curva_de_rendimiento"]["moneda"] == 0:
        select.select_by_value('1')
    else:
        select.select_by_value('2')

    waitFor(By.ID,'boton1',driver)
    click(By.ID,'boton1',driver)

    grafica  = driver.find_element("id","grafica")
    src = grafica.get_attribute('src')
    urllib.request.urlretrieve(src, "CurvaRendimiento.png")

    waitFor(By.XPATH,'//*[@id="tabla_rendimiento_wrapper"]/div[3]/a[2]',driver)
    click(By.XPATH,'//*[@id="tabla_rendimiento_wrapper"]/div[3]/a[2]',driver)

    waitFor(By.XPATH,'//*[@id="tabla_NSS_wrapper"]/div[3]/a[2]',driver)
    click(By.XPATH,'//*[@id="tabla_NSS_wrapper"]/div[3]/a[2]',driver)
    return True

def bvnsa_informe_diario(driver,site):
    driver.get(site)
    selector_fecha = var.bvnsa["informe_diario"]["id"]
    fecha = driver.find_element(By.ID, selector_fecha)
    fecha.click()
    fecha.clear()
    fecha.send_keys(var.bvnsa["informe_diario"]["fecha"])
    waitFor(By.ID,'boton1',driver)
    click(By.ID,'boton1',driver)
    import time
    time.sleep(1)
    import pandas as pd
    try:
        pg_so = driver.page_source
        tables = pd.read_html(pg_so)
        writer = pd.ExcelWriter("Informe Diario.xlsx", engine='xlsxwriter')
        tables[1].to_excel(writer, sheet_name=var.bvnsa["informe_diario"]["fecha"].replace("/","-"))
        writer.save()
    except Exception:
        traceback.print_exc()
        return False
    return True

def bvnsa_titulos_en_oferta(driver,site):
    driver.get(site)
    import pandas as pd
    try:
        pg_so = driver.page_source
        tables = pd.read_html(pg_so)
        writer = pd.ExcelWriter("Titulos en oferta.xlsx", engine='xlsxwriter')
        tables[1].to_excel(writer, sheet_name='TÍTULOS LISTADOS')
        tables[3].to_excel(writer, sheet_name='ACCIONES')
        writer.save()
    except Exception:
        traceback.print_exc()
        return False
    return True

def bvnsa_volumen_negociado(driver,site):
    driver.get(site)
    import pandas as pd
    try:
        pg_so = driver.page_source
        tables = pd.read_html(pg_so)
        writer = pd.ExcelWriter("Volumen negociado.xlsx", engine='xlsxwriter')
        n = 0
        for table in tables:
            table.to_excel(writer, sheet_name=["Quetzales","Dolares"][n])
            n = n +1
        writer.save()
    except Exception:
        traceback.print_exc()
        return False
    return True

def bvnsa_informe_SINEDI(driver,site):
    driver.get(site)
    waitFor(By.XPATH,'//*[@id="tabla_sinedi_wrapper"]/div[4]/a[2]',driver)
    click(By.XPATH,'//*[@id="tabla_sinedi_wrapper"]/div[4]/a[2]',driver)
    return True

def bvnsa_buzon_bursatil(site):
    import requests
    r = requests.get(site, stream=True)

    with open('Buzon bursatil.pdf', 'wb') as f:
        f.write(r.content)
    return True

def waitFor(by,name,driver):
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((by, name)))

def click(by,name,driver):
    driver.find_element(by, name).click()
