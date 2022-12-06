import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel,QFrame
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import main as descargar
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import os 
import variables as var
dir_path = os.path.dirname(os.path.realpath(__file__))
import shutil

sib_reporte = ["Perfil Financiero de todas las instituciones a una fecha",
"Tasas promedio ponderada de cartera de créditos por agrupaciones",
"Cartera de Créditos por Agrupación",
"Entidades Fuera de Plaza - Offshore Estado de Resultados y Balance General",
"Entidades Fuera de Plaza - Offshore Tasas de Interés",
"Entidades Fuera de Plaza - Offshore Saldos de Cartera",
"Tasas de Interés Aplicadas (Mensuales)",
"Indicadores Financieros del Sistema",
"Balance General y Estado de Resultados de las Instituciones a una Fecha",
"Tasas de Interés Aplicadas (Semanales)",
"Cartera de Créditos Clasificados GD y NGD",
"Balance General y Estado de Resultados de las Entidades a una fecha",
"Tasa promedio ponderada por Actividad Económica"]

sib_sites = ["https://www.sib.gob.gt/ConsultaDinamica/?cons=383",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=20",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=26",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=198",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=200",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=205",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=18",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=384",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=7",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=19",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=412",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=210",
"https://www.sib.gob.gt/ConsultaDinamica/?cons=21"]
labels = []

banguat_sites = [
#"https://banguat.gob.gt/es/page/exportaciones-e-importaciones-mensuales-por-producto",
"https://banguat.gob.gt/sites/default/files/banguat/estaeco/comercio/por_producto/prod_mensDB001.HTM",
"https://banguat.gob.gt/es/page/ingresos-y-egresos-de-divisas",
"https://banguat.gob.gt/es/page/balanza-comercial-comportamiento-semanal-grafica",
"https://banguat.gob.gt/es/page/ingreso-de-divisas-por-exportaciones-totales",
"https://banguat.gob.gt/es/page/oferta-y-demanda-del-mercado-institucional-de-divisas",
"https://banguat.gob.gt/page/emision-monetaria",
"https://banguat.gob.gt/page/posicion-neta-del-banco-de-guatemala-con-el-sector-publico",	
"https://banguat.gob.gt/page/al-sector-publico-neto",
"https://banguat.gob.gt/page/en-moneda-extranjera",
"https://banguat.gob.gt/page/base-monetaria",
"https://banguat.gob.gt/page/numerario-en-circulacion",
"https://banguat.gob.gt/page/medio-circulante",
"https://banguat.gob.gt/page/monetarios-en-moneda-nacional	",
"https://banguat.gob.gt/es/page/en-moneda-nacional-3",
"https://banguat.gob.gt/page/monetarios-en-moneda-extranjera"]

banguat_reporte = ["Exportaciones e importaciones mensuales por producto",
"Ingresos y Egresos de Divisas",
"Balanza Comercial, Comportamiento Semanal ",
"Ingreso de Divisas por Exportaciones Totales",
"Oferta y Demanda del Mercado Institucional de Divisas",
"Emisión Monetaria",
"Posición neta del Banco de Guatemala con el sector público",	
"Credito Bancario al sector público neto ",
"Credito Bancario al sector privado en Moneda Nacional ",
"Base Monetaria",
"Numerario en circulación",
"Medio Circulante",
"Depósitos Monetarios en los Bancos (Moneda Nacional)",
"Medios de Pago (Moneda Nacional)",
"Depósitos Monetarios en los Bancos (Moneda Extanjera)"]

banguat_tipo = [1,0,2,0,0,0,0,0,0,0,0,0,0,0,0]

labels_banguat = []

labels_bvnsa = []

bvnsa_reporte = ["Comportamiento Tasa Overnight",
"Reportos",
"Resumen Reportos",
"Mercado Primario - Licitaciones (Instrumentos Públicos)",
"Curva de Rendimiento",
"Informe diario",
"Títulos en Oferta",
"Volumen Negociado",
"Informe SINEDI",
"Buzón Bursátil"]

bvnsa_sites = ["http://www.bvnsa.com.gt/bvnsa/informes_tasa_overnight.php",
"http://www.bvnsa.com.gt/bvnsa/informes_reportos.php",
"http://www.bvnsa.com.gt/bvnsa/informes_resumen_reportos.php",
"http://www.bvnsa.com.gt/bvnsa/informes_primario_licitaciones.php",
"http://www.bvnsa.com.gt/bvnsa/informes_curva_rendimiento.php",
"http://www.bvnsa.com.gt/bvnsa/informes_informe_diario.php",
"http://www.bvnsa.com.gt/bvnsa/informes_titulos_oferta.php",
"http://www.bvnsa.com.gt/bvnsa/informes_volumen_negociado.php",
"http://www.bvnsa.com.gt/bvnsa/informes_sinedi.php",
"http://www.bvnsa.com.gt/WPaginas/buzon/buzonbursatil_2022.08.pdf"]


def window():
   app = QApplication(sys.argv)
   widget = QWidget()
   layout = QGridLayout()
   widget1 = QWidget()
   widget2 = QWidget()
   widget3 = QWidget()
   widget4 = QWidget()

   button1 = QPushButton(widget2)
   button1.setText("Descargar reportes SIB")
   button1.move(70,32)
   button1.clicked.connect(button1_clicked)


   x = 70
   y = 70
   for n in sib_reporte:
    label = QLabel(widget2)
    label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
    label.setText(n)
    label.move(x,y)
    y = y + 25
    labels.append(label)

   button_banguat = QPushButton(widget3)
   button_banguat.setText("Descargar reportes BANGUAT")
   button_banguat.move(70,32)
   button_banguat.clicked.connect(banguat_clicked)


   x = 70
   y = 70
   for n in banguat_reporte:
    label = QLabel(widget3)
    label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
    label.setText(n)
    label.move(x,y)
    y = y + 25
    labels_banguat.append(label)

   button_bvnsa = QPushButton(widget4)
   button_bvnsa.setText("Descargar BVNSA Comportamiento Tasa Overnight")
   button_bvnsa.move(70,32)
   button_bvnsa.clicked.connect(bvnsa_clicked)
   x = 70
   y = 70
   for n in bvnsa_reporte:
    label = QLabel(widget4)
    label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
    label.setText(n)
    label.move(x,y)
    y = y + 25
    labels_bvnsa.append(label)
   
   tabwidget = QTabWidget()
   tabwidget.addTab(widget1, "Todos")
   tabwidget.addTab(widget2, "SIB")
   tabwidget.addTab(widget3, "BANGUAT")
   tabwidget.addTab(widget4, "BVNSA")
   tabwidget.setLayout(layout)
   layout.addWidget(tabwidget, 0, 0)
   tabwidget.setGeometry(600,50,600,500)
   tabwidget.setWindowTitle("TEST1")
   tabwidget.show()
   sys.exit(app.exec_())


def button1_clicked():
    driver = descargar.get_driver()
    n = 1
    for label,site in zip(labels,sib_sites):
        print("DESCARGANDO REPORTE " +str(n))
        label.setStyleSheet("background-color: white")
        reporte1 = descargar.sib_1(driver,site)
        print("REPORTE " +str(n) + " DESCARGADO")
        n = n + 1
        label.setStyleSheet("background-color: green") if reporte1 else label.setStyleSheet("background-color: red")
    driver.close()

def banguat_clicked():
    driver = descargar.get_headless_driver("BANGUAT")
    n = 1
    for label,site,tipo in zip(labels_banguat,banguat_sites,banguat_tipo):
        print("DESCARGANDO REPORTE " +str(n))
        label.setStyleSheet("background-color: white")
        reporte = False
        if tipo == 0: reporte = descargar.banguat_descarga_excel(driver,site)
        elif tipo == 1: reporte = descargar.banguat_descarga_html(driver,site,n)
        elif tipo == 2: print("ES IMAGEN")
        print("REPORTE " +str(n) + " DESCARGADO")
        n = n + 1
        label.setStyleSheet("background-color: green") if reporte else label.setStyleSheet("background-color: red")
    driver.close()

def bvnsa_clicked():
    isExist = os.path.exists("./BVNSA")
    if isExist:
        shutil.rmtree('./BVNSA')
    os.makedirs("./BVNSA")
    driver = descargar.get_non_headless_driver("BVNSA")
    n = 1
    for label,site,nombre in zip(labels_bvnsa,bvnsa_sites,bvnsa_reporte):
        print("DESCARGANDO REPORTE " + nombre)
        if False and (n<5 or n>6):
            n=n+1
            continue
        label.setStyleSheet("background-color: white")
        reporte = False
        if n == 1 : reporte = descargar.bvnsa_comportamiento_tasa_overnight(driver,site)
        elif n == 2: reporte = descargar.bvnsa_reportos("reportos",driver,site)
        elif n == 3: reporte = descargar.bvnsa_resumen_reportos(driver,site)
        elif n == 4: print("PENDIENTE")
        elif n == 5: reporte = descargar.bvnsa_curva_de_rendimiento(driver,site)
        elif n == 6: reporte = descargar.bvnsa_informe_diario(driver,site)
        elif n == 7: reporte = descargar.bvnsa_titulos_en_oferta(driver,site)
        elif n == 8: reporte = descargar.bvnsa_volumen_negociado(driver,site)
        elif n == 9: reporte = descargar.bvnsa_informe_SINEDI(driver,site)
        elif n == 10: reporte = descargar.bvnsa_buzon_bursatil(site)
        #elif tipo == 1: reporte = descargar.banguat_descarga_html(driver,site,n)
        #elif tipo == 2: print("ES IMAGEN")
        print("REPORTE " + nombre + " DESCARGADO")
        n = n + 1
        label.setStyleSheet("background-color: green") if reporte else label.setStyleSheet("background-color: red")
    import time
    time.sleep(1)
    driver.close()
    
if __name__ == '__main__':
   window()