import etl_sib
import etl_banguat
import shutil
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import main as descargar
import os
import variables as var
dir_path = os.path.dirname(os.path.realpath(__file__))
import etl_bvnsa
import qdarktheme
import traceback

class MainWindow(QMainWindow):
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
    # "https://banguat.gob.gt/es/page/exportaciones-e-importaciones-mensuales-por-producto",
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

    banguat_tipo = [1, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

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


    def __init__(self):
        super(MainWindow, self).__init__()
        self.layout = QGridLayout()
        self.widget1 = QWidget()
        self.widget2 = QWidget()
        self.widget3 = QWidget()
        self.widget4 = QWidget()


        self.button1 = QPushButton(self.widget2)
        self.button1.setText("Descargar reportes SIB")
        self.button1.move(270,32)
        self.button1.setStyleSheet("QPushButton {background-color: #2B5DD1;color: #FFFFFF;border-style: outset;padding: 2px;font: bold 20px;border-width: 6px;border-radius: 10px; border-color: #2752B8;}QPushButton:hover{background-color: #11275c;}")
        self.button1.clicked.connect(self.button1_clicked)


        x = 70
        y = 90
        for n in self.sib_reporte:
            label = QLabel(self.widget2)
            label.setStyleSheet('font-size:16px')
            label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
            label.setText(n)
            label.move(x,y)
            y = y + 30
            self.labels.append(label)

        self.button_banguat = QPushButton(self.widget3)
        self.button_banguat.setText("Descargar reportes BANGUAT")
        self.button_banguat.move(220,32)
        self.button_banguat.setStyleSheet("QPushButton {background-color: #2B5DD1;color: #FFFFFF;border-style: outset;padding: 2px;font: bold 20px;border-width: 6px;border-radius: 10px; border-color: #2752B8;}QPushButton:hover{background-color: #11275c;}")
        self.button_banguat.clicked.connect(self.banguat_clicked)


        x = 70
        y = 90
        for n in self.banguat_reporte:
            label = QLabel(self.widget3)
            label.setStyleSheet('font-size:16px')
            label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
            label.setText(n)
            label.move(x,y)
            y = y + 30
            self.labels_banguat.append(label)

        self.button_bvnsa = QPushButton(self.widget4)
        self.button_bvnsa.setText("Descargar reportes BVNSA")
        self.button_bvnsa.move(250,32)
        self.button_bvnsa.clicked.connect(self.bvnsa_clicked)
        self.button_bvnsa.setStyleSheet("QPushButton {background-color: #2B5DD1;color: #FFFFFF;border-style: outset;padding: 2px;font: bold 20px;border-width: 6px;border-radius: 10px; border-color: #2752B8;}QPushButton:hover{background-color: #11275c;}")

        x = 70
        y = 90
        for n in self.bvnsa_reporte:
            label = QLabel(self.widget4)
            label.setStyleSheet('font-size:16px')
            label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
            label.setText(n)
            label.move(x,y)
            y = y + 30
            self.labels_bvnsa.append(label)
        
        y = y + 120
        #self.button.setText("Guardar cambios")
        self.button_cambios = QPushButton(self.widget4)
        self.button_cambios.setIcon(QIcon('settings.ico'))
        self.button_cambios.setText("Editar configuraciones")
        self.button_cambios.move(590,y)
        self.ajustes_bvnsa = AjustesBVNSA()
        self.button_cambios.clicked.connect(self.ajustes_bvnsa.show)


        self.button_todos = QPushButton(self.widget1)
        self.button_todos.setText("Descargar todos los reportes")
        self.button_todos.setStyleSheet("QPushButton {background-color: #2B5DD1;color: #FFFFFF;border-style: outset;padding: 2px;font: bold 20px;border-width: 6px;border-radius: 10px; border-color: #2752B8;}QPushButton:hover{background-color: #11275c;}")
        self.button_todos.move(230,50)
        self.button_todos.clicked.connect(self.todos)

        self.textBox = QPlainTextEdit(self.widget1)
        self.textBox.setGeometry(50,150,700,350)
        #self.textBox.setEnabled(False)



        self.tabwidget = QTabWidget(self)
        self.tabwidget.addTab(self.widget1, "Todos")
        self.tabwidget.addTab(self.widget2, "SIB")
        self.tabwidget.addTab(self.widget3, "BANGUAT")
        self.tabwidget.addTab(self.widget4, "BVNSA")
        #layout.addWidget(tabwidget, 0, 0)
        self.tabwidget.setGeometry(0,0,800,600)
        self.resize(800,600)
        self.setWindowTitle("Descarga Automatizada")
        #self.tabwidget.show()


    def button1_clicked(self):
        isExist = os.path.exists("./SIB")
        if isExist:
            shutil.rmtree('./SIB')
        os.makedirs("./SIB")
        driver = descargar.get_driver()
        n = 1
        for label,site in zip(self.labels,self.sib_sites):
            self.imprimir("DESCARGANDO REPORTE " +str(n))
            label.setStyleSheet("background-color: white")
            reporte= descargar.sib_1(driver,site)
            self.imprimir("REPORTE " +str(n) + " DESCARGADO")
            n = n + 1
            label.setStyleSheet("background-color: green;font-size:16px;color: white") if reporte else label.setStyleSheet("background-color: red;font-size:16px;color: white")
        driver.close()
        try:
            etl_sib.ExecuteETL()
        except Exception:
            self.imprimir("Hubo un error en el proceso de ETL SIB")
            traceback.print_exc()

    def banguat_clicked(self):
        isExist = os.path.exists("./BANGUAT")
        if isExist:
            shutil.rmtree('./BANGUAT')
        os.makedirs("./BANGUAT")
        driver = descargar.get_headless_driver("BANGUAT")
        n = 1
        for label,site,tipo in zip(self.labels_banguat,self.banguat_sites,self.banguat_tipo):
            self.imprimir("DESCARGANDO REPORTE " +str(n))
            label.setStyleSheet("background-color: white")
            reporte = False
            if tipo == 0: reporte = descargar.banguat_descarga_excel(driver,site)
            elif tipo == 1: reporte = descargar.banguat_descarga_html(driver,site,n)
            elif tipo == 2: reporte = descargar.banguat_descarga_img(driver,site)
            self.imprimir("REPORTE " +str(n) + " DESCARGADO")
            n = n + 1
            label.setStyleSheet("background-color: green;font-size:16px;color: white") if reporte else label.setStyleSheet("background-color: red;font-size:16px;color: white")
        driver.close()
        try:
            etl_banguat.ExecuteETL()
        except Exception:
            self.imprimir("Hubo un error en el proceso de ETL BANGUAT")
            traceback.print_exc()

    def bvnsa_clicked(self):
        isExist = os.path.exists("./BVNSA")
        if isExist:
            shutil.rmtree('./BVNSA')
        os.makedirs("./BVNSA")
        driver = descargar.get_headless_driver("BVNSA")
        n = 1
        for label,site,nombre in zip(self.labels_bvnsa,self.bvnsa_sites,self.bvnsa_reporte):
            self.imprimir("DESCARGANDO REPORTE " + nombre)
            if False and (n<4 or n>4):
                n=n+1
                continue
            label.setStyleSheet("background-color: white")
            reporte = False
            if n == 1 : reporte = descargar.bvnsa_comportamiento_tasa_overnight(driver,site)
            elif n == 2: reporte = descargar.bvnsa_reportos("reportos",driver,site)
            elif n == 3: reporte = descargar.bvnsa_resumen_reportos(driver,site)
            elif n == 4: reporte = True
            elif n == 5: reporte = descargar.bvnsa_curva_de_rendimiento(driver,site)
            elif n == 6: reporte = descargar.bvnsa_informe_diario(driver,site)
            elif n == 7: reporte = descargar.bvnsa_titulos_en_oferta(driver,site)
            elif n == 8: reporte = descargar.bvnsa_volumen_negociado(driver,site)
            elif n == 9: reporte = descargar.bvnsa_informe_SINEDI(driver,site)
            elif n == 10: reporte = descargar.bvnsa_buzon_bursatil(site)
            self.imprimir("REPORTE " + nombre + " DESCARGADO")
            n = n + 1
            label.setStyleSheet("background-color: green;font-size:16px;color: white") if reporte else label.setStyleSheet("background-color: red;font-size:16px;color: white")
        import time
        time.sleep(1)
        driver.close()
        #try:
        #    etl_bvnsa.ExecuteETL()
        #except Exception:
        #    self.imprimir("Hubo un error en el proceso de ETL BVNSA")
        #    traceback.print_exc()

    def todos(self):
        print("todos")
        self.imprimir("----------------SIB----------------")
        self.button1_clicked()
        self.imprimir("----------------FIN SIB----------------")
        self.imprimir("----------------BANGUAT----------------")
        self.banguat_clicked()
        self.imprimir("----------------FIN BANGUAT----------------")
        self.imprimir("----------------BVNSA----------------")
        self.bvnsa_clicked()
        self.imprimir("----------------FIN BVNSA----------------")

    def imprimir(self,texto):
        self.textBox.setPlainText(self.textBox.toPlainText() + texto + "\n")
        print(texto)

class AjustesBVNSA(QWidget):
    def __init__(self):
        super(AjustesBVNSA, self).__init__()
        self.resize(550, 300)
        
        self.setWindowTitle("Configuraciones BVNSA")

        # Label
        self.label = QLabel(self)
        self.label.setGeometry(0, 0, 500, 50)
        self.label.setText('Ajustes para descarga de reportes BVNSA')
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet('font-size:20px')

        x = 40
        y = 60
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y, 150, 30)
        self.label_reportos.setText('Reportos:')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:16px')
        x = x + 100
        self.reportos_fecha = QDateEdit(self,calendarPopup=True)
        self.reportos_fecha.setGeometry(x, y, 100, 20)
        d = QDate(int(var.bvnsa["reportos"]["fecha"].split("/")[2]), int(var.bvnsa["reportos"]["fecha"].split("/")[1]), int(var.bvnsa["reportos"]["fecha"].split("/")[0]))
        self.reportos_fecha.setDate(d)
        x = x - 100
        y = y + 40
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y, 175, 30)
        self.label_reportos.setText('Resumen reportos:  Del')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:16px')
        x = x + 180
        self.resumen_reportos_fecha_del = QDateEdit(self,calendarPopup=True)
        self.resumen_reportos_fecha_del.setGeometry(x, y, 100, 20)
        d = QDate(int(var.bvnsa["resumen_reportos"]["fechaDel"].split("/")[2]), int(var.bvnsa["resumen_reportos"]["fechaDel"].split("/")[1]), int(var.bvnsa["resumen_reportos"]["fechaDel"].split("/")[0]))
        self.resumen_reportos_fecha_del.setDate(d)
        x = x + 110
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y, 100, 40)
        self.label_reportos.setText('al')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:16px')
        x = x + 25
        self.resumen_reportos_fecha_al = QDateEdit(self,calendarPopup=True)
        self.resumen_reportos_fecha_al.setGeometry(x, y, 100, 20)
        d = QDate(int(var.bvnsa["resumen_reportos"]["fechaAl"].split("/")[2]), int(var.bvnsa["resumen_reportos"]["fechaAl"].split("/")[1]), int(var.bvnsa["resumen_reportos"]["fechaAl"].split("/")[0]))
        self.resumen_reportos_fecha_al.setDate(d)

        x = 40
        y = y + 40
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y, 160, 40)
        self.label_reportos.setText('Curva de rendimiento:')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:16px')
        x = x + 170
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y+2, 50, 30)
        self.label_reportos.setText('Fecha:')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:12px')
        x = x + 45
        self.curva_rendimiento_fecha = QDateEdit(self,calendarPopup=True)
        self.curva_rendimiento_fecha.setGeometry(x, y, 100, 20)
        d = QDate(int(var.bvnsa["curva_de_rendimiento"]["fecha"].split("/")[2]), int(var.bvnsa["curva_de_rendimiento"]["fecha"].split("/")[1]), int(var.bvnsa["curva_de_rendimiento"]["fecha"].split("/")[0]))
        self.curva_rendimiento_fecha.setDate(d)
        x = x + 110
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y+2, 60, 40)
        self.label_reportos.setText('Moneda:')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:12px')
        x = x + 60
        self.combobox_moneda = QComboBox(self)
        self.combobox_moneda.addItems(var.cr_moneda)
        self.combobox_moneda.setGeometry(x,y,100,25)
        self.combobox_moneda.setCurrentIndex(var.bvnsa["curva_de_rendimiento"]["moneda"])
  
        x = 40
        y = y + 40
        self.label_reportos = QLabel(self)
        self.label_reportos.setGeometry(x, y, 150, 30)
        self.label_reportos.setText('Informe diario:')
        self.label_reportos.setAlignment(Qt.AlignLeft)
        self.label_reportos.setStyleSheet('font-size:16px')
        x = x + 150
        self.informe_diario_fecha = QDateEdit(self,calendarPopup=True)
        self.informe_diario_fecha.setGeometry(x, y, 100, 20)
        d = QDate(int(var.bvnsa["informe_diario"]["fecha"].split("/")[2]), int(var.bvnsa["informe_diario"]["fecha"].split("/")[1]), int(var.bvnsa["informe_diario"]["fecha"].split("/")[0]))
        self.informe_diario_fecha.setDate(d)

        y = y + 50
        self.button = QPushButton(self)
        self.button.setText("Guardar cambios")
        self.button.move(60,y)
        self.button.clicked.connect(self.guardar_cambios)

    def guardar_cambios(self):
        ar_fecha_reportos = str(self.reportos_fecha.date().toPyDate()).split("-")
        var.bvnsa["reportos"]["fecha"] = ar_fecha_reportos[2] + "/" + ar_fecha_reportos[1] + "/" + ar_fecha_reportos[0]
        self.close()




if __name__ == '__main__':
    app = QApplication([])
    app.setStyleSheet(qdarktheme.load_stylesheet("light"))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())