from pandas.tseries.offsets import BMonthEnd
from datetime import *

d=datetime.now()
offset = BMonthEnd()

cr_moneda = ["Quetzales","Dolares"]

bvnsa = {
    "reportos" : {
        "id":"fechaInformeDel",
        "fecha": d.strftime("%d/%m/%Y")
    },
    "resumen_reportos": {
        "idDel":"fechaInformeDel",
        "idAl":"fechaInformeAl",
        "fechaDel": d.strftime("01/%m/%Y"),
        "fechaAl": d.strftime("%d/%m/%Y")
    },
    "curva_de_rendimiento":{
        "idMoneda":"select2-monedaInforme-container",
        "idFecha":"fechaInformeDel",
        "moneda": 0,
        "fecha": d.strftime("%d/%m/%Y")
    },
    "informe_diario":{
        "id":"fechaInformeDel",
        "fecha":d.strftime("%d/%m/%Y")
    },
    "mercado_primario":{
        "titulo":{
            "id":"tituloInforme",
            "values":[2,3,4,8],
            "opciones":["CERTIBONOS","CDP","CDP$","CERTIBONO$"]
        },
        "parametro":{
            "id":"Param_Graficar",
            "values":[1,2,3,4,5],
            "opciones":["Precio","Rendimiento","Precio/Rendimiento","Precio/Volumen","Rendimiento/Volumen"]
        },
        "vencimiento":[[],[],["20/04/2023","15/05/2024","16/09/2024","18/11/2024","21/07/2025","02/10/2025","15/04/2026","19/05/2027","15/11/2027","15/12/2027","21/02/2028","28/03/2028","23/05/2029","02/08/2029","18/03/2030","23/09/2030","25/06/2031","10/07/2031","18/08/2031","21/10/2031","27/04/2032","21/07/2032","15/03/2033","18/04/2033","27/06/2034","20/07/2034","21/08/2034","15/10/2035","15/11/2035","19/06/2036","22/09/2036","22/09/2037","15/02/2038","17/05/2039","23/11/2039","20/08/2040","05/09/2040","15/08/2041","20/11/2041", "20/08/2042"],
        ["06/03/2023","05/06/2023","04/09/2023","04/12/2023","04/03/2024","03/06/2024","02/09/2024","01/12/2025","01/12/2026","07/12/2027","03/12/2030"],[],[],[],[],
        ["15/08/2023","15/02/2027","25/07/2035","20/11/2036","18/11/2037"]],
        "duracion":{
            "id":"duracion",
            "values":[90,180,360,720,1800],
            "options":["3 Meses","6 Meses", "1 Año","2 Años","5 Años"]
        }
    }
}