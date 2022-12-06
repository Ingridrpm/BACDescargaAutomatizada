from pandas.tseries.offsets import BMonthEnd
from datetime import date

d=date.today()
offset = BMonthEnd()

cr_moneda = ["Quetzales","Dolares"]

bvnsa = {
    "reportos" : {
        "id":"fechaInformeDel",
        "fecha":"17/11/2022"#str(offset.rollforward(d))
    },
    "resumen_reportos": {
        "idDel":"fechaInformeDel",
        "idAl":"fechaInformeAl",
        "fechaDel": "01/11/2022",
        "fechaAl": "01/12/2022"
    },
    "curva_de_rendimiento":{
        "idMoneda":"select2-monedaInforme-container",
        "idFecha":"fechaInformeDel",
        "moneda": 1,
        "fecha":"17/11/2022"
    },
    "informe_diario":{
        "id":"fechaInformeDel",
        "fecha":"17/11/2022"#str(offset.rollforward(d))
    }
}