"""Genera plantilla/Plantilla_control_rendimiento.xlsx: registro por orden con balance de pérdidas.

Basada en el formato "Análisis de Rendimiento metálico 2010" (T1, T2, T3, T4, fuera de medida,
corto T4, oxidación, barras perdidas, pérdidas SAP, delta, hecho/causa/acción), con resúmenes,
carta de control semanal, Pareto y plan de acción que se calculan con fórmulas.

Uso:
    python3 scripts/crear_plantilla.py
"""
import os
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = Path(os.environ.get("PLANTILLA_SALIDA", RAIZ / "plantilla" / "Plantilla_control_rendimiento.xlsx"))

ROJO = "C8102E"
GRAFITO = "2B2B2B"
F = "Arial"
f_base = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_titulo = Font(name=F, size=14, bold=True, color=GRAFITO)
f_sub = Font(name=F, size=11, bold=True, color=ROJO)
f_nota = Font(name=F, size=9, italic=True, color="6F6E69")
f_input = Font(name=F, size=10, color="0000FF")
fill_head = PatternFill("solid", fgColor=ROJO)
fill_calc = PatternFill("solid", fgColor="5F5E5A")
fill_sub = PatternFill("solid", fgColor="F0EFEC")
fill_input = PatternFill("solid", fgColor="FFFF00")
fill_alerta = PatternFill("solid", fgColor="F8D7DA")
fill_ok = PatternFill("solid", fgColor="DDEFE0")
fill_aviso = PatternFill("solid", fgColor="FCE5D8")
borde = Border(bottom=Side(style="thin", color="E1E0D9"))

PCT = "0.00%"
KG = "#,##0;-#,##0;-"
PP = '+0.00" pp";-0.00" pp";0.00" pp"'
FECHA = "dd-mm-yyyy"

# Metas: Rend_Met del CSV de consumo 2026, ponderadas por kg consumidos (ver README).
META_GENERAL = 0.958
GRUPOS = [("B Hormigón", 0.9576), ("Redondo liso", 0.9538), ("Plana", 0.9609), ("Ángulo", 0.9551),
          ("Cuadrado", 0.9648), ("Estrella", None), ("Saferock", 0.9584), ("Hexágono", 0.9444)]
TURNOS = ["A", "B", "C", "D"]
CAUSAS = ["Palanquilla (defecto / largo)", "Horno (oxidación / temperatura)", "Cizallas (despunte)",
          "Cobles / barras perdidas", "Largo de corte / padrón", "Eléctrico / instrumentación",
          "Mecánico / guías / cilindros", "Cambio de medida / campaña corta", "Operacional / práctica",
          "Registro / información"]
ESTADOS = ["Abierta", "En curso", "Cerrada", "Cerrada - no eficaz"]
TIPOS_ACCION = ["Contención", "Correctiva", "Preventiva"]

N_REG = 1000          # filas de órdenes disponibles
R0 = 6                # primera fila de datos en Registro
RN = R0 + N_REG - 1
N_ACC = 200


def encabezado(ws, fila, textos, col0=1, calc=()):
    for j, t in enumerate(textos):
        c = ws.cell(fila, col0 + j, t)
        c.font = f_head
        c.fill = fill_calc if (col0 + j) in calc else fill_head
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 42


def anchos(ws, valores):
    for j, w in enumerate(valores, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def titulo(ws, texto, sub=None):
    ws["A1"] = texto
    ws["A1"].font = f_titulo
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f_nota


def entrada(c, valor=None, fmt=None):
    if valor is not None:
        c.value = valor
    c.font, c.fill = f_input, fill_input
    if fmt:
        c.number_format = fmt


def lista(ws, rango_origen, celdas):
    dv = DataValidation(type="list", formula1=rango_origen, allow_blank=True)
    dv.error, dv.errorTitle = "Elija un valor de la lista (hoja Parámetros).", "Valor no válido"
    ws.add_data_validation(dv)
    dv.add(celdas)


def fuente_arial(ws):
    for fila in ws.iter_rows():
        for c in fila:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic,
                              color=c.font.color)


def texto_largo(ws, fila, texto, font=f_base, col=1, hasta=12, alto=None):
    ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=hasta)
    c = ws.cell(fila, col, texto)
    c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if alto:
        ws.row_dimensions[fila].height = alto


# ---------------------------------------------------------------- Registro (columnas)
COLS = [
    # (clave, encabezado, ancho, tipo, formato)  tipo: in = dato, f = fórmula
    ("fecha", "Fecha", 11, "in", FECHA),
    ("anio", "Año ISO", 7, "f", "0"),
    ("sem", "Semana ISO", 7, "f", "0"),
    ("mes", "Mes", 9, "f", "mmm-yy"),
    ("turno", "Turno", 6, "in", None),
    ("orden", "Orden", 11, "in", "0"),
    ("cod", "Código", 9, "in", "0"),
    ("mat", "Material", 34, "in", None),
    ("grupo", "Grupo", 13, "in", None),
    ("largo", "Largo palanquilla (m)", 10, "in", "0.00"),
    ("npal", "N° palanquillas", 10, "in", "0"),
    ("horas", "Horas", 7, "in", "0.00"),
    ("cons", "Consumo (kg)", 11, "in", KG),
    ("prod", "Producción (kg)", 11, "in", KG),
    ("rend", "Rend. metálico", 9, "f", PCT),
    ("meta", "Meta", 8, "f", PCT),
    ("brecha", "Brecha vs meta", 9, "f", PP),
    ("falt", "Kg faltantes vs meta", 10, "f", KG),
    ("prodv", "Productividad (t/h)", 10, "f", "0.00"),
    ("t1", "T1 (kg)", 8, "in", KG),
    ("t2", "T2 (kg)", 8, "in", KG),
    ("t3", "T3 (kg)", 8, "in", KG),
    ("t4", "T4 (kg)", 8, "in", KG),
    ("fm", "Fuera de medida (kg)", 9, "in", KG),
    ("corto", "Corto T4 (kg)", 8, "in", KG),
    ("cobles", "Cobles / barras perdidas (kg)", 10, "in", KG),
    ("nbarras", "N° barras perdidas", 8, "in", "0"),
    ("oxm", "Oxidación medida (kg)", 10, "in", KG),
    ("ox", "Oxidación usada (kg)", 10, "f", KG),
    ("medidas", "Pérdidas medidas (kg)", 10, "f", KG),
    ("sap", "Pérdidas SAP (kg)", 10, "f", KG),
    ("delta", "Delta no explicado (kg)", 10, "f", KG),
    ("pdelta", "Delta / consumo", 8, "f", PCT),
    ("kgton", "Pérdidas (kg/t prod.)", 9, "f", "0.0"),
    ("cxm", "aux: consumo × meta", 11, "f", KG),
    ("estado", "Estado", 15, "f", None),
    ("hecho", "Hecho (qué pasó)", 30, "in", None),
    ("causa", "Causa (categoría)", 24, "in", None),
    ("detalle", "Causa (detalle)", 30, "in", None),
    ("accion", "Acción", 30, "in", None),
    ("resp", "Responsable", 13, "in", None),
]
C = {k: get_column_letter(i) for i, (k, *_r) in enumerate(COLS, 1)}
PERDIDAS = [("T1", "t1"), ("T2", "t2"), ("T3", "t3"), ("T4", "t4"), ("Fuera de medida", "fm"),
            ("Corto T4", "corto"), ("Cobles / barras perdidas", "cobles"), ("Oxidación", "ox"),
            ("Delta no explicado", "delta")]

REG = "Registro"
PAR = "'Parámetros'"


def rango(clave):
    return f"{REG}!${C[clave]}${R0}:${C[clave]}${RN}"


def formula_registro(k, r):
    c = {key: f"{col}{r}" for key, col in C.items()}
    vacio = f'{c["cons"]}=""'
    return {
        "anio": f'=IF({c["fecha"]}="","",YEAR({c["fecha"]}-WEEKDAY({c["fecha"]},2)+4))',
        "sem": f'=IF({c["fecha"]}="","",WEEKNUM({c["fecha"]},21))',
        "mes": f'=IF({c["fecha"]}="","",DATE(YEAR({c["fecha"]}),MONTH({c["fecha"]}),1))',
        "rend": f'=IF(OR({vacio},{c["cons"]}=0,{c["prod"]}=""),"",{c["prod"]}/{c["cons"]})',
        "meta": f'=IF({vacio},"",IFERROR(IF(INDEX({PAR}!$B$14:$B$21,MATCH({c["grupo"]},{PAR}!$A$14:$A$21,0))="",'
                f'{PAR}!$B$5,INDEX({PAR}!$B$14:$B$21,MATCH({c["grupo"]},{PAR}!$A$14:$A$21,0))),{PAR}!$B$5))',
        "brecha": f'=IF({c["rend"]}="","",({c["rend"]}-{c["meta"]})*100)',
        "falt": f'=IF({c["rend"]}="","",MAX(0,{c["cons"]}*({c["meta"]}-{c["rend"]})))',
        "prodv": f'=IF(OR({c["horas"]}="",{c["horas"]}=0,{c["prod"]}=""),"",{c["prod"]}/1000/{c["horas"]})',
        "ox": f'=IF({vacio},"",IF({c["oxm"]}="",{c["cons"]}*{PAR}!$B$6,{c["oxm"]}))',
        "medidas": f'=IF({vacio},"",SUM({c["t1"]}:{c["cobles"]})+{c["ox"]})',
        "sap": f'=IF(OR({vacio},{c["prod"]}=""),"",{c["cons"]}-{c["prod"]})',
        "delta": f'=IF({c["sap"]}="","",{c["sap"]}-{c["medidas"]})',
        "pdelta": f'=IF(OR({c["delta"]}="",{c["cons"]}=0),"",{c["delta"]}/{c["cons"]})',
        "kgton": f'=IF(OR({c["sap"]}="",{c["prod"]}=0),"",{c["sap"]}/({c["prod"]}/1000))',
        "cxm": f'=IF({c["rend"]}="","",{c["cons"]}*{c["meta"]})',
        "estado": f'=IF({c["rend"]}="","",IF({c["rend"]}>1,"Error: rend > 100 %",'
                  f'IF(SUM({c["t1"]}:{c["cobles"]})=0,"Falta registrar pérdidas",'
                  f'IF(ABS({c["pdelta"]})>{PAR}!$B$7,"Revisar balance",'
                  f'IF({c["rend"]}<{c["meta"]},"Bajo meta","OK")))))',
    }[k]


def hoja_registro(wb):
    ws = wb.create_sheet(REG)
    titulo(ws, "Registro de rendimiento metálico por orden",
           "Una fila por orden de laminación. Amarillo = dato a ingresar; gris = fórmula (no editar). "
           "Pérdidas en kg. La fila 6 es un ejemplo (orden 1008027, noviembre 2010): reemplácela.")
    ws["A3"] = "Balance: Pérdidas SAP = Consumo − Producción.  Delta = Pérdidas SAP − Pérdidas medidas " \
               "(T1…Cobles + Oxidación).  Un delta grande indica pérdidas que no se están pesando o registrando."
    ws["A3"].font = f_nota
    bloques = [("IDENTIFICACIÓN", "fecha", "npal"), ("PRODUCCIÓN", "horas", "prodv"),
               ("PÉRDIDAS MEDIDAS (kg)", "t1", "medidas"), ("BALANCE", "sap", "estado"),
               ("ANÁLISIS DE LA DESVIACIÓN", "hecho", "resp")]
    idx = {k: i for i, (k, *_r) in enumerate(COLS, 1)}
    for texto, a, b in bloques:
        ws.merge_cells(start_row=4, start_column=idx[a], end_row=4, end_column=idx[b])
        c = ws.cell(4, idx[a], texto)
        c.font, c.fill = f_bold, fill_sub
        c.alignment = Alignment(horizontal="center")
    calc = {i for i, (_k, _h, _w, t, _f) in enumerate(COLS, 1) if t == "f"}
    encabezado(ws, 5, [h for _k, h, *_r in COLS], calc=calc)
    anchos(ws, [w for _k, _h, w, *_r in COLS])

    for r in range(R0, RN + 1):
        for i, (k, _h, _w, t, fmt) in enumerate(COLS, 1):
            cel = ws.cell(r, i)
            if t == "f":
                cel.value = formula_registro(k, r)
                cel.font = f_base
            else:
                cel.font, cel.fill = f_input, fill_input
            if fmt:
                cel.number_format = fmt
            if k in ("hecho", "detalle", "accion"):
                cel.alignment = Alignment(wrap_text=False)

    # fila de ejemplo (datos reales del archivo 2010, orden 1008027, noviembre)
    ej = {"fecha": date(2010, 11, 2), "turno": "A", "orden": 1008027, "cod": 3027,
          "mat": "Plana 50x06 mm 6 m Comercial (N)", "grupo": "Plana", "largo": 3.2, "npal": 356,
          "horas": 8.23, "cons": 139814, "prod": 131826, "t1": 2000, "t2": 395, "t3": 0, "t4": 1140,
          "fm": 0, "corto": 1730, "cobles": 390, "nbarras": 1,
          "hecho": "Cobles (trozo de coble y barra de sacrificio)", "causa": "Eléctrico / instrumentación",
          "detalle": "HMD A5-A6 pierde señal",
          "accion": "Chequear señales de fotoceldas y HMD, graficar en IBA", "resp": "Jefe de turno"}
    for k, v in ej.items():
        ws[f"{C[k]}{R0}"] = v
    ws[f"{C['mat']}{R0}"].comment = Comment(
        "Ejemplo tomado de 'Rend. metálico Noviembre 2010'. N° palanquillas, horas y kg de cobles "
        "son supuestos para mostrar el formato. Borre o sobrescriba esta fila.", "Plantilla")

    ws.freeze_panes = ws[f"{C['mat']}{R0}"]
    ws.auto_filter.ref = f"A5:{C['resp']}{RN}"
    ws.column_dimensions[C["cxm"]].hidden = True

    lista(ws, f"={PAR}!$A$14:$A$21", f"{C['grupo']}{R0}:{C['grupo']}{RN}")
    lista(ws, f"={PAR}!$D$14:$D$17", f"{C['turno']}{R0}:{C['turno']}{RN}")
    lista(ws, f"={PAR}!$F$14:$F$23", f"{C['causa']}{R0}:{C['causa']}{RN}")
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv.error = "Ingrese un número ≥ 0 (kg)."
    ws.add_data_validation(dv)
    for k in ("cons", "prod", "t1", "t2", "t3", "t4", "fm", "corto", "cobles", "oxm", "horas"):
        dv.add(f"{C[k]}{R0}:{C[k]}{RN}")

    est = f"{C['estado']}{R0}:{C['estado']}{RN}"
    e0 = f"${C['estado']}{R0}"
    ws.conditional_formatting.add(est, FormulaRule(formula=[f'LEFT({e0},5)="Error"'], fill=fill_alerta,
                                                   font=Font(name=F, bold=True, color="A3121F")))
    ws.conditional_formatting.add(est, FormulaRule(formula=[f'{e0}="Bajo meta"'], fill=fill_alerta))
    ws.conditional_formatting.add(est, FormulaRule(formula=[f'OR({e0}="Revisar balance",{e0}="Falta registrar pérdidas")'],
                                                   fill=fill_aviso))
    ws.conditional_formatting.add(est, FormulaRule(formula=[f'{e0}="OK"'], fill=fill_ok))
    rend = f"{C['rend']}{R0}:{C['rend']}{RN}"
    ws.conditional_formatting.add(rend, FormulaRule(
        formula=[f'AND(${C["rend"]}{R0}<>"",${C["rend"]}{R0}<${C["meta"]}{R0})'],
        font=Font(name=F, bold=True, color="C8102E")))
    return ws


# ---------------------------------------------------------------- Parámetros
def hoja_parametros(wb):
    ws = wb.create_sheet("Parámetros")
    titulo(ws, "Parámetros", "Celdas amarillas editables. Todas las hojas leen estos valores.")
    filas = [
        (4, "Año a analizar", 2026, "0", "Año de los resúmenes, la carta semanal y el Pareto."),
        (5, "Meta general de rendimiento", META_GENERAL, PCT,
         "Se usa cuando el grupo no tiene meta propia. Valor por defecto: meta ponderada 2026 del CSV."),
        (6, "Oxidación estimada (% del consumo)", 0.02, PCT,
         "Se aplica cuando no se ingresa la oxidación medida. Archivo 2010: 2 %."),
        (7, "Tolerancia del delta (% del consumo)", 0.01, PCT,
         "Si |Delta| / Consumo supera este valor, la orden queda como 'Revisar balance'. Supuesto: 1 %."),
        (8, "Constante d2 (carta I-MR, n = 2)", 1.128, "0.000", "σ = MR̄ / d2. Constante estadística; no cambiar."),
    ]
    for r, et, v, fmt, nota in filas:
        ws.cell(r, 1, et).font = f_bold
        entrada(ws.cell(r, 2), v, fmt)
        ws.cell(r, 3, nota).font = f_nota
    ws["B8"].font, ws["B8"].fill = f_base, fill_sub

    ws["A11"] = "Listas y metas por grupo"
    ws["A11"].font = f_sub
    encabezado(ws, 13, ["Grupo", "Meta"], 1)
    encabezado(ws, 13, ["Turno"], 4)
    encabezado(ws, 13, ["Causa (categoría)"], 6)
    encabezado(ws, 13, ["Estado de acción"], 8)
    encabezado(ws, 13, ["Tipo de acción"], 10)
    encabezado(ws, 13, ["Tipo de pérdida"], 12)
    for i, (g, m) in enumerate(GRUPOS):
        entrada(ws.cell(14 + i, 1), g)
        entrada(ws.cell(14 + i, 2), m, PCT)
    for col, valores in ((4, TURNOS), (6, CAUSAS), (8, ESTADOS), (10, TIPOS_ACCION),
                         (12, [p for p, _k in PERDIDAS])):
        for i, v in enumerate(valores):
            entrada(ws.cell(14 + i, col), v)
    ws["A23"] = ("Metas por grupo: Rend_Met del CSV de consumo 2026, ponderado por kg consumidos. "
                 "Estrella no tiene órdenes en el CSV: queda en blanco y usa la meta general.")
    ws["A23"].font = f_nota
    ws["A25"] = "Puede agregar valores en las filas vacías de cada lista (Grupo hasta la fila 21, Causa hasta la 23)."
    ws["A25"].font = f_nota
    anchos(ws, [34, 11, 4, 8, 4, 32, 4, 20, 4, 14, 4, 24])
    ws.column_dimensions["C"].width = 4
    # la nota de la columna C de los parámetros es larga: se deja desbordar
    return ws


# ---------------------------------------------------------------- Resumen mensual
def hoja_resumen(wb):
    ws = wb.create_sheet("Resumen mensual")
    titulo(ws, "Resumen mensual", "Rendimiento ponderado = Σ producción / Σ consumo. Año en Parámetros!B4.")
    ws["A3"] = "Año"
    ws["A3"].font = f_bold
    ws["B3"] = f"={PAR}!B4"
    ws["B3"].font = Font(name=F, bold=True, color="008000")
    cab = ["Mes", "N° órdenes", "Consumo (kg)", "Producción (kg)", "Rend. metálico", "Meta ponderada",
           "Brecha", "Kg faltantes vs meta", "Órdenes bajo meta"] + \
          [f"{p} (kg)" for p, _k in PERDIDAS[:-1]] + \
          ["Pérdidas medidas (kg)", "Pérdidas SAP (kg)", "Delta no explicado (kg)", "Delta / consumo",
           "Pérdidas (kg/t)", "Productividad (t/h)"]
    encabezado(ws, 5, cab)
    anchos(ws, [9, 8, 13, 13, 9, 9, 9, 11, 8] + [10] * 8 + [11, 11, 11, 9, 9, 10])
    crit = lambda r: f'{rango("mes")},$A{r}'  # noqa: E731
    for m in range(12):
        r = 6 + m
        ws.cell(r, 1, f"=DATE($B$3,{m + 1},1)").number_format = "mmm-yy"
        ws.cell(r, 2, f'=COUNTIFS({crit(r)},{rango("cons")},">0")')
        ws.cell(r, 3, f'=SUMIFS({rango("cons")},{crit(r)})')
        ws.cell(r, 4, f'=SUMIFS({rango("prod")},{crit(r)})')
        ws.cell(r, 5, f'=IF(C{r}=0,"",D{r}/C{r})')
        ws.cell(r, 6, f'=IF(C{r}=0,"",SUMIFS({rango("cxm")},{crit(r)})/C{r})')
        ws.cell(r, 7, f'=IF(E{r}="","",(E{r}-F{r})*100)')
        ws.cell(r, 8, f'=SUMIFS({rango("falt")},{crit(r)})')
        ws.cell(r, 9, f'=COUNTIFS({crit(r)},{rango("estado")},"Bajo meta")')
        for j, (_p, k) in enumerate(PERDIDAS[:-1]):
            ws.cell(r, 10 + j, f'=SUMIFS({rango(k)},{crit(r)})')
        ws.cell(r, 18, f'=SUM(J{r}:Q{r})')
        ws.cell(r, 19, f'=SUMIFS({rango("sap")},{crit(r)})')
        ws.cell(r, 20, f'=S{r}-R{r}')
        ws.cell(r, 21, f'=IF(C{r}=0,"",T{r}/C{r})')
        ws.cell(r, 22, f'=IF(D{r}=0,"",S{r}/(D{r}/1000))')
        ws.cell(r, 23, f'=IFERROR(D{r}/1000/SUMIFS({rango("horas")},{crit(r)}),"")')
    t = 18
    ws.cell(t, 1, "Año").font = f_bold
    for col in [2, 3, 4, 8, 9] + list(range(10, 21)):
        L = get_column_letter(col)
        ws.cell(t, col, f"=SUM({L}6:{L}17)")
    ws.cell(t, 5, f'=IF(C{t}=0,"",D{t}/C{t})')
    ws.cell(t, 6, f'=IF(C{t}=0,"",SUMIFS({rango("cxm")},{rango("mes")},">="&A6,{rango("mes")},"<="&A17)/C{t})')
    ws.cell(t, 7, f'=IF(E{t}="","",(E{t}-F{t})*100)')
    ws.cell(t, 21, f'=IF(C{t}=0,"",T{t}/C{t})')
    ws.cell(t, 22, f'=IF(D{t}=0,"",S{t}/(D{t}/1000))')
    ws.cell(t, 23, f'=IFERROR(D{t}/1000/SUMIFS({rango("horas")},{rango("mes")},">="&A6,{rango("mes")},"<="&A17),"")')
    fmts = {2: "0", 3: KG, 4: KG, 5: PCT, 6: PCT, 7: PP, 8: KG, 9: "0", 21: PCT, 22: "0.0", 23: "0.00"}
    for r in range(6, t + 1):
        for col in range(2, 24):
            c = ws.cell(r, col)
            c.number_format = fmts.get(col, KG)
            c.font = f_bold if r == t else f_base
            if r == t:
                c.fill = fill_sub
        ws.cell(r, 1).border = borde
    ws.cell(t, 1).fill = fill_sub
    ws.conditional_formatting.add("G6:G18", CellIsRule(operator="lessThan", formula=["0"],
                                                      font=Font(name=F, bold=True, color="C8102E")))
    ws.conditional_formatting.add("U6:U18", FormulaRule(formula=[f'AND(U6<>"",ABS(U6)>{PAR}!$B$7)'],
                                                       fill=fill_aviso))
    ws.freeze_panes = "B6"

    # grupo x mes
    g0 = 22
    ws.cell(g0 - 1, 1, "Rendimiento por grupo y mes").font = f_sub
    encabezado(ws, g0, ["Grupo"] + [None] * 12 + ["Año", "Meta"])
    for m in range(12):
        ws.cell(g0, 2 + m).value = f"=A{6 + m}"
        ws.cell(g0, 2 + m).number_format = "mmm"
    ws.column_dimensions["A"].width = 13
    for i in range(len(GRUPOS)):
        r = g0 + 1 + i
        ws.cell(r, 1, f"={PAR}!A{14 + i}").font = f_bold
        for m in range(12):
            L = get_column_letter(2 + m)
            cr = f'{rango("grupo")},$A{r},{rango("mes")},{L}${g0}'
            ws.cell(r, 2 + m, f'=IFERROR(SUMIFS({rango("prod")},{cr})/SUMIFS({rango("cons")},{cr}),"")')
        cr = f'{rango("grupo")},$A{r},{rango("mes")},">="&$A$6,{rango("mes")},"<="&$A$17'
        ws.cell(r, 14, f'=IFERROR(SUMIFS({rango("prod")},{cr})/SUMIFS({rango("cons")},{cr}),"")')
        ws.cell(r, 15, f'=IF({PAR}!B{14 + i}="",{PAR}!$B$5,{PAR}!B{14 + i})')
        for col in range(2, 16):
            ws.cell(r, col).number_format = "0.0%"
            ws.cell(r, col).font = f_base
    gl = g0 + len(GRUPOS)
    ws.conditional_formatting.add(f"B{g0 + 1}:N{gl}", FormulaRule(
        formula=[f'AND(B{g0 + 1}<>"",B{g0 + 1}<$O{g0 + 1})'], fill=fill_alerta))
    ws.conditional_formatting.add(f"B{g0 + 1}:N{gl}", FormulaRule(
        formula=[f'AND(B{g0 + 1}<>"",B{g0 + 1}>=$O{g0 + 1})'], fill=fill_ok))
    ws.cell(gl + 1, 1, "Rosado = bajo la meta del grupo; verde = cumple.").font = f_nota

    ch = BarChart()
    ch.type = "col"
    ch.title = "Rendimiento mensual vs meta"
    ch.height, ch.width = 7.5, 17
    ch.add_data(Reference(ws, min_col=5, max_col=6, min_row=5, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    ch.y_axis.number_format = "0%"
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0.88, 1.0
    ch.x_axis.number_format = "mmm"
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.series[0].graphicalProperties.solidFill = "5F5E5A"
    ch.series[1].graphicalProperties.solidFill = ROJO
    ch.gapWidth = 60
    ch.legend.position = "b"
    ws.add_chart(ch, f"A{gl + 3}")

    st = BarChart()
    st.type, st.grouping, st.overlap = "col", "stacked", 100
    st.title = "Pérdidas por tipo (kg)"
    st.height, st.width = 7.5, 17
    st.add_data(Reference(ws, min_col=10, max_col=17, min_row=5, max_row=17), titles_from_data=True)
    st.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17))
    st.x_axis.number_format = "mmm"
    st.x_axis.delete = st.y_axis.delete = False
    st.legend.position = "b"
    ws.add_chart(st, f"J{gl + 3}")
    return ws


# ---------------------------------------------------------------- Carta semanal
def hoja_carta(wb):
    ws = wb.create_sheet("Carta semanal")
    titulo(ws, "Carta de control I-MR semanal",
           "LC = rendimiento ponderado del año; σ = MR̄ / d2; LSC / LIC = LC ± 3σ. Semanas ISO (lunes a domingo).")
    params = [("LC (rend. ponderado del año)", "='Resumen mensual'!E18", PCT),
              ("MR̄ (rango móvil promedio)", '=IFERROR(AVERAGE(F12:F64),"")', "0.000%"),
              ("σ", f'=IF(B5="","",B5/{PAR}!B8)', "0.000%"),
              ("LSC = LC + 3σ", '=IF(OR(B4="",B6=""),"",B4+3*B6)', PCT),
              ("LIC = LC − 3σ", '=IF(OR(B4="",B6=""),"",B4-3*B6)', PCT),
              ("Meta ponderada del año", "='Resumen mensual'!F18", PCT)]
    for i, (et, f, fmt) in enumerate(params):
        ws.cell(4 + i, 1, et).font = f_bold
        c = ws.cell(4 + i, 2, f)
        c.number_format, c.font = fmt, Font(name=F, bold=True, color="008000" if "Resumen" in f else "000000")
    encabezado(ws, 11, ["Semana", "Etiqueta", "Consumo (kg)", "Producción (kg)", "Rend. semanal",
                        "Rango móvil (MR)", "LC", "LSC", "LIC", "Meta", "Semanas seguidas bajo LC", "Señal"])
    anchos(ws, [30, 14, 13, 13, 10, 10, 9, 9, 9, 9, 10, 26])
    ws.column_dimensions["A"].width = 30
    for w in range(1, 54):
        r = 11 + w
        cr = f'{rango("anio")},{PAR}!$B$4,{rango("sem")},{w}'
        ws.cell(r, 1, w)
        ws.cell(r, 2, f'="S"&TEXT(A{r},"00")')
        ws.cell(r, 3, f"=SUMIFS({rango('cons')},{cr})")
        ws.cell(r, 4, f"=SUMIFS({rango('prod')},{cr})")
        ws.cell(r, 5, f'=IF(C{r}=0,"",D{r}/C{r})')
        if w > 1:
            ws.cell(r, 6, f'=IF(OR(E{r}="",E{r - 1}=""),"",ABS(E{r}-E{r - 1}))')
        ws.cell(r, 7, "=$B$4")
        ws.cell(r, 8, "=$B$7")
        ws.cell(r, 9, "=$B$8")
        ws.cell(r, 10, "=$B$9")
        ws.cell(r, 11, f'=IF(OR(E{r}="",$B$4=""),"",IF(E{r}<G{r},{"0" if w == 1 else f"N(K{r - 1})"}+1,0))')
        ws.cell(r, 12, f'=IF(E{r}="","",IF(AND(H{r}<>"",OR(E{r}>H{r},E{r}<I{r})),"Fuera de control",'
                       f'IF(K{r}>=8,"Tendencia: 8+ semanas bajo LC",IF(E{r}<J{r},"Bajo meta","OK"))))')
        for col, fmt in ((3, KG), (4, KG), (5, PCT), (6, "0.000%"), (7, PCT), (8, PCT), (9, PCT), (10, PCT), (11, "0")):
            ws.cell(r, col).number_format = fmt
        for col in range(1, 13):
            ws.cell(r, col).font = f_base
            ws.cell(r, col).border = borde
    sen = "L12:L64"
    ws.conditional_formatting.add(sen, FormulaRule(formula=['L12="Fuera de control"'], fill=fill_alerta,
                                                   font=Font(name=F, bold=True, color="A3121F")))
    ws.conditional_formatting.add(sen, FormulaRule(formula=['LEFT(L12,9)="Tendencia"'], fill=fill_aviso))
    ws.conditional_formatting.add(sen, FormulaRule(formula=['L12="Bajo meta"'], fill=fill_alerta))
    ws.conditional_formatting.add(sen, FormulaRule(formula=['L12="OK"'], fill=fill_ok))
    ws.freeze_panes = "C12"

    bar = BarChart()
    bar.type = "col"
    bar.title = "Carta de control semanal"
    bar.height, bar.width = 9, 26
    bar.add_data(Reference(ws, min_col=5, min_row=11, max_row=64), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=2, min_row=12, max_row=64))
    bar.series[0].graphicalProperties.solidFill = "5F5E5A"
    bar.series[0].graphicalProperties.line.noFill = True
    bar.gapWidth = 40
    bar.y_axis.scaling.min, bar.y_axis.scaling.max = 0.88, 1.0
    bar.y_axis.number_format = "0%"
    bar.y_axis.majorGridlines = None
    bar.x_axis.delete = bar.y_axis.delete = False
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=7, max_col=10, min_row=11, max_row=64), titles_from_data=True)
    for s, color, dash in zip(ln.series, (GRAFITO, "C8102E", "C8102E", "2A78D6"), (None, "dash", "dash", "sysDot")):
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 19000
        if dash:
            s.graphicalProperties.line.dashStyle = dash
        s.smooth = False
    bar += ln
    bar.legend.position = "b"
    ws.add_chart(bar, "N4")
    ws.cell(66, 1, "Eje vertical de 88 % a 100 %: las semanas sin datos no muestran barra. MR solo se calcula entre "
                   "semanas consecutivas con datos. Reglas: un punto fuera de LSC/LIC o 8 semanas seguidas bajo LC "
                   "indican una causa especial que hay que investigar.").font = f_nota
    return ws


# ---------------------------------------------------------------- Pareto
def pareto(ws, fila, titulo_txt, nombres, valor_fn, total_fn=None, nota=None, grafico_col="H"):
    """Tabla base (sin ordenar) + tabla ordenada de mayor a menor con % acumulado."""
    n = len(nombres)
    ws.cell(fila, 1, titulo_txt).font = f_sub
    if nota:
        ws.cell(fila + 1, 1, nota).font = f_nota
    h = fila + 2
    encabezado(ws, h, ["#", "Ítem", "kg", "% del total", "% acumulado", "Prioridad"])
    base = h + n + 3  # tabla base, a la derecha no: debajo, compacta
    ws.cell(base - 1, 1, "Base de cálculo (sin ordenar)").font = f_nota
    for i, nom in enumerate(nombres):
        r = base + i
        ws.cell(r, 2, nom).font = f_nota
        c = ws.cell(r, 3, valor_fn(i, r))
        c.number_format, c.font = KG, f_nota
        ws.cell(r, 4, f"=C{r}+ROW()/1000000").font = f_nota  # desempate
        ws.cell(r, 4).number_format = "0.000000"
    rb = f"$D${base}:$D${base + n - 1}"
    tot = f"SUMIF($C${base}:$C${base + n - 1},\">0\")"
    for k in range(1, n + 1):
        r = h + k
        ws.cell(r, 1, k)
        ws.cell(r, 2, f"=INDEX($B${base}:$B${base + n - 1},MATCH(LARGE({rb},{k}),{rb},0))")
        ws.cell(r, 3, f"=INDEX($C${base}:$C${base + n - 1},MATCH(LARGE({rb},{k}),{rb},0))")
        ws.cell(r, 4, f'=IF({tot}=0,"",MAX(0,C{r})/{tot})')
        ws.cell(r, 5, f'=IF(D{r}="","",SUM($D${h + 1}:D{r}))')
        ws.cell(r, 6, f'=IF(OR(E{r}="",C{r}<=0),"",IF(E{r}-D{r}<0.8,"Prioridad (80 %)","Monitorear"))')
        for col, fmt in ((3, KG), (4, "0.0%"), (5, "0.0%")):
            ws.cell(r, col).number_format = fmt
        for col in range(1, 7):
            ws.cell(r, col).font = f_base
            ws.cell(r, col).border = borde
    ws.conditional_formatting.add(f"F{h + 1}:F{h + n}", FormulaRule(
        formula=[f'F{h + 1}="Prioridad (80 %)"'], fill=fill_alerta, font=Font(name=F, bold=True)))
    ch = BarChart()
    ch.type = "col"
    ch.title = titulo_txt
    ch.height, ch.width = 9, 18
    ch.add_data(Reference(ws, min_col=3, min_row=h, max_row=h + n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=h + 1, max_row=h + n))
    ch.series[0].graphicalProperties.solidFill = ROJO
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.legend = None
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=5, min_row=h, max_row=h + n), titles_from_data=True)
    ln.y_axis.axId = 200
    ln.y_axis.number_format = "0%"
    ln.y_axis.scaling.min, ln.y_axis.scaling.max = 0, 1
    ln.y_axis.crosses = "max"
    ln.y_axis.delete = False
    ln.series[0].graphicalProperties.line.solidFill = GRAFITO
    ch += ln
    ws.add_chart(ch, f"{grafico_col}{fila}")
    return base + n + 2


def hoja_pareto(wb):
    ws = wb.create_sheet("Pareto")
    titulo(ws, "Pareto de pérdidas y de causas",
           "Elija el mes en C4 (1 a 12) o deje 0 para el año completo (Parámetros!B4).")
    ws["A4"], ws["A4"].font = "Mes", f_bold
    entrada(ws["C4"], 0, "0")
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="12")
    dv.error = "Ingrese un mes entre 1 y 12, o 0 para el año completo."
    ws.add_data_validation(dv)
    dv.add("C4")
    ws["A5"], ws["A5"].font = "Desde", f_bold
    ws["C5"] = f"=IF($C$4=0,DATE({PAR}!$B$4,1,1),DATE({PAR}!$B$4,$C$4,1))"
    ws["A6"], ws["A6"].font = "Hasta", f_bold
    ws["C6"] = f"=IF($C$4=0,DATE({PAR}!$B$4,12,1),DATE({PAR}!$B$4,$C$4,1))"
    for c in ("C5", "C6"):
        ws[c].number_format = "mmm-yy"
    per = f'{rango("mes")},">="&$C$5,{rango("mes")},"<="&$C$6'
    anchos(ws, [5, 32, 12, 11, 11, 16])

    fila = pareto(ws, 8, "Pérdidas por tipo (kg)", [p for p, _k in PERDIDAS],
                  lambda i, r: f"=SUMIFS({rango(PERDIDAS[i][1])},{per})",
                  nota="Qué etapa del proceso concentra los kg perdidos. Un delta alto = pérdidas sin registrar.")
    nombres = CAUSAS + ["Sin causa asignada"]

    def val_causa(i, r):
        if i < len(CAUSAS):
            return f"=SUMIFS({rango('falt')},{per},{rango('causa')},{PAR}!$F${14 + i})"
        return f"=SUMIFS({rango('falt')},{per})-SUM(C{r - len(CAUSAS)}:C{r - 1})"
    fila = pareto(ws, fila + 2, "Kg faltantes vs meta por causa", nombres, val_causa,
                  nota="Kg faltantes = consumo × (meta − rendimiento) de las órdenes bajo meta, "
                       "asignados a la causa registrada en la hoja Registro.")

    def val_grupo(i, r):
        return f"=SUMIFS({rango('falt')},{per},{rango('grupo')},{PAR}!$A${14 + i})"
    pareto(ws, fila + 2, "Kg faltantes vs meta por grupo de producto", [g for g, _m in GRUPOS], val_grupo)
    return ws


# ---------------------------------------------------------------- Plan de acción
def hoja_acciones(wb):
    ws = wb.create_sheet("Plan de acción")
    titulo(ws, "Plan de acción y verificación de eficacia",
           "Una fila por problema. Cierre la acción solo cuando el rendimiento 'después' confirme la mejora.")
    cols = [("N°", 5, None), ("Fecha detección", 11, FECHA), ("Orden / semana", 12, None),
            ("Grupo / producto", 18, None), ("Hecho (problema observado)", 32, None),
            ("Tipo de pérdida", 16, None), ("Causa (categoría)", 24, None),
            ("Causa raíz (5 porqués)", 34, None), ("Acción", 34, None), ("Tipo de acción", 12, None),
            ("Responsable", 14, None), ("Fecha compromiso", 11, FECHA), ("Estado", 13, None),
            ("Fecha cierre", 11, FECHA), ("Rend. antes", 9, PCT), ("Rend. después", 9, PCT),
            ("Mejora", 9, PP), ("Días de atraso", 8, "0")]
    encabezado(ws, 4, [c[0] for c in cols], calc={17, 18})
    anchos(ws, [c[1] for c in cols])
    for r in range(5, 5 + N_ACC):
        for j, (_h, _w, fmt) in enumerate(cols, 1):
            c = ws.cell(r, j)
            if j == 17:
                c.value = f'=IF(OR(O{r}="",P{r}=""),"",(P{r}-O{r})*100)'
                c.font = f_base
            elif j == 18:
                c.value = f'=IF(OR(L{r}="",LEFT(M{r},7)="Cerrada"),"",MAX(0,TODAY()-L{r}))'
                c.font = f_base
            else:
                c.font, c.fill = f_input, fill_input
            if fmt:
                c.number_format = fmt
    ej = [1, date(2010, 11, 2), "1008027", "Plana 50x06", "Cobles (trozo de coble y barra de sacrificio)",
          "Cobles / barras perdidas", "Eléctrico / instrumentación",
          "HMD A5-A6 pierde señal: flanco de bajada al pasar la barra (visto en IBA)",
          "Revisar y reemplazar HMD A5-A6; verificar señal en IBA durante 1 semana", "Correctiva",
          "Eléctrico", date(2010, 11, 9), "Cerrada", date(2010, 11, 8), 0.9429, 0.9510]
    for j, v in enumerate(ej, 1):
        ws.cell(5, j, v)
    ws["O5"].comment = Comment("Ejemplo basado en la orden 1008027 (noviembre 2010). "
                               "Rend. después es un supuesto para mostrar el formato.", "Plantilla")
    lista(ws, f"={PAR}!$L$14:$L$22", f"F5:F{4 + N_ACC}")
    lista(ws, f"={PAR}!$F$14:$F$23", f"G5:G{4 + N_ACC}")
    lista(ws, f"={PAR}!$J$14:$J$16", f"J5:J{4 + N_ACC}")
    lista(ws, f"={PAR}!$H$14:$H$17", f"M5:M{4 + N_ACC}")
    ult = 4 + N_ACC
    ws.conditional_formatting.add(f"R5:R{ult}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill_alerta,
                                                          font=Font(name=F, bold=True, color="A3121F")))
    ws.conditional_formatting.add(f"M5:M{ult}", FormulaRule(formula=['M5="Cerrada"'], fill=fill_ok))
    ws.conditional_formatting.add(f"M5:M{ult}", FormulaRule(formula=['M5="Cerrada - no eficaz"'], fill=fill_aviso))
    ws.conditional_formatting.add(f"Q5:Q{ult}", CellIsRule(operator="lessThan", formula=["0"],
                                                          font=Font(name=F, bold=True, color="C8102E")))
    ws.freeze_panes = "F5"
    ws.auto_filter.ref = f"A4:R{ult}"
    return ws


# ---------------------------------------------------------------- Inicio y Guía
def hoja_inicio(wb):
    ws = wb.active
    ws.title = "Inicio"
    titulo(ws, "Control de rendimiento metálico — plantilla")
    ws["A3"] = "Indicadores del año (Parámetros!B4)"
    ws["A3"].font = f_sub
    kpis = [("Rendimiento metálico", "='Resumen mensual'!E18", PCT),
            ("Meta ponderada", "='Resumen mensual'!F18", PCT),
            ("Brecha", "='Resumen mensual'!G18", PP),
            ("Kg faltantes vs meta", "='Resumen mensual'!H18", KG),
            ("Pérdidas SAP (kg)", "='Resumen mensual'!S18", KG),
            ("Delta no explicado (kg)", "='Resumen mensual'!T18", KG),
            ("Delta / consumo", "='Resumen mensual'!U18", PCT),
            ("Órdenes registradas", "='Resumen mensual'!B18", "0"),
            ("Órdenes bajo meta", "='Resumen mensual'!I18", "0"),
            ("Órdenes con error o balance a revisar",
             f'=COUNTIF({rango("estado")},"Error*")+COUNTIF({rango("estado")},"Revisar balance")', "0"),
            ("Acciones abiertas con atraso", "=COUNTIF('Plan de acción'!R5:R204,\">0\")", "0")]
    for i, (et, f, fmt) in enumerate(kpis):
        r = 4 + i
        ws.cell(r, 1, et).font = f_base
        c = ws.cell(r, 2, f)
        c.number_format, c.font = fmt, Font(name=F, size=11, bold=True, color="008000")
        ws.cell(r, 1).border = ws.cell(r, 2).border = borde

    r = 17
    ws.cell(r, 1, "Cómo usar").font = f_sub
    pasos = [
        "1. Parámetros: fije el año, la meta general, las metas por grupo, el % de oxidación y la tolerancia del delta.",
        "2. Registro: ingrese una fila por orden al cierre de cada turno o campaña (celdas amarillas). "
        "La fila 6 es un ejemplo: reemplácela.",
        "3. Revise la columna Estado: 'Error', 'Revisar balance' y 'Falta registrar pérdidas' son problemas de dato; "
        "'Bajo meta' es un problema de proceso y debe llevar Hecho, Causa y Acción.",
        "4. Resumen mensual, Carta semanal y Pareto se actualizan solos.",
        "5. Semanalmente: lea la carta, tome los ítems 'Prioridad (80 %)' del Pareto y abra acciones en Plan de acción.",
        "6. Cierre cada acción comparando el rendimiento antes/después del mismo producto (columna Mejora).",
    ]
    for i, p in enumerate(pasos):
        texto_largo(ws, r + 1 + i, p, hasta=6, alto=28)
    r += len(pasos) + 2
    ws.cell(r, 1, "Leyenda").font = f_sub
    leyenda = [(fill_input, f_input, "Dato a ingresar (texto azul, fondo amarillo)"),
               (fill_calc, f_head, "Encabezado de columna calculada: no escribir"),
               (None, Font(name=F, bold=True, color="008000"), "Texto verde: valor traído de otra hoja"),
               (fill_alerta, f_base, "Bajo meta / fuera de control / atrasada"),
               (fill_aviso, f_base, "Dato a revisar (balance, tendencia)"),
               (fill_ok, f_base, "Cumple")]
    for i, (fill, font, txt) in enumerate(leyenda):
        c = ws.cell(r + 1 + i, 1, "Ejemplo")
        c.font = font
        if fill:
            c.fill = fill
        ws.cell(r + 1 + i, 2, txt).font = f_base
    r += len(leyenda) + 2
    ws.cell(r, 1, "Hojas").font = f_sub
    hojas = [("Registro", "Base por orden: producción, consumo, pérdidas por etapa, balance y análisis."),
             ("Resumen mensual", "Rendimiento, meta, kg faltantes y pérdidas por tipo por mes; rendimiento grupo × mes."),
             ("Carta semanal", "Carta I-MR con LC, LSC, LIC, meta y señales de fuera de control o tendencia."),
             ("Pareto", "Pérdidas por tipo, kg faltantes por causa y por grupo, con el 80 % prioritario."),
             ("Plan de acción", "Seguimiento de acciones, atraso y verificación de eficacia."),
             ("Guía de control", "Qué medir, cómo leer los indicadores y rutina de control."),
             ("Parámetros", "Metas, supuestos y listas desplegables.")]
    for i, (h, d) in enumerate(hojas):
        ws.cell(r + 1 + i, 1, h).font = f_bold
        ws.cell(r + 1 + i, 2, d).font = f_base
    anchos(ws, [38, 16, 14, 14, 14, 14])
    return ws


GUIA = [
    ("1. Qué es y cómo se calcula", [
        "Rendimiento metálico = kg producidos (barras buenas) / kg consumidos (palanquilla cargada). "
        "Es el indicador de cuánto acero se convierte en producto vendible.",
        "Se calcula siempre ponderado: Σ producción / Σ consumo. No promedie porcentajes de órdenes: "
        "una orden chica pesaría igual que una grande.",
        "Cada punto de rendimiento vale mucho: en 1.000 t consumidas, 1 pp = 10 t de acero que vuelven como chatarra o se pierden como laminilla.",
    ]),
    ("2. Dónde se pierde el metal (balance de masa)", [
        "Consumo − Producción = Pérdidas SAP. Esa diferencia se reparte en:",
        "• Oxidación / laminilla (horno y descascarillado): típicamente 1–2,5 % del consumo. Depende de temperatura, tiempo "
        "de residencia en el horno y atmósfera. Si se pesa la laminilla, ingrésela en 'Oxidación medida'.",
        "• Despuntes de cizallas T1, T2, T3: cortes de cabeza y cola. Crecen con palanquillas con defectos de extremo, "
        "cabezas frías o despunte mal calibrado.",
        "• T4 / corte a largo comercial: sobrantes de la barra madre respecto del múltiplo del largo comercial. Dependen de "
        "que el largo y el peso de la palanquilla estén optimizados para el producto (padrón).",
        "• Corto T4 y fuera de medida: barras cortas o con dimensión fuera de tolerancia (desgaste de cilindros, guías, ajuste).",
        "• Cobles / barras perdidas: atascos o salidas de línea. Cada coble puede costar cientos de kg.",
        "• Delta no explicado: la parte de las pérdidas SAP que nadie pesó. Si es grande, el problema es de medición/registro, "
        "o de la balanza (consumo o producción mal pesados).",
    ]),
    ("3. Qué controlar y con qué frecuencia", [
        "Por orden / turno: rendimiento vs meta, pérdidas por etapa, n° de cobles, estado del balance (hoja Registro).",
        "Semanal: carta de control I-MR y Pareto; reunión corta para asignar acciones a los ítems del 80 %.",
        "Mensual: tendencia por grupo de producto, kg/t de pérdidas, eficacia de acciones cerradas.",
        "Calidad del dato: órdenes con rendimiento > 100 % o con delta fuera de tolerancia se corrigen antes de analizar.",
    ]),
    ("4. Cómo leer la carta de control", [
        "LC es el nivel actual del proceso; LSC y LIC (± 3σ) es la variación normal. La meta es lo que se quiere lograr.",
        "Punto bajo LIC: algo pasó esa semana (causa especial): buscar en el Registro qué órdenes lo bajaron.",
        "8 semanas seguidas bajo LC: el proceso cambió de nivel (desgaste, cambio de proveedor de palanquilla, práctica).",
        "Si LC está bajo la meta pero no hay señales, el proceso es estable pero incapaz: hay que cambiarlo (proyecto de mejora), "
        "no reaccionar a cada semana.",
        "Recalcule los límites cuando una mejora se consolide (fije el año nuevo o una fecha de corte).",
    ]),
    ("5. Palancas típicas de mejora", [
        "Palanquilla: largo y peso optimizados por producto para minimizar el sobrante en T4 (padrón); controlar defectos "
        "(hojas, rechupes, romboidad) en la recepción; usar el largo estandarizado (el archivo 2010 muestra pérdidas por usar 3,1 m en vez de 3,2 m).",
        "Horno: controlar temperatura y tiempo de residencia para reducir laminilla; evitar esperas largas con material caliente.",
        "Cizallas: calibrar el largo de despunte con fotoceldas / HMD; revisar señales (IBA) cuando el despunte varía.",
        "Laminación: cambio de cilindros y guías según tonelaje, pruebas de primera barra, estandarizar la puesta a punto "
        "para reducir fuera de medida y cobles en cambios de medida y campañas cortas.",
        "Planificación: campañas más largas y secuencia de medidas que reduzca los cambios.",
        "Registro: pesar y anotar todas las pérdidas en el turno; el delta debe tender a cero.",
    ]),
    ("6. Cómo analizar una desviación", [
        "Hecho: qué se observó (dato, no opinión). Ej.: 'Mayor pérdida en T1 y T2, 2 barras perdidas'.",
        "Causa: preguntar ¿por qué? hasta llegar a algo que se pueda corregir (5 porqués). Clasificarla en la categoría de la lista.",
        "Acción: contención (hoy), correctiva (elimina la causa) y preventiva (evita que se repita en otros productos).",
        "Eficacia: comparar el rendimiento del mismo producto antes y después. Si no mejora, marcar 'Cerrada - no eficaz' y volver a analizar.",
    ]),
]


def hoja_guia(wb):
    ws = wb.create_sheet("Guía de control")
    titulo(ws, "Guía para controlar el rendimiento metálico")
    anchos(ws, [110])
    r = 3
    for sec, items in GUIA:
        ws.cell(r, 1, sec).font = f_sub
        r += 1
        for t in items:
            c = ws.cell(r, 1, t)
            c.font = f_base
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 14 * (1 + len(t) // 120)
            r += 1
        r += 1
    return ws


def main():
    wb = Workbook()
    hoja_inicio(wb)
    hoja_registro(wb)
    hoja_resumen(wb)
    hoja_carta(wb)
    hoja_pareto(wb)
    hoja_acciones(wb)
    hoja_guia(wb)
    hoja_parametros(wb)
    for ws in wb.worksheets:
        fuente_arial(ws)
        ws.sheet_view.showGridLines = False
    wb["Registro"].sheet_properties.tabColor = "FFFF00"
    wb["Parámetros"].sheet_properties.tabColor = "FFFF00"
    wb["Plan de acción"].sheet_properties.tabColor = "FFFF00"
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    wb.save(SALIDA)
    print(f"Plantilla generada: {SALIDA}")


if __name__ == "__main__":
    main()
