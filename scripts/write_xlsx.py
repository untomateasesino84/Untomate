"""Genera el entregable Excel de conciliación 60 / 61 / 2x."""
import pickle, sys, re
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.workbook.defined_name import DefinedName

W, OUT = sys.argv[1], sys.argv[2]
M = pickle.load(open(f"{W}/model.pkl", "rb"))
bi, kt, so, lt, ac, sbc_pend, sbc_month = M["bi"], M["kt"], M["so"], M["lt"], M["ac"], M["sbc_pend"], M["sbc_month"]
CAT_NAME = M["CAT_NAME"]
d61 = pd.read_pickle(f"{W}/Detalle_61.pkl")
d2 = pd.read_pickle(f"{W}/Base_Cta2.pkl")

# ------------------------------------------------------------------ recodificación final
HMAP = {"H19": "H01", "H20": "H02", "H04": "H03", "H05": "H04", "H02": "H05", "H03": "H06", "H06": "H07", "H08": "H08",
        "H07": "H09", "H16": "H10", "H14": "H11", "H12": "H12", "H13": "H13", "H10": "H14", "H17": "H15", "H01": "H16",
        "H18": "H17", "H11": "H18", "H09": "H19", "": ""}
AMAP = {"AC-01": "AC-01", "AC-06": "AC-02", "AC-07": "AC-03", "AC-08": "AC-04", "AC-09": "AC-05", "AC-10": "AC-06",
        "AC-14": "AC-07", "AC-04": "AC-08", "AC-02": "AC-09", "AC-03": "AC-10", "AC-12": "AC-11", "AC-11": "AC-12",
        "AC-13": "AC-13", "AC-15": "AC-14"}
bi["Código hallazgo"] = bi["Código hallazgo"].fillna("").map(HMAP)
kt["Código"] = kt["Código"].fillna("").map(HMAP)
so["Código"] = so["Código"].fillna("").map(HMAP)
lt["Código"] = lt["Código"].fillna("").map(HMAP)
ac["Hallazgo"] = ac["Hallazgo"].map(HMAP)
ac["AC"] = ac["AC"].map(AMAP)
ac = ac.sort_values(["AC", "Línea"]).reset_index(drop=True)
ac["Línea"] = np.arange(1, len(ac) + 1)
CODE2AC = ac.groupby("Hallazgo")["AC"].first().to_dict()


def fam_conc(acc):
    if acc.startswith(("6091", "6011", "6121", "20")):
        return "Mercaderías"
    if acc.startswith(("6092", "6021", "6123", "24", "21", "23")):
        return "Materias primas"
    if acc.startswith(("6093", "6032", "6033", "6132", "6133", "25")):
        return "Suministros y repuestos"
    if acc.startswith(("6094", "6041", "6042", "6141", "6142", "26")):
        return "Envases y embalajes"
    return "Otros"


ac["Familia"] = ac["Cuenta"].map(fam_conc)


def ac_ref(code, estado):
    if estado.startswith("Error") and code in CODE2AC:
        return CODE2AC[code]
    if estado.startswith("Conciliado (corregido)"):
        return "Ya corregido en Odoo"
    if estado.startswith("Observación"):
        return "Revisión"
    if estado.startswith("Diferencia temporal") and code in CODE2AC:
        return CODE2AC[code]
    return ""


bi["Asiento corrección"] = [ac_ref(c, e) for c, e in zip(bi["Código hallazgo"], bi["Estado conciliación"])]
kt["Asiento corrección"] = [ac_ref(c, e) for c, e in zip(kt["Código"], kt["Estado"])]

MEAS = {"H01": ["61C"], "H02": ["61C"], "H03": ["61K"], "H04": ["60", "609"], "H05": ["61C"], "H06": ["61C"], "H07": ["609"],
        "H08": ["609", "60"], "H09": ["2x"], "H10": ["61C"], "H14": ["60"], "H15": ["2x"], "H16": ["61C", "61K"],
        "H17": ["61C"], "H18": ["61C"], "H19": ["60", "61C"]}
MEAS.pop("H08")


def meas(s, cl, c, op):
    if c == "H07":
        return -s if (cl == "609" and op.startswith("OP06")) else 0.0
    return s if cl in MEAS.get(c, []) else 0.0


bi["Importe medición"] = [meas(s, cl, c, op) for s, cl, c, op in zip(bi.Saldo, bi.Clase, bi["Código hallazgo"], bi["Tipo de operación"])]
bi["Importe abs"] = bi.Saldo.abs()
bi["Destino 61C"] = np.where(bi.Clase == "61C", -bi.Saldo, 0.0)

# ------------------------------------------------------------------ estilos
NAVY, BLUE, CYAN, GREY, LGREY, WHITE = "051C2C", "2251FF", "00A9F4", "7F7F7F", "F2F2F2", "FFFFFF"
GREEN, AMBER, RED, PURPLE = "E2F0D9", "FFF2CC", "FCE4E4", "EDE7F6"
FONT = "Arial"
f_title = Font(name=FONT, size=18, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=11, color=GREY, italic=True)
f_h = Font(name=FONT, size=9, bold=True, color=WHITE)
f_b = Font(name=FONT, size=9, color="000000")
f_bb = Font(name=FONT, size=9, bold=True, color="000000")
f_sec = Font(name=FONT, size=12, bold=True, color=NAVY)
f_link = Font(name=FONT, size=10, color=BLUE, underline="single")
fill_h = PatternFill("solid", fgColor=NAVY)
fill_h2 = PatternFill("solid", fgColor=BLUE)
fill_tot = PatternFill("solid", fgColor=LGREY)
fill_in = PatternFill("solid", fgColor="EAF0FF")
thin = Side(style="thin", color="BFBFBF")
b_bot = Border(bottom=thin)
b_top = Border(top=Side(style="thin", color=NAVY), bottom=Side(style="double", color=NAVY))
NUM = '#,##0;(#,##0);"–"'
NUM2 = '#,##0.00;(#,##0.00);"–"'
PCT = '0.0%;(0.0%);"–"'
STATUS_FILL = {"Conciliado": GREEN, "Conciliado (corregido)": "D9EAD3", "Diferencia temporal (corte)": AMBER,
               "Error – requiere asiento": RED, "Observación": PURPLE}

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, sub, tab=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = title; ws["A1"].font = f_title
    ws["A2"] = sub; ws["A2"].font = f_sub
    ws["A3"] = "◂ Volver al índice"; ws["A3"].hyperlink = "#'Inicio'!A1"; ws["A3"].font = Font(name=FONT, size=8, color=BLUE, underline="single")
    if tab:
        ws.sheet_properties.tabColor = tab
    return ws


def header(ws, r, c, labels, fill=fill_h, h=30):
    for i, t in enumerate(labels):
        x = ws.cell(r, c + i, t); x.font = f_h; x.fill = fill
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = h


def cellw(ws, r, c, v, fmt=None, bold=False, fill=None, wrap=False, align=None, font=None):
    x = ws.cell(r, c, v)
    x.font = font or (f_bb if bold else f_b)
    if fmt: x.number_format = fmt
    if fill: x.fill = PatternFill("solid", fgColor=fill) if isinstance(fill, str) else fill
    if wrap or align: x.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    return x


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# ------------------------------------------------------------------ Base_Integrada (primero para conocer rangos)
BI_COLS = [("ID línea", "ID línea", 7), ("Fuente", "Fuente", 6), ("ID asiento", "ID asiento", 9), ("Asiento", "Asiento", 18),
           ("Fecha", "Fecha", 10), ("Mes", "Mes", 5), ("Diario", "Diario", 18), ("Tipo documento", "Tipo documento", 12),
           ("Cuenta", "Cuenta", 10), ("Clase", "Clase", 6), ("Familia", "Familia", 14), ("Familia conciliación", "Familia conc.", 18), ("Subcuenta 61", "Subcuenta", 9),
           ("Debe", "Debe", 12), ("Haber", "Haber", 12), ("Saldo (D-H)", "Saldo", 13), ("Destino 61C (+)", "Destino 61C", 12),
           ("Importe abs", "Importe abs", 12), ("Producto", "Producto", 40), ("Cantidad", "Cantidad", 10), ("Socio", "Socio", 28),
           ("OC", "OC principal", 9), ("OC servicio / LC", "OC secundaria", 9), ("ID Clave OC|Prod", "ID Clave", 8),
           ("Motivo Odoo", "Motivo", 30), ("Tipo de operación", "Tipo de operación", 34), ("Categoría de conciliación", "Categoría de conciliación", 40),
           ("Estado conciliación", "Estado conciliación", 22), ("Código hallazgo", "Código hallazgo", 8), ("Explicación", "Explicación", 70),
           ("Asiento corrección", "Asiento corrección", 12), ("Importe medición", "Importe medición", 12),
           ("Asiento: 60", "Asiento 60", 11), ("Asiento: 609", "Asiento 609", 11), ("Asiento: 61C", "Asiento 61C", 11),
           ("Asiento: 61K", "Asiento 61K", 11), ("Asiento: 2x", "Asiento 2x", 11), ("Asiento: otras ctas", "Asiento otras ctas", 11),
           ("Glosa", "Glosa", 50), ("Ref. mov. stock", "Ref. movimiento stock", 16), ("Creado por", "Creado por", 20)]
BIC = {lab: L(i + 1) for i, (lab, _, _) in enumerate(BI_COLS)}
N = len(bi)
R0, R1 = 6, 6 + N  # cabecera en fila 6
rng = lambda lab: f"Base_Integrada!${BIC[lab]}${R0 + 1}:${BIC[lab]}${R1}"

wsb = wb.create_sheet("Base_Integrada")
wsb.sheet_view.showGridLines = False
wsb["A1"] = "Base integrada de compras – líneas 60 / 61 / 2x clasificadas"; wsb["A1"].font = f_title
wsb["A2"] = (f"{N:,} líneas Odoo (Ene–Set 2026). Cada línea indica tipo de operación, categoría de conciliación, estado, hallazgo y asiento de corrección. "
             "Filtre por 'Código hallazgo', 'ID Clave' o 'Asiento' para rastrear cualquier cifra del informe.").replace(",", ",")
wsb["A2"].font = f_sub
wsb["A3"] = "◂ Volver al índice"; wsb["A3"].hyperlink = "#'Inicio'!A1"; wsb["A3"].font = Font(name=FONT, size=8, color=BLUE, underline="single")
wsb.sheet_properties.tabColor = GREY
header(wsb, R0, 1, [c[0] for c in BI_COLS])
num_cols = {"Debe", "Haber", "Saldo (D-H)", "Destino 61C (+)", "Importe abs", "Importe medición", "Asiento: 60", "Asiento: 609",
            "Asiento: 61C", "Asiento: 61K", "Asiento: 2x", "Asiento: otras ctas"}
for j, (lab, src, w) in enumerate(BI_COLS, 1):
    wsb.column_dimensions[L(j)].width = w
vals = bi[[c[1] for c in BI_COLS]].copy()
vals["Fecha"] = vals["Fecha"].dt.to_pydatetime()
vals = vals.astype(object).where(pd.notna(vals), None)
for i, row in enumerate(vals.itertuples(index=False), R0 + 1):
    for j, v in enumerate(row, 1):
        c = wsb.cell(i, j, v)
for j, (lab, _, _) in enumerate(BI_COLS, 1):
    fmt = NUM2 if lab in num_cols else ("dd/mm/yyyy" if lab == "Fecha" else ("0" if lab in ("ID línea", "ID asiento", "Mes", "ID Clave OC|Prod") else None))
    if fmt:
        for i in range(R0 + 1, R1 + 1):
            wsb.cell(i, j).number_format = fmt
tab = Table(displayName="tBase", ref=f"A{R0}:{L(len(BI_COLS))}{R1}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
wsb.add_table(tab)
wsb.freeze_panes = wsb.cell(R0 + 1, 5)
ec = BIC["Estado conciliación"]
for st, col in STATUS_FILL.items():
    wsb.conditional_formatting.add(f"{ec}{R0 + 1}:{ec}{R1}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=col)))
# totales de control
wsb["M4"] = "Control:"; wsb["M4"].font = f_bb
wsb["O4"] = f"=SUBTOTAL(9,{BIC['Saldo (D-H)']}{R0 + 1}:{BIC['Saldo (D-H)']}{R1})"; wsb["O4"].number_format = NUM2; wsb["O4"].font = f_bb
wsb["N4"] = "Σ saldo visible →"; wsb["N4"].font = f_b

# orden de revisión: primero lo que requiere acción
_ord = {"Error – requiere asiento": 0, "Diferencia temporal (corte)": 1, "Observación": 2, "Conciliado (corregido)": 3, "Conciliado": 4}
kt = kt.assign(_o=kt.Estado.map(_ord), _a=-kt.D1.abs()).sort_values(["_o", "_a"]).drop(columns=["_o", "_a"]).reset_index(drop=True)
# ------------------------------------------------------------------ Conc_OC_Producto
wsk = sheet("Conc_OC_Producto", "Conciliación factura (60) ↔ recepción (61) ↔ existencias (2x) por OC y producto",
            "Cada fila es una clave OC|Producto. Los importes se calculan con SUMIFS sobre Base_Integrada (columna 'ID Clave'). Tolerancia ±S/ 1. "
            "(*) Las diferencias consideran la reclasificación AC-11 de devoluciones/ajustes de almacén imputados a 60 (col. X).", tab=BLUE)
KH = ["ID Clave", "OC", "Producto", "Proveedor", "Familia (60)", "Familia (61)", "Mes factura", "Mes recepción",
      "Importe 60 (factura/NC)", "Cant. 60", "Destino 61-Compra", "Cant. 61", "Ingreso 2x", "Cant. 2x",
      "Dif 60 vs 61 (*)", "Dif 61 vs 2x (*)", "P.U. factura", "P.U. recepción", "Validación", "Estado", "Código", "Explicación", "Asiento corrección",
      "Devol./ajuste almacén imputado a 60 (AC-11)"]
KR0 = 6
header(wsk, KR0, 1, KH)
widths(wsk, {"A": 7, "B": 9, "C": 42, "D": 28, "E": 13, "F": 13, "G": 8, "H": 8, "I": 13, "J": 10, "K": 13, "L": 10, "M": 13, "N": 10,
             "O": 12, "P": 12, "Q": 10, "R": 10, "S": 11, "T": 22, "U": 7, "V": 70, "W": 13, "X": 16})
idc, cls, sal, qty, top = rng("ID Clave OC|Prod"), rng("Clase"), rng("Saldo (D-H)"), rng("Cantidad"), rng("Tipo de operación")
for i, r in enumerate(kt.itertuples(index=False), KR0 + 1):
    cellw(wsk, i, 1, int(r[kt.columns.get_loc("ID Clave")]), "0")
    cellw(wsk, i, 2, r.OC); cellw(wsk, i, 3, r.Producto); cellw(wsk, i, 4, r.Proveedor)
    cellw(wsk, i, 5, r.FamF if isinstance(r.FamF, str) else ""); cellw(wsk, i, 6, r.FamR if isinstance(r.FamR, str) else "")
    mf = "" if pd.isna(r.MesFmin) else (f"{int(r.MesFmin)}" if r.MesFmin == r.MesFmax else f"{int(r.MesFmin)}-{int(r.MesFmax)}")
    mr = "" if pd.isna(r.MesRmin) else (f"{int(r.MesRmin)}" if r.MesRmin == r.MesRmax else f"{int(r.MesRmin)}-{int(r.MesRmax)}")
    cellw(wsk, i, 7, mf, align="center"); cellw(wsk, i, 8, mr, align="center")
    cellw(wsk, i, 9, f'=SUMIFS({sal},{idc},$A{i},{cls},"60",{top},"OP01*")+SUMIFS({sal},{idc},$A{i},{cls},"60",{top},"OP02*")', NUM2)
    cellw(wsk, i, 10, f'=SUMIFS({qty},{idc},$A{i},{cls},"60",{top},"OP01*")+SUMIFS({qty},{idc},$A{i},{cls},"60",{top},"OP02*")', NUM2)
    cellw(wsk, i, 24, f'=SUMIFS({sal},{idc},$A{i},{cls},"60",{top},"OP05*")+SUMIFS({sal},{idc},$A{i},{cls},"60",{top},"OP16*")', NUM2)
    cellw(wsk, i, 11, f'=-SUMIFS({sal},{idc},$A{i},{cls},"61C")', NUM2)
    cellw(wsk, i, 12, f'=SUMIFS({qty},{idc},$A{i},{cls},"61C")', NUM2)
    cellw(wsk, i, 13, f'=SUMIFS({sal},{idc},$A{i},{cls},"2x")', NUM2)
    cellw(wsk, i, 14, f'=SUMIFS({qty},{idc},$A{i},{cls},"2x")', NUM2)
    cellw(wsk, i, 15, f"=I{i}-K{i}+X{i}", NUM2)
    cellw(wsk, i, 16, f"=K{i}-X{i}-M{i}", NUM2)
    cellw(wsk, i, 17, f'=IFERROR(I{i}/J{i},"")', NUM2)
    cellw(wsk, i, 18, f'=IFERROR(K{i}/L{i},"")', NUM2)
    cellw(wsk, i, 19, f'=IF(AND(ABS(O{i})<=1,ABS(P{i})<=1),"Cuadra","Diferencia")', align="center")
    cellw(wsk, i, 20, r.Estado); cellw(wsk, i, 21, r.Código, align="center"); cellw(wsk, i, 22, r.Explicación)
    cellw(wsk, i, 23, r[kt.columns.get_loc("Asiento corrección")], align="center")
KR1 = KR0 + len(kt)
tab = Table(displayName="tClaves", ref=f"A{KR0}:X{KR1}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
wsk.add_table(tab)
wsk.freeze_panes = "D7"
for st, col in STATUS_FILL.items():
    wsk.conditional_formatting.add(f"T{KR0 + 1}:T{KR1}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=col)))
wsk.conditional_formatting.add(f"S{KR0 + 1}:S{KR1}", CellIsRule(operator="equal", formula=['"Diferencia"'], font=Font(color="C00000", bold=True)))
wsk["H4"] = "Totales:"; wsk["H4"].font = f_bb
for c in "IKMOPX":
    wsk[f"{c}4"] = f"=SUBTOTAL(9,{c}{KR0 + 1}:{c}{KR1})"; wsk[f"{c}4"].number_format = NUM; wsk[f"{c}4"].font = f_bb
KT_RNG = lambda col: f"Conc_OC_Producto!${col}${KR0 + 1}:${col}${KR1}"

# ------------------------------------------------------------------ Asientos_Correccion
wsa = sheet("Asientos_Correccion", "Asientos de corrección propuestos (al 30/09/2026)",
            "Detalle Debe/Haber por asiento, con referencia al asiento/OC de origen. El resumen se calcula con SUMIFS sobre el detalle; la columna 'Cuadre' valida Debe = Haber.", tab="C00000")
AC_DESC = {
    "AC-01": ("H01 / H02", "Revertir el saldo neto de las provisiones de cuadre de 'compras no ingresadas' (set-26 y jul-26 no extornada)."),
    "AC-02": ("H03", "Reclasificar el consumo de componentes enviados a subcontratación (jun–set y residuo ene–may) de 61-Consumo a 61-Compra / 71."),
    "AC-03": ("H05", "Desmontajes: trasladar el efecto neto registrado en 61-Compra a 61-Consumo."),
    "AC-04": ("H06", "Producción propia ingresada con abono a 61-Compra: reclasificar a 71 (producción almacenada)."),
    "AC-05": ("H07", "Costo en destino abonado a 609: trasladar el abono a 61-Compra."),
    "AC-06": ("H08", "Gastos vinculados (609) no capitalizados: destino a existencias (Debe 2x / Haber 61)."),
    "AC-07": ("H10", "Costo en destino con 61 de familia distinta a la existencia: reclasificar entre subcuentas 61."),
    "AC-08": ("H11", "Diferencias de precio/TC factura vs recepción: ajustar existencias y destino."),
    "AC-09": ("H12", "Corte: bienes facturados no recibidos → 28 Existencias por recibir."),
    "AC-10": ("H13", "Corte: bienes recibidos no facturados → 60 / 42 Facturas por recibir."),
    "AC-11": ("H14", "Devoluciones y revaloraciones de almacén imputadas a 60: trasladar a 61."),
    "AC-12": ("H15", "Compras de repuestos cargadas directo a 2x: registrar la compra en 60 y su destino en 61."),
    "AC-13": ("H16", "Gasto/activo menor imputado a 61-Consumo: reclasificar al gasto."),
    "AC-14": ("H17", "Ajustes de inventario/revaloraciones registrados en 61-Compra: trasladar a 61-Consumo."),
}
acs = sorted(ac.AC.unique())
header(wsa, 6, 1, ["Asiento", "Hallazgo", "Descripción", "N° líneas", "Total Debe", "Total Haber", "Cuadre", "Efecto en 60+609", "Efecto en destino 61-Compra", "Efecto en 2x"])
widths(wsa, {"A": 9, "B": 10, "C": 70, "D": 9, "E": 14, "F": 14, "G": 10, "H": 14, "I": 14, "J": 14, "K": 32, "L": 34, "M": 14, "N": 14, "O": 60, "P": 46, "Q": 16, "R": 8, "S": 8, "T": 50})
AD0 = 6 + len(acs) + 4  # cabecera detalle
AD1 = AD0 + len(ac)
arng = lambda c: f"${c}${AD0 + 1}:${c}${AD1}"
for i, a in enumerate(acs, 7):
    h, d = AC_DESC.get(a, ("", ""))
    cellw(wsa, i, 1, a, bold=True); cellw(wsa, i, 2, h); cellw(wsa, i, 3, d, wrap=True)
    cellw(wsa, i, 4, f'=COUNTIFS({arng("A")},A{i})', "0")
    cellw(wsa, i, 5, f'=SUMIFS({arng("M")},{arng("A")},A{i})', NUM2)
    cellw(wsa, i, 6, f'=SUMIFS({arng("N")},{arng("A")},A{i})', NUM2)
    cellw(wsa, i, 7, f'=IF(ROUND(E{i}-F{i},2)=0,"✔ Cuadra","✖ Revisar")', align="center")
    cellw(wsa, i, 8, f'=SUMIFS({arng("M")},{arng("A")},A{i},{arng("R")},"60")-SUMIFS({arng("N")},{arng("A")},A{i},{arng("R")},"60")'
                     f'+SUMIFS({arng("M")},{arng("A")},A{i},{arng("R")},"609")-SUMIFS({arng("N")},{arng("A")},A{i},{arng("R")},"609")', NUM2)
    cellw(wsa, i, 9, f'=SUMIFS({arng("N")},{arng("A")},A{i},{arng("R")},"61C")-SUMIFS({arng("M")},{arng("A")},A{i},{arng("R")},"61C")', NUM2)
    cellw(wsa, i, 10, f'=SUMIFS({arng("M")},{arng("A")},A{i},{arng("R")},"2x")-SUMIFS({arng("N")},{arng("A")},A{i},{arng("R")},"2x")', NUM2)
    wsa.row_dimensions[i].height = 26
it = 7 + len(acs)
cellw(wsa, it, 1, "Total", bold=True, fill=LGREY)
for c in "BC": cellw(wsa, it, " ABC".index(c), None, fill=LGREY)
for c in "DEFHIJ":
    cellw(wsa, it, " ABCDEFGHIJ".index(c), f"=SUM({c}7:{c}{it - 1})", NUM2 if c != "D" else "0", bold=True, fill=LGREY)
cellw(wsa, it, 7, f'=IF(ROUND(E{it}-F{it},2)=0,"✔ Cuadra","✖ Revisar")', bold=True, fill=LGREY, align="center")
wsa.conditional_formatting.add(f"G7:G{it}", CellIsRule(operator="equal", formula=['"✔ Cuadra"'], font=Font(color="00703C", bold=True)))
wsa.conditional_formatting.add(f"G7:G{it}", CellIsRule(operator="equal", formula=['"✖ Revisar"'], font=Font(color="C00000", bold=True)))
wsa.cell(AD0 - 1, 1, "Detalle de líneas").font = f_sec
ADH = ["Asiento", "Línea", "Fecha", "Hallazgo", "Categoría", "Cuenta", "Nombre de la cuenta", "", "", "", "Familia", "Referencia origen (asiento / OC / producto)",
       "Debe", "Haber", "Glosa", "Nota / condición", "", "Clase", "", ""]
ADH = ["Asiento", "Línea", "Fecha", "Hallazgo", "Categoría", "Cuenta", "Nombre de la cuenta", "Familia", "Referencia origen", "Condición / nota",
       "Clase cta", "Glosa", "Debe", "Haber"]
# columnas: A Asiento, B Línea, C Fecha, D Hallazgo, E Cat, F Cuenta, G Nombre, H Familia, I Ref, J Nota, K (vacía), L Glosa, M Debe, N Haber, ... R Clase
cols_detail = {"A": "Asiento", "B": "Línea", "C": "Glosa", "D": "Hallazgo", "E": "Categoría", "F": "Cuenta", "G": "Nombre de la cuenta",
               "H": "Familia", "I": "Referencia origen", "J": "Condición / nota", "L": "Fecha", "M": "Debe", "N": "Haber", "R": "Clase cta"}
for c, t in cols_detail.items():
    x = wsa[f"{c}{AD0}"]; x.value = t; x.font = f_h; x.fill = fill_h; x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for c in "KOPQ":
    wsa[f"{c}{AD0}"].fill = fill_h
for i, r in enumerate(ac.itertuples(index=False), AD0 + 1):
    cellw(wsa, i, 1, r.AC, bold=True); cellw(wsa, i, 2, int(r.Línea), "0"); cellw(wsa, i, 3, r.Glosa)
    cellw(wsa, i, 4, r.Hallazgo, align="center"); cellw(wsa, i, 5, r.Cat, align="center"); cellw(wsa, i, 6, r.Cuenta)
    cellw(wsa, i, 7, r.Nombre); cellw(wsa, i, 8, r.Familia); cellw(wsa, i, 9, r.Referencia); cellw(wsa, i, 10, r.Nota)
    cellw(wsa, i, 12, r.Fecha.to_pydatetime(), "dd/mm/yyyy"); cellw(wsa, i, 13, r.Debe if r.Debe else None, NUM2); cellw(wsa, i, 14, r.Haber if r.Haber else None, NUM2)
    cellw(wsa, i, 18, r.Clase, align="center")
widths(wsa, {"G": 40, "I": 46, "J": 50, "K": 2, "L": 11, "O": 2, "P": 2, "Q": 2})
wsa.freeze_panes = wsa[f"A{AD0 + 1}"] if False else None
wsa.auto_filter.ref = f"A{AD0}:R{AD1}"

# ------------------------------------------------------------------ Hallazgos
wsh = sheet("Hallazgos", "Catálogo de hallazgos: qué no concilia, por qué y cómo se corrige",
            "Cada hallazgo se cuantifica con fórmulas sobre Base_Integrada / Conc_OC_Producto / Asientos_Correccion. Filtre la base por 'Código hallazgo' para ver las líneas.", tab="C00000")
HALL = [
    ("H01", "Provisión 'compras no ingresadas' usada como asiento de cuadre", "Error de práctica contable", "Alta", "Contabilidad",
     "Cada mes se registra Debe/Haber 61-Compra contra 20911101/24/25/26 por la diferencia (60 − 61) de cada familia. Iguala 60 = 61 mensualmente, pero mezcla en un solo importe todas las causas estructurales (SBC, desmontajes, 609, precios) y altera el saldo de existencias (2x) sin movimiento físico.",
     "61x1102 / 20911101, 24111101, 25241101, 26111101, 26211101", "Eliminar el cuadre. Registrar sólo el corte real (AC-09/AC-10) y corregir las causas raíz (AC-02 a AC-14)."),
    ("H02", "Provisión de julio (CG/2026/07/0207) no extornada", "Error de registro", "Alta", "Contabilidad",
     "La mecánica provisión/extorno se rompe en agosto: CG/2026/08/0229 no revierte la de julio, por lo que su efecto (S/ 19.8k) permanece en 61 y 2x.",
     "61x1102 / 2x", "Incluida en AC-01."),
    ("H03", "Subcontratación: consumo de componentes no reclasificado", "Error de registro", "Alta", "Contabilidad / Producción",
     "En Odoo el ingreso del producto subcontratado (1101/SBC) abona 61-Compra por componentes + servicio, mientras los componentes enviados se cargan a 61-Consumo. La reclasificación manual CG/2026/07/0097 sólo cubrió ene–may.",
     "61x1101 → 61x1102 / 71311102", "AC-02. Automatizar: cuenta de producción de la ubicación de subcontratación = 61-Compra, o reclasificación mensual estándar."),
    ("H04", "Subcontratación: servicio facturado ≠ servicio valorizado / sin OC", "Observación", "Media", "Compras / Producción",
     "El servicio de armado/selección (LIDAN, LOGSELAR) se factura en 609/604, a veces sin OC; el valor incorporado al producto (precio OC) difiere de lo facturado, y hay mermas/rendimientos en la selección de castaña.",
     "609x / 604 / 61-Compra", "Revisar por OC en Conc_Subcontratacion. Facturar siempre contra la OC de subcontratación."),
    ("H05", "Desmontajes (Unbuild) registrados en 61-Compra", "Error de registro", "Media", "Contabilidad / Producción",
     "Los componentes devueltos por desmontaje se abonaron a 61-Compra y la reclasificación CG/2026/07/0143 llevó también el consumo a 61-Compra. Un desmontaje es transformación interna, no compra.",
     "61x1102 ↔ 61x1101", "AC-03."),
    ("H06", "Producción propia ingresada con abono a 61-Compra", "Error de configuración", "Media", "Sistemas / Contabilidad",
     "Órdenes de fabricación de productos con categoría de MP/envase abonan la cuenta de entrada de la categoría (61-Compra) en vez de producción almacenada (71).",
     "61x1102 → 71311102", "AC-04. Configurar cuenta de producción de la ubicación virtual 'Production'."),
    ("H07", "Costo en destino abonado a la 609", "Error de configuración", "Media", "Sistemas / Contabilidad",
     "En 18 OC de servicio el LC usó como cuenta de abono la 609 del producto servicio: la 609 queda anulada y el destino no pasa por 61.",
     "609x → 61x1102", "AC-05. El producto de servicio del LC debe tener como cuenta de gasto la 609 y el LC debe abonar 61 (o cuenta puente)."),
    ("H08", "Gastos vinculados (609) no capitalizados", "Error de registro", "Media", "Contabilidad / Comercio exterior",
     "Facturas 609 (fletes, almacenaje, agenciamiento) sin costo en destino asociado: el costo vinculado queda en 60 sin destino en 61 ni mayor valor del inventario.",
     "609x / 2x / 61x1102", "AC-06 (capitalizar) o reclasificar a 63 si no corresponde a compras. Preferible: crear el LC en Odoo."),
    ("H09", "Costos en destino sin factura vinculada (cacao: fletes y costos indirectos)", "Observación", "Alta", "Contabilidad / Costos",
     "LC de cacao (Quillabamba) por S/ 657k cargados a 24 con contrapartida fuera de 60/61 (no visible en la data). El ingreso a 2x no tiene compra en 60 ni destino en 61.",
     "24111101 / contrapartida a identificar", "Identificar la cuenta abonada; si es flete/servicio de compra, registrar en 609 y abonar el LC a 61."),
    ("H10", "Costo en destino con 61 de familia distinta a la existencia", "Error de configuración", "Baja", "Contabilidad",
     "LC aplicado a productos de otra familia (p.ej. 61421102 embalajes contra 26111101 envases).",
     "61x1102 ↔ 61y1102", "AC-07."),
    ("H11", "Diferencia de precio/TC entre factura y recepción", "Error de registro", "Media", "Compras / Contabilidad",
     "Misma cantidad facturada y recibida con distinto valor (precio OC ≠ precio factura, TC distinto en compras USD). Odoo no trasladó la diferencia a la valorización.",
     "2x / 61x1102", "AC-08. Activar/validar 'Diferencias de precio' de la categoría y actualizar el precio de la OC antes de recibir."),
    ("H12", "Bienes facturados no recibidos al 30/09", "Diferencia temporal", "Media", "Almacén",
     "Facturas con OC cuya recepción no está registrada en Odoo (o recepción parcial).",
     "28x / 61x1102", "AC-09 (extornable). Validar físicamente y registrar la recepción."),
    ("H13", "Bienes recibidos no facturados al 30/09", "Diferencia temporal", "Media", "Compras / Cuentas por pagar",
     "Recepciones valorizadas sin factura registrada (o factura parcial).",
     "60x / 42 Facturas por recibir", "AC-10 (extornable)."),
    ("H14", "Devoluciones y revaloraciones de almacén imputadas a 60", "Error de configuración", "Baja", "Sistemas",
     "En devoluciones a proveedor y cambios manuales de costo, parte del ajuste de valor se contabilizó en la 60 (cuenta de la categoría), no en 61.",
     "60x → 61x", "AC-11."),
    ("H15", "Compras de repuestos cargadas directamente a 2x", "Error de configuración", "Media", "Sistemas / Contabilidad",
     "Facturas de repuestos (SACMI, BUHLER, HACOS) contabilizadas Debe 25311102 sin pasar por 60 ni 61: la compra no figura en la 60.",
     "25311102 / 60 / 61331102", "AC-12. Usar recepción de almacén y cuenta de compra 60 en la categoría de repuestos."),
    ("H16", "Gastos/servicios imputados a cuenta 61", "Error de registro", "Baja", "Contabilidad",
     "Alquileres, comisiones bancarias, movilidad y activos menores imputados a 61. Ya reclasificados en julio (CG/2026/07/0101-0102), salvo una balanza en 61321101.",
     "61 → 63/65", "AC-13 (pendiente menor)."),
    ("H17", "Ajustes de inventario/revaloraciones en 61-Compra", "Observación", "Baja", "Contabilidad",
     "Ajustes de cantidad y cambios de costo registrados en la subcuenta de compra de la 61.",
     "61x1102 → 61x1101", "AC-14."),
    ("H18", "Recepciones de merchandising abonadas a 71 (ya corregido)", "Corregido", "Baja", "Sistemas",
     "Productos de merchandising (gorros, tazas, tote bags) se recibieron contra 71; regularizado con CG/2026/07/0139.",
     "71111101 → 61211101", "Ninguna (ya corregido). Corregir la categoría para evitar recurrencia."),
    ("H19", "Compras, NC y devoluciones sin OC", "Observación", "Baja", "Compras",
     "Facturas, NC y devoluciones sin orden de compra vinculada: no se pueden conciliar por clave OC|Producto.",
     "60x / 61x1102", "Exigir OC en compras de bienes inventariables."),
]
HH = ["Código", "Hallazgo", "Tipo", "Prioridad", "Responsable", "Descripción y causa raíz (Odoo)", "Cuentas", "Líneas (base)", "Importe medido (S/)",
      "Importe bruto asiento (Debe)", "Asiento propuesto", "Acción correctiva / preventiva"]
header(wsh, 6, 1, HH)
widths(wsh, {"A": 7, "B": 38, "C": 17, "D": 9, "E": 18, "F": 70, "G": 26, "H": 10, "I": 14, "J": 14, "K": 11, "L": 55})
cod_r, med_r = rng("Código hallazgo"), rng("Importe medición")
for i, h in enumerate(HALL, 7):
    code = h[0]
    for j, v in enumerate([code, h[1], h[2], h[3], h[4], h[5], h[6]], 1):
        cellw(wsh, i, j, v, wrap=True, bold=(j == 1))
    cellw(wsh, i, 8, f'=COUNTIFS({cod_r},A{i})', "#,##0")
    if code in ("H11", "H12", "H13"):
        cellw(wsh, i, 9, f'=SUMIFS({KT_RNG("O")},{KT_RNG("U")},A{i})', NUM)
    elif code == "H08":
        cellw(wsh, i, 9, f'=SUMIFS(Conc_609_LC!$K$7:$K${6 + len(lt)},Conc_609_LC!$M$7:$M${6 + len(lt)},A{i})', NUM)
    else:
        cellw(wsh, i, 9, f'=SUMIFS({med_r},{cod_r},A{i})', NUM)
    cellw(wsh, i, 10, f'=SUMIFS(Asientos_Correccion!{arng("M")},Asientos_Correccion!{arng("D")},A{i})', NUM)
    cellw(wsh, i, 11, ", ".join(sorted(ac.loc[ac.Hallazgo == code, "AC"].unique())) or ("—" if h[2] != "Observación" else "Revisión"), align="center")
    cellw(wsh, i, 12, h[7], wrap=True)
    fillc = {"Error de práctica contable": RED, "Error de registro": RED, "Error de configuración": RED, "Observación": PURPLE,
             "Diferencia temporal": AMBER, "Corregido": GREEN}[h[2]]
    wsh.cell(i, 3).fill = PatternFill("solid", fgColor=fillc)
    wsh.row_dimensions[i].height = 64
HL = 6 + len(HALL)
wsh.freeze_panes = "C7"
wsh.conditional_formatting.add(f"D7:D{HL}", CellIsRule(operator="equal", formula=['"Alta"'], font=Font(color="C00000", bold=True)))

# ------------------------------------------------------------------ Triangulo
wst = sheet("Triangulo", "Triángulo de conciliación: compras (60) ↔ destino (61-Compra) ↔ existencias (2x)",
            "Importes en S/ por categoría de operación. Antes = registros Odoo; Ajustes = asientos propuestos; Después = antes + ajustes. "
            "Destino 61-Compra se muestra en positivo (abono). Todas las cifras son SUMIFS sobre Base_Integrada y Asientos_Correccion.", tab=NAVY)
cats = list(CAT_NAME.keys())
cat_r, cls_r, sal_r = rng("Categoría de conciliación"), rng("Clase"), rng("Saldo (D-H)")
ac_cat, ac_cls, ac_d, ac_h = f"Asientos_Correccion!{arng('E')}", f"Asientos_Correccion!{arng('R')}", f"Asientos_Correccion!{arng('M')}", f"Asientos_Correccion!{arng('N')}"
blocks = [("ANTES – registros en Odoo", 6), ("AJUSTES PROPUESTOS", 13), ("DESPUÉS – saldos corregidos", 20)]
sub = ["Compras 60", "Gastos vinc. 609", "Destino 61-Compra", "Existencias 2x", "Dif. 60+609 vs 61", "Dif. 61 vs 2x"]
TR0 = 7
cellw(wst, TR0, 1, "Categoría de conciliación", font=f_h, fill=fill_h)
cellw(wst, TR0 + 1, 1, None, fill=fill_h)
wst.merge_cells(start_row=TR0, start_column=1, end_row=TR0 + 1, end_column=1)
widths(wst, {"A": 52})
col0 = {}
for bi_, (t, _) in enumerate(blocks):
    c0 = 2 + bi_ * 7
    col0[t] = c0
    wst.cell(TR0, c0, t).font = f_h
    for k in range(6): wst.cell(TR0, c0 + k).fill = fill_h if bi_ != 1 else fill_h2
    wst.merge_cells(start_row=TR0, start_column=c0, end_row=TR0, end_column=c0 + 5)
    wst.cell(TR0, c0).alignment = Alignment(horizontal="center")
    header(wst, TR0 + 1, c0, sub, fill=fill_h if bi_ != 1 else fill_h2, h=32)
    for k in range(6): wst.column_dimensions[L(c0 + k)].width = 13
    wst.column_dimensions[L(c0 + 6)].width = 2
cellw(wst, TR0, 23, "Explicación de la diferencia remanente (después de ajustes)", font=f_h, fill=fill_h)
wst.merge_cells(start_row=TR0, start_column=23, end_row=TR0 + 1, end_column=23)
wst.column_dimensions["W"].width = 80
after_tbl = M.get("after_expl", {})
EXPL = {
    "C01": "Remanente: devoluciones/retornos sin OC (H19) y redondeo de la reclasificación CG/2026/07/0139.",
    "C02": "Compras/NC sin OC (H19), importe menor.",
    "C03": "60 vs 61: NC 609 sin OC y capitalización parcial de LC (porción de unidades ya consumidas). 61 vs 2x: LC de cacao S/ 657k capitalizado en 24 sin pasar por 60/61 (H09).",
    "C04": "Servicio de subcontratación facturado vs. valor del servicio incorporado al producto (precio OC, mermas/rendimiento de componentes). 61 vs 2x: los componentes salen de 2x contra 61-Consumo/71 (ver Conc_Subcontratacion).",
    "C05": "61 vs 2x: el ingreso de producción y los desmontajes se contrapartidan en 71 y 61-Consumo (no son compras).",
    "C06": "Sin diferencia: la compra de repuestos queda registrada en 60 y destinada en 61.",
    "C07": "Ajuste de inventario de mercaderías (cuenta 61 única) – importe menor.",
    "C08": "Extorno (ene-26) de la provisión de cuadre de dic-25 (CG/2025/12/0123): efecto de apertura; validar contra el corte 2025.",
}
for i, c in enumerate(cats, TR0 + 2):
    cellw(wst, i, 1, CAT_NAME[c], bold=True)
    cr = f'{cat_r},$A{i}'
    a = col0[blocks[0][0]]
    cellw(wst, i, a, f'=SUMIFS({sal_r},{cr},{cls_r},"60")', NUM)
    cellw(wst, i, a + 1, f'=SUMIFS({sal_r},{cr},{cls_r},"609")', NUM)
    cellw(wst, i, a + 2, f'=-SUMIFS({sal_r},{cr},{cls_r},"61C")', NUM)
    cellw(wst, i, a + 3, f'=SUMIFS({sal_r},{cr},{cls_r},"2x")', NUM)
    b = col0[blocks[1][0]]
    code = c
    for k, cl in enumerate(["60", "609"]):
        cellw(wst, i, b + k, f'=SUMIFS({ac_d},{ac_cat},"{code}",{ac_cls},"{cl}")-SUMIFS({ac_h},{ac_cat},"{code}",{ac_cls},"{cl}")', NUM, fill=fill_in)
    cellw(wst, i, b + 2, f'=SUMIFS({ac_h},{ac_cat},"{code}",{ac_cls},"61C")-SUMIFS({ac_d},{ac_cat},"{code}",{ac_cls},"61C")', NUM, fill=fill_in)
    cellw(wst, i, b + 3, f'=SUMIFS({ac_d},{ac_cat},"{code}",{ac_cls},"2x")-SUMIFS({ac_h},{ac_cat},"{code}",{ac_cls},"2x")', NUM, fill=fill_in)
    d = col0[blocks[2][0]]
    for k in range(4):
        cellw(wst, i, d + k, f"={L(a + k)}{i}+{L(b + k)}{i}", NUM)
    for c0 in (a, b, d):
        cellw(wst, i, c0 + 4, f"={L(c0)}{i}+{L(c0 + 1)}{i}-{L(c0 + 2)}{i}", NUM, bold=True)
        cellw(wst, i, c0 + 5, f"={L(c0 + 2)}{i}-{L(c0 + 3)}{i}", NUM, bold=True)
    cellw(wst, i, 23, EXPL[c], wrap=True)
    wst.row_dimensions[i].height = 40
TT = TR0 + 2 + len(cats)
cellw(wst, TT, 1, "TOTAL", bold=True, fill=LGREY)
for c0 in (2, 9, 16):
    for k in range(6):
        x = cellw(wst, TT, c0 + k, f"=SUM({L(c0 + k)}{TR0 + 2}:{L(c0 + k)}{TT - 1})", NUM, bold=True, fill=LGREY)
        x.border = b_top
# fila: sin provisión
cellw(wst, TT + 1, 1, "TOTAL sin provisión de cuadre (C08)", bold=True)
for c0 in (2, 16):
    for k in range(6):
        cellw(wst, TT + 1, c0 + k, f"={L(c0 + k)}{TT}-{L(c0 + k)}{TT - 1}", NUM, bold=True)
wst.cell(TT + 3, 1, "Lectura").font = f_sec
notes = [
    "• En el acumulado Ene–Set la 60 (incluida 609) es igual al destino 61-Compra sólo porque la provisión de cuadre (C08) absorbe la diferencia.",
    "• Sin la provisión, el destino 61-Compra supera a las compras: la diferencia nace de operaciones que no son compras (subcontratación, producción, desmontajes) y de 609 sin capitalizar.",
    "• Los asientos propuestos llevan cada categoría a su lógica correcta; el remanente queda explicado en la columna W y no tiene naturaleza de error de registro.",
    "• Dif. 61 vs 2x ≠ 0 es esperable en C03–C05: hay ingresos a existencias cuya contrapartida no es 61 (71, 61-Consumo, cuentas fuera del alcance).",
]
for k, t in enumerate(notes):
    cellw(wst, TT + 4 + k, 1, t)
wst.freeze_panes = "B9"
for c0 in (2, 9, 16):
    for k in (4, 5):
        col = L(c0 + k)
        wst.conditional_formatting.add(f"{col}{TR0 + 2}:{col}{TT}", CellIsRule(operator="notBetween", formula=["-1", "1"], font=Font(color="C00000", bold=True)))
TRI = dict(TT=TT, TR0=TR0)

# ------------------------------------------------------------------ Puente_Familia
wsf = sheet("Puente_Familia", "Puente por familia: la provisión de cuadre iguala 60 y 61 familia por familia",
            "Familias de conciliación: la 609 se asigna a la familia del bien (6091 mercaderías, 6092 MP, 6093 suministros, 6094 envases y embalajes). "
            "Antes = Odoo; Ajustes = asientos AC; Después = Antes + Ajustes.", tab=NAVY)
fams = ["Mercaderías", "Materias primas", "Suministros y repuestos", "Envases y embalajes"]
FH = ["Familia de conciliación", "Compras 60", "Gastos vinc. 609", "Compras 60 + 609", "Destino 61-Compra (total)", "  de ello: provisión de cuadre",
      "Destino 61-Compra sin provisión", "Dif. con provisión", "Dif. real sin provisión", "Ajustes 60/609 (AC)", "Ajustes destino 61 (AC-02 a AC-14)",
      "Ajuste destino 61 (AC-01 reversión)", "Compras 60+609 corregidas", "Destino 61 corregido", "Dif. después de ajustes", "Explicación del remanente"]
header(wsf, 6, 1, FH, h=48)
widths(wsf, {"A": 24, **{L(k): 14 for k in range(2, 16)}, "P": 70})
fam_r = rng("Familia conciliación")
ac_fam, ac_ac = f"Asientos_Correccion!{arng('H')}", f"Asientos_Correccion!{arng('A')}"
FEXPL = {"Mercaderías": "Extorno de la provisión dic-25 (+S/ 42.5k, apertura), subcontratación (−S/ 18.5k) y devoluciones sin OC (−S/ 6.4k).",
         "Materias primas": "Subcontratación – selección de castaña: servicio facturado vs valorizado (−S/ 53.0k, H04); extorno dic-25 (+S/ 9.4k); compras/NC sin OC (+S/ 6.1k).",
         "Suministros y repuestos": "Extorno de la provisión dic-25 (−S/ 13.7k) y costos en destino (+S/ 1.5k).",
         "Envases y embalajes": "Extorno de la provisión dic-25 (+S/ 93.3k), servicio de armado LIDAN facturado vs valorizado (−S/ 22.1k), NC sin OC (−S/ 5.8k) y LC (−S/ 5.0k)."}
for i, f in enumerate(fams, 7):
    cellw(wsf, i, 1, f, bold=True)
    cellw(wsf, i, 2, f'=SUMIFS({sal_r},{fam_r},$A{i},{cls_r},"60")', NUM)
    cellw(wsf, i, 3, f'=SUMIFS({sal_r},{fam_r},$A{i},{cls_r},"609")', NUM)
    cellw(wsf, i, 4, f"=B{i}+C{i}", NUM, bold=True)
    cellw(wsf, i, 5, f'=-SUMIFS({sal_r},{fam_r},$A{i},{cls_r},"61C")', NUM)
    cellw(wsf, i, 6, f'=-SUMIFS({sal_r},{fam_r},$A{i},{cls_r},"61C",{cat_r},"{CAT_NAME["C08"]}")', NUM)
    cellw(wsf, i, 7, f"=E{i}-F{i}", NUM)
    cellw(wsf, i, 8, f"=D{i}-E{i}", NUM, bold=True)
    cellw(wsf, i, 9, f"=D{i}-G{i}", NUM, bold=True)
    cellw(wsf, i, 10, f'=SUMIFS({ac_d},{ac_fam},$A{i},{ac_cls},"60")-SUMIFS({ac_h},{ac_fam},$A{i},{ac_cls},"60")'
                      f'+SUMIFS({ac_d},{ac_fam},$A{i},{ac_cls},"609")-SUMIFS({ac_h},{ac_fam},$A{i},{ac_cls},"609")', NUM, fill=fill_in)
    cellw(wsf, i, 11, f'=SUMIFS({ac_h},{ac_fam},$A{i},{ac_cls},"61C",{ac_ac},"<>AC-01")-SUMIFS({ac_d},{ac_fam},$A{i},{ac_cls},"61C",{ac_ac},"<>AC-01")', NUM, fill=fill_in)
    cellw(wsf, i, 12, f'=SUMIFS({ac_h},{ac_fam},$A{i},{ac_cls},"61C",{ac_ac},"AC-01")-SUMIFS({ac_d},{ac_fam},$A{i},{ac_cls},"61C",{ac_ac},"AC-01")', NUM, fill=fill_in)
    cellw(wsf, i, 13, f"=D{i}+J{i}", NUM)
    cellw(wsf, i, 14, f"=E{i}+K{i}+L{i}", NUM)
    cellw(wsf, i, 15, f"=M{i}-N{i}", NUM, bold=True)
    cellw(wsf, i, 16, FEXPL[f], wrap=True)
    wsf.row_dimensions[i].height = 30
ft = 7 + len(fams)
cellw(wsf, ft, 1, "TOTAL", bold=True, fill=LGREY)
for k in range(2, 16):
    x = cellw(wsf, ft, k, f"=SUM({L(k)}7:{L(k)}{ft - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
cellw(wsf, ft + 1, 1, "Σ |diferencias| por familia", bold=True)
for k in (8, 9, 15):
    cellw(wsf, ft + 1, k, f"=SUMPRODUCT(ABS({L(k)}7:{L(k)}{ft - 1}))", NUM, bold=True)
for k in (8, 9, 15):
    wsf.conditional_formatting.add(f"{L(k)}7:{L(k)}{ft - 1}", CellIsRule(operator="notBetween", formula=["-1000", "1000"], font=Font(color="C00000", bold=True)))
# composición de la diferencia real por categoría y familia
r2 = ft + 4
wsf.cell(r2 - 1, 1, "Composición de la diferencia real (60+609 − 61-Compra, sin provisión) por categoría de operación").font = f_sec
header(wsf, r2, 1, ["Categoría"] + fams + ["Total"], h=32)
for j, c in enumerate([c for c in CAT_NAME if c != "C08"], r2 + 1):
    cellw(wsf, j, 1, CAT_NAME[c], bold=True)
    for k, f in enumerate(fams, 2):
        fl = f'{fam_r},"{f}",{cat_r},"{CAT_NAME[c]}"'
        cellw(wsf, j, k, f'=SUMIFS({sal_r},{fl},{cls_r},"60")+SUMIFS({sal_r},{fl},{cls_r},"609")+SUMIFS({sal_r},{fl},{cls_r},"61C")', NUM)
    cellw(wsf, j, 2 + len(fams), f"=SUM(B{j}:{L(1 + len(fams))}{j})", NUM, bold=True)
je = r2 + len(CAT_NAME)
cellw(wsf, je, 1, "Diferencia real sin provisión", bold=True, fill=LGREY)
for k in range(2, 3 + len(fams)):
    x = cellw(wsf, je, k, f"=SUM({L(k)}{r2 + 1}:{L(k)}{je - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
cellw(wsf, je + 1, 1, "Control vs. columna I (debe ser 0)")
for k, f in enumerate(fams, 2):
    cellw(wsf, je + 1, k, f"={L(k)}{je}-I{7 + k - 2}", NUM)
wsf.cell(je + 3, 1, "Notas").font = f_sec
for k, t in enumerate([
    "• Con la 609 asignada a su familia, 'Dif. con provisión' es 0 en todas las familias: la provisión mensual se calcula exactamente como el cuadre 60+609 − 61 por familia.",
    "• 'Dif. real sin provisión' es lo que la provisión ocultaba; la tabla inferior la descompone por tipo de operación (subcontratación, 609/LC, producción, etc.).",
    "• 'Dif. después de ajustes' es la que queda tras los asientos AC-01 a AC-14; su explicación está en la columna P.",
]):
    cellw(wsf, je + 4 + k, 1, t)
wsf.freeze_panes = "B7"

# ------------------------------------------------------------------ Matriz_Mensual
wsm = sheet("Matriz_Mensual", "Matriz mensual: la provisión de cuadre iguala 60 y 61 cada mes",
            "Por mes: compras 60+609, destino 61-Compra sin provisión, la provisión neta registrada y la diferencia. Fórmulas SUMIFS sobre Base_Integrada.", tab=NAVY)
MH = ["Mes", "Compras 60+609", "Destino 61-Compra sin provisión", "Dif. real (60 − 61)", "Provisión / extorno neto (Debe 61)", "Dif. después de provisión",
      "% de la dif. real cubierta por la provisión", "Ingreso 2x por compras (C01)", "Facturas 60 de bienes (C01)", "Recepciones 61 (C01)", "Desfase factura vs recepción (C01)"]
header(wsm, 6, 1, MH, h=42)
widths(wsm, {"A": 8, **{L(k): 16 for k in range(2, 12)}})
mes_r = rng("Mes")
c08 = CAT_NAME["C08"]; c01 = CAT_NAME["C01"]
for i, m in enumerate(range(1, 10), 7):
    cellw(wsm, i, 1, m, "0", bold=True, align="center")
    cellw(wsm, i, 2, f'=SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"60")+SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"609")', NUM)
    cellw(wsm, i, 3, f'=-SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"61C")+SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"61C",{cat_r},"{c08}")', NUM)
    cellw(wsm, i, 4, f"=B{i}-C{i}", NUM, bold=True)
    cellw(wsm, i, 5, f'=SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"61C",{cat_r},"{c08}")', NUM)
    cellw(wsm, i, 6, f"=D{i}+E{i}", NUM, bold=True)
    cellw(wsm, i, 7, f'=IFERROR(-E{i}/D{i},"")', PCT)
    cellw(wsm, i, 8, f'=SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"2x",{cat_r},"{c01}")', NUM)
    cellw(wsm, i, 9, f'=SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"60",{cat_r},"{c01}")', NUM)
    cellw(wsm, i, 10, f'=-SUMIFS({sal_r},{mes_r},$A{i},{cls_r},"61C",{cat_r},"{c01}")', NUM)
    cellw(wsm, i, 11, f"=I{i}-J{i}", NUM)
mt = 16
cellw(wsm, mt, 1, "Total", bold=True, fill=LGREY)
for k in range(2, 12):
    if k == 7:
        cellw(wsm, mt, k, f'=IFERROR(-E{mt}/D{mt},"")', PCT, bold=True, fill=LGREY); continue
    x = cellw(wsm, mt, k, f"=SUM({L(k)}7:{L(k)}15)", NUM, bold=True, fill=LGREY); x.border = b_top
wsm.conditional_formatting.add("F7:F15", CellIsRule(operator="between", formula=["-1000", "1000"], fill=PatternFill("solid", fgColor=GREEN)))
ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.title = "60+609 vs 61-Compra sin provisión (S/)"
ch.add_data(Reference(wsm, min_col=2, max_col=3, min_row=6, max_row=15), titles_from_data=True)
ch.set_categories(Reference(wsm, min_col=1, min_row=7, max_row=15))
ch.height, ch.width = 8, 18; ch.y_axis.numFmt = "#,##0"; ch.y_axis.majorGridlines = None
ch2 = LineChart(); ch2.add_data(Reference(wsm, min_col=4, max_col=5, min_row=6, max_row=15), titles_from_data=True)
ch2.y_axis.axId = 200; ch2.y_axis.title = "Dif. / provisión"; ch2.y_axis.crosses = "max"
ch += ch2
wsm.add_chart(ch, "B19")
wsm.freeze_panes = "B7"

# ------------------------------------------------------------------ Provisiones
wsp = sheet("Provisiones", "Provisiones de 'compras no ingresadas': detalle por asiento y familia",
            "Importes tal como fueron registrados en la 61 (Debe + / Haber −). Su contrapartida es 20911101 / 24 / 25 / 26. El saldo neto acumulado es lo que AC-01 revierte (Set + Jul) y lo que queda como apertura (dic-25).", tab=NAVY)
pv = d61[d61.Motivo.str.contains("rovisi")].copy()
pv["Fam"] = pv["Cuenta 61"].astype(str).map({"61211101": "Mercaderías", "61231102": "Materias primas", "61321102": "Suministros",
                                              "61411102": "Envases", "61421102": "Embalajes"})
pt = pv.pivot_table(index=["Asiento", "Mes", "Glosa línea"], columns="Fam", values="Saldo (D-H)", aggfunc="sum", fill_value=0).reset_index()
pf = ["Mercaderías", "Materias primas", "Suministros", "Envases", "Embalajes"]
header(wsp, 6, 1, ["Asiento", "Mes", "Glosa"] + pf + ["Total 61", "Tratamiento"])
widths(wsp, {"A": 17, "B": 6, "C": 62, "D": 13, "E": 14, "F": 13, "G": 13, "H": 13, "I": 13, "J": 46})
for i, r in enumerate(pt.itertuples(index=False), 7):
    cellw(wsp, i, 1, r.Asiento, bold=True); cellw(wsp, i, 2, int(r.Mes), "0", align="center"); cellw(wsp, i, 3, r[2])
    for k, f in enumerate(pf):
        cellw(wsp, i, 4 + k, float(getattr(r, f.replace(" ", "_"), 0) if False else r[3 + list(pt.columns[3:]).index(f)]), NUM)
    cellw(wsp, i, 9, f"=SUM(D{i}:H{i})", NUM, bold=True)
    trat = {"CG/2026/09/0195": "Revertir (AC-01): provisión de cuadre vigente al 30/09",
            "CG/2026/07/0207": "Revertir (AC-01): no fue extornada en agosto",
            "CG/2026/01/0082": "Extorno de dic-25: efecto de apertura (validar con cierre 2025)"}.get(r.Asiento, "Se compensa con su extorno del mes siguiente")
    cellw(wsp, i, 10, trat, fill=RED if "Revertir" in trat else (AMBER if "apertura" in trat else None))
pe = 7 + len(pt)
cellw(wsp, pe, 3, "Saldo neto acumulado al 30/09/2026", bold=True, fill=LGREY)
for k in range(4, 10):
    x = cellw(wsp, pe, k, f"=SUM({L(k)}7:{L(k)}{pe - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
cellw(wsp, pe + 2, 3, "Prueba de 'cuadre': provisión neta del mes = diferencia real 60 − 61 del mes (hoja Matriz_Mensual, col. G ≈ 100%)", bold=True)
wsp.freeze_panes = "D7"

# ------------------------------------------------------------------ Conc_Subcontratacion
wss = sheet("Conc_Subcontratacion", "Subcontratación (SBC): componentes, producto recibido y servicio facturado",
            "Bloque 1: reclasificación de consumos por familia. Bloque 2: conciliación por OC de subcontratación (servicio implícito = ingreso del producto − componentes).", tab=BLUE)
header(wss, 6, 1, ["Familia", "Consumo SBC (61-Consumo)", "Reclasificado CG/2026/07/0097", "  a 61-Compra", "  a 71", "Pendiente de reclasificar", "Asiento propuesto"], h=36)
widths(wss, {"A": 18, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 13, "H": 30, "I": 15, "J": 15, "K": 15, "L": 15, "M": 15, "N": 22, "O": 7, "P": 60})
for i, (f, r) in enumerate(sbc_pend.iterrows(), 7):
    cellw(wss, i, 1, f, bold=True)
    cellw(wss, i, 2, f'=SUMIFS({sal_r},{fam_r},$A{i},{cls_r},"61K",{rng("Tipo de operación")},"OP09*")', NUM)
    cellw(wss, i, 3, float(r["Reclasificado (Haber 61K)"]), NUM)
    cellw(wss, i, 4, float(r["Reclasificado a 61C"]), NUM); cellw(wss, i, 5, float(r["Reclasificado a 71"]), NUM)
    cellw(wss, i, 6, f"=B{i}-C{i}", NUM, bold=True); cellw(wss, i, 7, "AC-02", align="center")
se = 7 + len(sbc_pend)
cellw(wss, se, 1, "Total", bold=True, fill=LGREY)
for k in range(2, 7):
    x = cellw(wss, se, k, f"=SUM({L(k)}7:{L(k)}{se - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
# mensual
r0 = se + 2
wss.cell(r0, 1, "Consumo de componentes SBC por mes (61-Consumo) – la reclasificación sólo alcanzó Ene–May").font = f_sec
mf = list(sbc_month.columns)
header(wss, r0 + 1, 1, ["Mes"] + mf + ["Total", "¿Reclasificado?"])
for i, m in enumerate(range(1, 10), r0 + 2):
    cellw(wss, i, 1, m, "0", align="center", bold=True)
    for k, f in enumerate(mf):
        cellw(wss, i, 2 + k, f'=SUMIFS({sal_r},{mes_r},$A{i},{fam_r},"{f}",{cls_r},"61K",{rng("Tipo de operación")},"OP09*")', NUM)
    cellw(wss, i, 2 + len(mf), f"=SUM(B{i}:{L(1 + len(mf))}{i})", NUM, bold=True)
    cellw(wss, i, 3 + len(mf), "Sí (CG/2026/07/0097)" if m <= 5 else "No", fill=GREEN if m <= 5 else RED)
# por OC
r1 = r0 + 13
wss.cell(r1, 1, "Conciliación por OC de subcontratación").font = f_sec
SH = ["OC", "Subcontratista", "Producto recibido", "Ingreso producto (2x)", "Abono 61-Compra", "Componentes (61-Consumo)",
      "Servicio facturado (60/609)", "Servicio implícito", "Dif. facturado vs implícito", "Estado", "Código", "Explicación"]
header(wss, r1 + 1, 1, SH, h=36)
SO = so.sort_values("Estado")
for i, r in enumerate(SO.itertuples(index=False), r1 + 2):
    cellw(wss, i, 1, r.OC, bold=True); cellw(wss, i, 2, r.Subcontratista); cellw(wss, i, 3, r[2])
    cellw(wss, i, 4, float(r[3]), NUM); cellw(wss, i, 5, float(-r[4]), NUM); cellw(wss, i, 6, float(r[5]), NUM); cellw(wss, i, 7, float(r[6]), NUM)
    cellw(wss, i, 8, f"=D{i}-F{i}", NUM); cellw(wss, i, 9, f"=G{i}-H{i}", NUM, bold=True)
    cellw(wss, i, 10, r.Estado); cellw(wss, i, 11, r.Código, align="center"); cellw(wss, i, 12, r.Explicación)
s1 = r1 + 2 + len(SO)
cellw(wss, s1, 1, "Total", bold=True, fill=LGREY)
for k in range(4, 10):
    x = cellw(wss, s1, k, f"=SUM({L(k)}{r1 + 2}:{L(k)}{s1 - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
for st, col in STATUS_FILL.items():
    wss.conditional_formatting.add(f"J{r1 + 2}:J{s1}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=col)))
widths(wss, {"C": 40, "L": 70})

# ------------------------------------------------------------------ Conc_609_LC
wsl = sheet("Conc_609_LC", "Gastos vinculados (609) vs costos en destino (LC) por OC de servicio",
            "Una factura 609 debe trasladarse al inventario mediante un costo en destino (Debe 2x / Haber 61). Diferencia = 609 facturado − LC capitalizado.", tab=BLUE)
LH = ["OC servicio", "Proveedor", "Glosa", "Cuenta 60/609", "Mes", "Factura 609", "LC abonado a 61", "LC abonado a 609", "LC cargado a 2x",
      "Capitalizado (61+609)", "Dif. 609 vs capitalizado", "Estado", "Código", "Explicación", "Asiento"]
header(wsl, 6, 1, LH, h=36)
widths(wsl, {"A": 11, "B": 30, "C": 46, "D": 18, "E": 6, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 22, "M": 7, "N": 60, "O": 9})
LT = lt.sort_values(["Estado", "Factura 609"], ascending=[True, False])
for i, r in enumerate(LT.itertuples(index=False), 7):
    cellw(wsl, i, 1, r[0], bold=True); cellw(wsl, i, 2, r.Proveedor); cellw(wsl, i, 3, r.Glosa); cellw(wsl, i, 4, r[3])
    cellw(wsl, i, 5, None if pd.isna(r.Mes) else int(r.Mes), "0", align="center")
    cellw(wsl, i, 6, float(r[5]), NUM2); cellw(wsl, i, 7, float(-r[6]), NUM2); cellw(wsl, i, 8, float(-r[7]), NUM2); cellw(wsl, i, 9, float(r[8]), NUM2)
    cellw(wsl, i, 10, f"=G{i}+H{i}", NUM2); cellw(wsl, i, 11, f"=F{i}-J{i}", NUM2, bold=True)
    cellw(wsl, i, 12, r.Estado); cellw(wsl, i, 13, r.Código, align="center"); cellw(wsl, i, 14, r.Explicación)
    cellw(wsl, i, 15, CODE2AC.get(r.Código, "") if r.Estado.startswith("Error") else ("Revisión" if r.Estado.startswith("Obs") else ""), align="center")
le = 7 + len(LT)
cellw(wsl, le, 1, "Total", bold=True, fill=LGREY)
for k in range(6, 12):
    x = cellw(wsl, le, k, f"=SUM({L(k)}7:{L(k)}{le - 1})", NUM, bold=True, fill=LGREY); x.border = b_top
for st, col in STATUS_FILL.items():
    wsl.conditional_formatting.add(f"L7:L{le}", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=col)))
wsl.auto_filter.ref = f"A6:O{le - 1}"
wsl.freeze_panes = "B7"

# ------------------------------------------------------------------ Resumen
wsr = sheet("Resumen", "Resumen ejecutivo – Conciliación de compras 60 / 61 / 2x, La Ibérica, Ene–Set 2026",
            "Fuente: Odoo (Detalle_61, Base_Cta60, Base_Cta2). Todas las cifras están formuladas y enlazadas a las hojas de detalle.", tab=NAVY)
widths(wsr, {"A": 2, "B": 5, "C": 30, "D": 14, "E": 14, "F": 18, "G": 20, "H": 64, "I": 12, "J": 16, "K": 11})
# KPIs
kpis = [("Compras 60 + 609", f"=Triangulo!B{TT}+Triangulo!C{TT}"),
        ("Destino 61-Compra", f"=Triangulo!D{TT}"),
        ("Diferencia oculta por la provisión", f"=-Triangulo!D{TT - 1}"),
        ("Claves OC|Producto que cuadran", f'=COUNTIFS({KT_RNG("S")},"Cuadra")/COUNTA({KT_RNG("S")})'),
        ("Importe bruto de asientos propuestos", f"=Asientos_Correccion!E{it}")]
for k, (t, f) in enumerate(kpis):
    c = 2 + k if k < 1 else 2 + k
for k, (t, f) in enumerate(kpis):
    col = [3, 4, 5, 6, 7][k]
    x = wsr.cell(5, col, t); x.font = Font(name=FONT, size=8, color=GREY, bold=True); x.alignment = Alignment(wrap_text=True, vertical="bottom")
    y = wsr.cell(6, col, f); y.font = Font(name=FONT, size=16, bold=True, color=NAVY)
    y.number_format = PCT if "Claves" in t else '#,##0,"k"'
    for rr in (5, 6):
        wsr.cell(rr, col).fill = PatternFill("solid", fgColor="EEF2F7")
    wsr.cell(6, col).border = Border(bottom=Side(style="thick", color=BLUE))
wsr.row_dimensions[5].height = 30; wsr.row_dimensions[6].height = 30

# mensajes clave (calculados)
d_sinprov = -(bi.loc[(bi.Clase == "61C") & (bi.Cat == "C08"), "Saldo"].sum())
_x = bi[bi.Clase.isin(["60", "609", "61C"]) & (bi.Cat != "C08")]
famabs = _x.groupby(_x.Cuenta.map(fam_conc))["Saldo"].sum().abs().sum()
n_ok = int((kt.Estado == "Conciliado").sum()); n_k = len(kt)
v_ok = kt.loc[kt.Estado == "Conciliado", "F"].sum(); v_k = kt.F.sum()
sbc_p = sbc_pend["Pendiente"].sum()
ub = -bi.loc[(bi["Tipo de operación"].str.startswith("OP11")) & (bi.Clase == "61C"), "Saldo"].sum()
pp = -bi.loc[(bi["Tipo de operación"].str.startswith("OP10")) & (bi.Clase == "61C"), "Saldo"].sum()
lc609 = ac.loc[ac.AC == "AC-05", "Debe"].sum(); cap609 = ac.loc[ac.AC == "AC-06", "Debe"].sum()
rep = ac.loc[ac.AC == "AC-12", "Debe"].sum()
cacao = bi.loc[(bi["Código hallazgo"] == "H09") & (bi.Clase == "2x"), "Saldo"].sum()
_p = ac[(ac.AC == "AC-01") & (ac.Clase == "61C")]
prov_sep = (_p.Haber - _p.Debe).sum()
msgs = [
    ("1", f"60 = 61 en el acumulado sólo gracias a una provisión manual de cuadre",
     f"La provisión mensual 'compras no ingresadas' se calcula como 60 − 61 por familia y se registra contra existencias (2x). "
     f"Sin ella, el destino 61-Compra excede a las compras en S/ {abs(d_sinprov)/1e3:,.0f}k y las diferencias por familia suman S/ {famabs/1e3:,.0f}k en valor absoluto (ver Puente_Familia)."),
    ("2", f"El núcleo de compras sí concilia: {n_ok:,} de {n_k:,} claves OC|Producto ({v_ok / v_k:.0%} del valor facturado)",
     "Factura = recepción = ingreso a existencias. Las diferencias por clave son pocas y se explican por precio/TC (H11) y corte (H12/H13)."),
    ("3", "Las diferencias estructurales vienen de operaciones que no son compras y que Odoo registra en 61-Compra",
     f"Subcontratación: S/ {sbc_p/1e3:,.0f}k de componentes sin reclasificar (jun–set); desmontajes: S/ {abs(ub)/1e3:,.0f}k netos; "
     f"producción propia: S/ {pp/1e3:,.0f}k abonados a 61-Compra en lugar de 71."),
    ("4", "Gastos vinculados y costos en destino no siguen un circuito único",
     f"S/ {lc609/1e3:,.0f}k de LC abonados a 609, S/ {cap609/1e3:,.0f}k de 609 sin capitalizar y S/ {cacao/1e3:,.0f}k de LC de cacao cargados a 24 sin pasar por 60/61 (por identificar contrapartida)."),
    ("5", f"14 asientos de corrección ({len(ac):,} líneas, todos cuadrados) dejan la conciliación explicada",
     f"Incluyen la reversión neta del cuadre (S/ {prov_sep/1e3:,.0f}k en 61), el corte real, las reclasificaciones 61-Compra/61-Consumo/71 y la compra de repuestos cargada directo a 2x (S/ {rep/1e3:,.0f}k). "
     "El remanente se limita al extorno de apertura dic-25 y a la revisión de subcontratación."),
]
wsr.cell(8, 2, "Mensajes clave").font = f_sec
r = 9
for n, t, d in msgs:
    a = wsr.cell(r, 2, n); a.font = Font(name=FONT, size=14, bold=True, color=BLUE); a.alignment = Alignment(horizontal="center", vertical="top")
    b = wsr.cell(r, 3, t); b.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    wsr.merge_cells(start_row=r, start_column=3, end_row=r, end_column=11)
    c = wsr.cell(r + 1, 3, d); c.font = f_b; c.alignment = Alignment(wrap_text=True, vertical="top")
    wsr.merge_cells(start_row=r + 1, start_column=3, end_row=r + 1, end_column=11)
    wsr.row_dimensions[r + 1].height = 30
    r += 2
# estados
r += 1
wsr.cell(r, 2, "Estado de conciliación de las líneas").font = f_sec
header(wsr, r + 1, 2, ["", "Estado", "Líneas", "% líneas", "Importe abs. (S/)"])
wsr.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=2)
est_r, abs_r = rng("Estado conciliación"), rng("Importe abs")
sts = ["Conciliado", "Conciliado (corregido)", "Diferencia temporal (corte)", "Error – requiere asiento", "Observación"]
for k, s_ in enumerate(sts, r + 2):
    wsr.cell(k, 2).fill = PatternFill("solid", fgColor=STATUS_FILL[s_])
    cellw(wsr, k, 3, s_, bold=True)
    cellw(wsr, k, 4, f'=COUNTIFS({est_r},C{k})', "#,##0")
    cellw(wsr, k, 5, f"=D{k}/SUM($D${r + 2}:$D${r + 1 + len(sts)})", PCT)
    cellw(wsr, k, 6, f'=SUMIFS({abs_r},{est_r},C{k})', NUM)
ke = r + 2 + len(sts)
cellw(wsr, ke, 3, "Total", bold=True, fill=LGREY)
for kk in (4, 5, 6):
    x = cellw(wsr, ke, kk, f"=SUM({L(kk)}{r + 2}:{L(kk)}{ke - 1})", "#,##0" if kk == 4 else (PCT if kk == 5 else NUM), bold=True, fill=LGREY); x.border = b_top
# top hallazgos (derecha)
wsr.cell(r, 8, "Principales hallazgos (importe medido)").font = f_sec
header(wsr, r + 1, 8, ["Hallazgo", "Asiento", "Importe medido (abs.)", "Prioridad"])
top = ["H01", "H03", "H05", "H06", "H08", "H07", "H15", "H11", "H09"]
for k, code in enumerate(top, r + 2):
    hrow = 7 + [h[0] for h in HALL].index(code)
    cellw(wsr, k, 8, f'=Hallazgos!A{hrow}&" – "&Hallazgos!B{hrow}', wrap=False)
    cellw(wsr, k, 9, f"=Hallazgos!K{hrow}", align="center")
    cellw(wsr, k, 10, f"=ABS(Hallazgos!I{hrow})", NUM)
    cellw(wsr, k, 11, f"=Hallazgos!D{hrow}", align="center")
wsr.conditional_formatting.add(f"J{r + 2}:J{r + 1 + len(top)}", DataBarRule(start_type="num", start_value=0, end_type="max", color=BLUE))
# gráfico mensual
ch = BarChart(); ch.type = "col"; ch.title = "Compras 60+609 vs destino 61-Compra sin provisión, por mes"
ch.add_data(Reference(wsm, min_col=2, max_col=3, min_row=6, max_row=15), titles_from_data=True)
ch.set_categories(Reference(wsm, min_col=1, min_row=7, max_row=15)); ch.height, ch.width = 7.5, 16; ch.y_axis.numFmt = '#,##0,"k"'
ch.y_axis.majorGridlines = None
cr0 = max(ke, r + 1 + len(top)) + 3
wsr.add_chart(ch, f"B{cr0}")
ch3 = BarChart(); ch3.type = "col"; ch3.title = "Provisión de cuadre vs diferencia real del mes"
ch3.add_data(Reference(wsm, min_col=4, max_col=5, min_row=6, max_row=15), titles_from_data=True)
ch3.set_categories(Reference(wsm, min_col=1, min_row=7, max_row=15)); ch3.height, ch3.width = 7.5, 16; ch3.y_axis.numFmt = '#,##0,"k"'
ch3.y_axis.majorGridlines = None
wsr.add_chart(ch3, f"G{cr0}")

# ------------------------------------------------------------------ Metodología
wsx = sheet("Metodologia", "Metodología, definiciones y supuestos", "Cómo se construyó la conciliación y cómo leer cada hoja.", tab=GREY)
widths(wsx, {"A": 30, "B": 110})
METH = [
    ("Objetivo", "Conciliar las compras entre la cuenta 60 (compras), la 61 (variación de existencias – destino) y la clase 2 (existencias), identificar las líneas que no cuadran, su causa en Odoo y los asientos que las corrigen."),
    ("Período y fuente", "Enero–Setiembre 2026. Hojas Detalle_61 (85,110 líneas), Base_Cta60 (3,989) y Base_Cta2 (229,938) exportadas de Odoo. Se trabajó sobre el perímetro de compras: toda la 60/609, la 61 subcuenta Compra/Única y la 61-Consumo vinculada a subcontratación, desmontajes y costos en destino, y las líneas 2x de esos asientos o con motivo de compra."),
    ("Lógica contable (PCGE)", "Compra: Debe 60 / Haber 42. Destino: Debe 2x / Haber 61 (subcuenta Compra x1102 o Única 61211101). Consumo: Debe 61 (subcuenta Consumo x1101) / Haber 2x. Por tanto, en un circuito sano: Compras 60 (incl. 609 capitalizada) = Destino 61-Compra = Ingreso a 2x por compras."),
    ("Clases", "60 = compras de bienes; 609 = costos vinculados; 61C = 61 subcuentas Compra (x1102) y Única (61211101); 61K = 61 subcuentas Consumo (x1101); 2x = existencias."),
    ("Destino 61-Compra", "Se presenta en positivo (= abono). Dif. 60 vs 61 = (60 + 609) − Destino 61-Compra. Dif. 61 vs 2x = Destino 61-Compra − Ingreso 2x."),
    ("Conciliación por clave", "Clave = OC principal | Producto. Para cada clave se compara factura/NC (60), recepción/devolución (61-Compra) e ingreso/salida en 2x. Tolerancia ±S/ 1. Una clave con importes iguales pero meses distintos se considera conciliada (desfase temporal resuelto)."),
    ("Categorías de conciliación", "C01 compras con OC; C02 compras sin OC; C03 609 y costos en destino; C04 subcontratación; C05 producción propia y desmontajes; C06 gastos imputados a 61/2x y reclasificaciones; C07 ajustes de inventario; C08 provisión de cuadre."),
    ("Estados", "Conciliado; Conciliado (corregido) = ya regularizado en Odoo con asiento manual; Diferencia temporal (corte) = se resuelve con el registro pendiente (extornable); Error – requiere asiento = tiene asiento propuesto; Observación = requiere revisión documental antes de corregir."),
    ("Asiento: 60 / 609 / 61C / 61K / 2x / otras ctas", "Totales por clase del asiento Odoo al que pertenece la línea (con todas sus líneas 60/61/2x). 'Otras ctas' = −(suma), es decir, el importe contabilizado en cuentas fuera del alcance (42, 40, 63, 69, 71, 79…)."),
    ("Asientos de corrección", "Fecha 30/09/2026. Cuentas marcadas '(cuenta sugerida)' (28x, 4211x, 6033x) no existen en el extracto y deben crearse o sustituirse por la equivalente del plan de La Ibérica. AC-09/AC-10 son de corte y se extornan al registrar la recepción/factura."),
    ("Supuestos y límites", "No se dispone de las cuentas 42, 63, 69, 71 y 79 completas: las contrapartidas fuera de 60/61/2x se infieren por diferencia. El extorno de la provisión de dic-25 se deja como efecto de apertura. Las cantidades pueden estar en UdM distintas entre factura y recepción; por eso la diferencia de precio se evalúa también por importe (≤2%)."),
    ("Cómo rastrear una cifra", "1) Ubicar el hallazgo o la categoría en Resumen/Triangulo. 2) Filtrar Base_Integrada por 'Código hallazgo' o 'Categoría de conciliación'. 3) Para compras con OC, usar 'ID Clave' y la hoja Conc_OC_Producto. 4) El asiento propuesto está en 'Asiento corrección' y en Asientos_Correccion."),
]
header(wsx, 5, 1, ["Tema", "Detalle"])
for i, (a, b) in enumerate(METH, 6):
    cellw(wsx, i, 1, a, bold=True, wrap=True); cellw(wsx, i, 2, b, wrap=True)
    wsx.row_dimensions[i].height = 15 * max(2, len(b) // 120 + 1)
r = 6 + len(METH) + 1
wsx.cell(r, 1, "Tipos de operación").font = f_sec
ops = sorted(bi["Tipo de operación"].unique())
for k, o in enumerate(ops, r + 1):
    cellw(wsx, k, 1, o[:4], bold=True); cellw(wsx, k, 2, o[5:])

# ------------------------------------------------------------------ Inicio
wsi = wb.create_sheet("Inicio", 0)
wsi.sheet_view.showGridLines = False
wsi.sheet_properties.tabColor = NAVY
widths(wsi, {"A": 3, "B": 28, "C": 90})
wsi["B2"] = "La Ibérica – Conciliación de compras 60 / 61 / 2x"; wsi["B2"].font = Font(name=FONT, size=22, bold=True, color=NAVY)
wsi["B3"] = "Enero – Setiembre 2026 · Fuente: Odoo · Cifras en soles (S/)"; wsi["B3"].font = f_sub
wsi["B5"] = "Contenido"; wsi["B5"].font = f_sec
IDX = [("Resumen", "Mensajes clave, KPIs, estado de conciliación y principales hallazgos."),
       ("Triangulo", "Conciliación 60 ↔ 61-Compra ↔ 2x por categoría de operación: antes, ajustes y después."),
       ("Puente_Familia", "Diferencias por familia, con y sin la provisión de cuadre, antes y después de ajustes."),
       ("Matriz_Mensual", "Evolución mensual y prueba de que la provisión iguala 60 y 61 cada mes."),
       ("Hallazgos", "Catálogo H01–H19: causa raíz en Odoo, importe, prioridad, responsable y asiento propuesto."),
       ("Asientos_Correccion", "Asientos AC-01 a AC-14 con detalle Debe/Haber, referencias y validación de cuadre."),
       ("Conc_OC_Producto", "Conciliación factura ↔ recepción ↔ 2x por OC y producto (fórmulas sobre la base)."),
       ("Conc_Subcontratacion", "Componentes SBC, reclasificaciones y conciliación por OC de subcontratación."),
       ("Conc_609_LC", "Gastos vinculados (609) vs costos en destino por OC de servicio."),
       ("Provisiones", "Detalle de las provisiones de 'compras no ingresadas' y su tratamiento."),
       ("Base_Integrada", "Todas las líneas del perímetro con tipo de operación, estado, hallazgo y asiento de corrección."),
       ("Metodologia", "Definiciones, supuestos, límites y cómo rastrear una cifra.")]
for i, (s_, d) in enumerate(IDX, 6):
    a = wsi.cell(i, 2, s_); a.hyperlink = f"#'{s_}'!A1"; a.font = f_link
    cellw(wsi, i, 3, d)
r = 6 + len(IDX) + 1
wsi.cell(r, 2, "Leyenda de estados").font = f_sec
for k, (s_, col) in enumerate(STATUS_FILL.items(), r + 1):
    a = wsi.cell(k, 2, s_); a.fill = PatternFill("solid", fgColor=col); a.font = f_bb
    cellw(wsi, k, 3, {"Conciliado": "60 = 61 = 2x para la clave / asiento.",
                      "Conciliado (corregido)": "Diferencia ya regularizada en Odoo con un asiento manual existente.",
                      "Diferencia temporal (corte)": "Se resuelve con el registro pendiente (recepción o factura). Asiento de corte extornable.",
                      "Error – requiere asiento": "Registro o configuración incorrecta. Tiene asiento de corrección propuesto.",
                      "Observación": "Requiere revisión documental antes de decidir el ajuste."}[s_])
wsi.cell(r + 7, 2, "Convenciones").font = f_sec
cellw(wsi, r + 8, 2, "Signos"); cellw(wsi, r + 8, 3, "Saldo = Debe − Haber. 'Destino 61-Compra' se muestra en positivo (abono). Paréntesis = negativo.")
cellw(wsi, r + 9, 2, "Celdas azul claro"); cellw(wsi, r + 9, 3, "Ajustes propuestos (provienen de Asientos_Correccion).")
cellw(wsi, r + 10, 2, "Recalcular"); cellw(wsi, r + 10, 3, "El libro se recalcula al abrir (todas las cifras son fórmulas). Si ve celdas vacías, presione F9.")

order = ["Inicio", "Resumen", "Triangulo", "Puente_Familia", "Matriz_Mensual", "Hallazgos", "Asientos_Correccion", "Conc_OC_Producto",
         "Conc_Subcontratacion", "Conc_609_LC", "Provisiones", "Base_Integrada", "Metodologia"]
wb._sheets = [wb[s_] for s_ in order]
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    if ws.title not in ("Base_Integrada",):
        ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1
wb.calculation.fullCalcOnLoad = True
wb.active = 0
wb.save(OUT)
print("ok", OUT)
