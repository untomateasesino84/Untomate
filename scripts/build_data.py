"""Motor de conciliación 60 / 61 / 2x – La Ibérica Ene-Set 2026 (datos Odoo)."""
import pandas as pd, numpy as np, pickle, sys

W = sys.argv[1]
d60 = pd.read_pickle(f"{W}/Base_Cta60.pkl")
d61 = pd.read_pickle(f"{W}/Detalle_61.pkl")
d2 = pd.read_pickle(f"{W}/Base_Cta2.pkl")
TOL = 1.0

# ---------------------------------------------------------------- normalización
d61 = d61.rename(columns={"Cuenta 61": "Cuenta", "Glosa línea": "Glosa", "Referencia asiento": "Referencia",
                          "Categoría producto": "Categoría", "OC secundaria (servicio / LC)": "OC secundaria",
                          "Orden de fabricación (OF)": "OF"})
for d, src in [(d60, "60"), (d61, "61"), (d2, "2x")]:
    d["Fuente"] = src
    d["Cuenta"] = d["Cuenta"].astype(str).str.replace(".0", "", regex=False)
    d["Mes"] = d["Mes"].astype(int)
    for c in ["Debe", "Haber", "Saldo (D-H)", "Cantidad"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
d61["Subcuenta"] = d61["Subcuenta"].fillna("")

FAM = {"6011": "Mercaderías", "6021": "Materias primas", "6032": "Suministros", "6041": "Envases", "6042": "Embalajes",
       "609": "Gastos vinculados (609)",
       "6121": "Mercaderías", "6123": "Materias primas", "6132": "Suministros", "6133": "Repuestos",
       "6141": "Envases", "6142": "Embalajes",
       "20": "Mercaderías", "21": "Productos terminados", "23": "Productos en proceso", "24": "Materias primas",
       "251": "Suministros", "252": "Suministros", "253": "Repuestos", "261": "Envases", "262": "Embalajes"}


def fam(acc):
    for k in sorted(FAM, key=len, reverse=True):
        if acc.startswith(k):
            return FAM[k]
    return "Otros"


def clase(row):
    a = row["Cuenta"]
    if a.startswith("609"):
        return "609"
    if a.startswith("60"):
        return "60"
    if a.startswith("61"):
        return "61C" if row["Subcuenta"] in ("COMPRA", "ÚNICA") else "61K"
    return "2x"


# OCs especiales
SBC_OC = set(d61.loc[d61["Tipo vínculo OC"].fillna("").str.startswith("OC de subcontratación"), "OC principal"].dropna())
LC_OC = set(d61["OC secundaria"].dropna()) | set(d60.loc[d60.Motivo == "Costo en destino (Landed cost)", "OC secundaria"].dropna()) \
    | set(d2.loc[d2.Motivo == "Costo en destino (Landed cost)", "OC secundaria"].dropna())

# ---------------------------------------------------------------- alcance
MOT_SCOPE61K = ["Consumo MP enviada a subcontratación (SBC)", "Reclasificación de consumos (SBC/Unbuild)",
                "Desmontaje OP (Unbuild)", "Costo en destino (Landed cost)", "Factura proveedor imputada a 61"]
s61 = d61[(d61.Subcuenta.isin(["COMPRA", "ÚNICA"])) | (d61.Motivo.isin(MOT_SCOPE61K))].copy()
MOT_PURCH = ["Compra – recepción de proveedor", "Devolución a proveedor", "Reingreso por devolución",
             "Costo en destino (Landed cost)", "Provisión compras no ingresadas",
             "Extorno provisión compras no ingresadas", "Factura de proveedor"]
ids = set(d60["ID asiento"]) | set(s61["ID asiento"])
s2 = d2[(d2.Motivo.isin(MOT_PURCH)) | (d2["ID asiento"].isin(ids) & d2.Motivo.isin(
    MOT_PURCH + ["Ingreso de producción (OP)", "Consumo MP enviada a subcontratación (SBC)", "Desmontaje OP (Unbuild)",
                 "Reclasificación manual", "Ajuste de inventario", "Ajuste de valor sin movimiento de stock"]))].copy()
# Ingresos SBC en 2x (para conciliar subcontratación) aunque su asiento no tenga 61
s2 = pd.concat([s2, d2[(d2.Motivo == "Ingreso de producción (OP)") & (d2["Tipo operación almacén"].fillna("").str.contains("Subcontrat"))
                       & ~d2.index.isin(s2.index)]])

COLS = ["Fuente", "ID asiento", "Asiento", "Fecha", "Mes", "Diario", "Tipo documento", "Referencia", "Glosa", "Cuenta",
        "Subcuenta", "Debe", "Haber", "Saldo (D-H)", "Producto", "Categoría", "Cantidad", "UdM", "Socio", "Motivo",
        "Origen asiento", "Tipo operación almacén", "Ref. movimiento stock", "Flujo", "OC principal", "OC secundaria",
        "Costo en destino (LC)", "Clave OC|Producto", "Creado por", "Tipo vínculo OC", "OF"]
for d in (d60, s61, s2):
    for c in COLS:
        if c not in d:
            d[c] = np.nan
d60["Subcuenta"] = ""
s2["Subcuenta"] = ""
bi = pd.concat([d60[COLS], s61[COLS], s2[COLS]], ignore_index=True)
bi["Fecha"] = pd.to_datetime(bi["Fecha"], unit="D", origin="1899-12-30", errors="coerce")
bi["Clase"] = bi.apply(clase, axis=1)
bi["Familia"] = bi["Cuenta"].map(fam)
bi["Saldo"] = bi["Saldo (D-H)"].round(2)
bi["Glosa"] = bi["Glosa"].fillna(bi["Referencia"])
bi["Clave"] = bi["Clave OC|Producto"].where(bi["OC principal"].notna())

# ---------------------------------------------------------------- tipo de operación
def tipo_op(r):
    m, c, oc = r.Motivo, r.Clase, r["OC principal"]
    tov = str(r["Tipo vínculo OC"]) if pd.notna(r["Tipo vínculo OC"]) else ""
    tom = str(r["Tipo operación almacén"]) if pd.notna(r["Tipo operación almacén"]) else ""
    if r.Fuente == "60":
        if m in ("Factura de proveedor", "Nota de crédito de proveedor") and (
                oc in SBC_OC or (pd.isna(oc) and "ARMAD" in str(r.Producto).upper())):
            return "OP07 Factura servicio subcontratación"
        if m == "Factura de proveedor":
            if c == "609":
                return "OP03 Factura gasto vinculado (609)"
            if oc in LC_OC:
                return "OP03 Factura gasto vinculado (609)"
            return "OP01 Factura de compra de bienes"
        if m == "Nota de crédito de proveedor":
            return "OP03 Factura gasto vinculado (609)" if c == "609" else "OP02 Nota de crédito de proveedor"
        if m == "Costo en destino (Landed cost)":
            return "OP06 Costo en destino (LC)"
        if m == "Devolución a proveedor":
            return "OP05 Devolución a proveedor"
        if m == "Ajuste de valor sin movimiento de stock":
            return "OP16 Ajuste de inventario / revaloración"
        return "OP15 Reclasificación manual"
    if m == "Compra – recepción de proveedor":
        return "OP04 Recepción de compra"
    if m in ("Devolución a proveedor", "Reingreso por devolución"):
        return "OP05 Devolución a proveedor"
    if m == "Costo en destino (Landed cost)":
        return "OP06 Costo en destino (LC)"
    if m == "Ingreso de producción (OP)":
        if tov.startswith("OC de subcontratación") or "Subcontrat" in tom:
            return "OP08 Subcontratación – ingreso producto"
        return "OP10 Producción propia – ingreso"
    if m == "Consumo MP enviada a subcontratación (SBC)":
        return "OP09 Subcontratación – consumo componentes"
    if m == "Reclasificación de consumos (SBC/Unbuild)":
        return "OP12 Reclasificación SBC/UB (manual)" if "TIPO UB" not in str(r.Glosa) else "OP11 Desmontaje (Unbuild)"
    if m == "Desmontaje OP (Unbuild)":
        return "OP11 Desmontaje (Unbuild)"
    if m == "Provisión compras no ingresadas":
        return "OP13 Provisión compras no ingresadas"
    if m == "Extorno provisión compras no ingresadas":
        return "OP14 Extorno provisión"
    if m == "Factura proveedor imputada a 61" or m == "Factura de proveedor":
        return "OP17 Factura imputada a 61 / 2x"
    if m == "Reclasificación manual":
        return "OP15 Reclasificación manual"
    if m in ("Ajuste de inventario", "Ajuste de valor sin movimiento de stock", "Merma / Scrap"):
        return "OP16 Ajuste de inventario / revaloración"
    return "OP18 Otros movimientos de inventario"


bi["Tipo de operación"] = bi.apply(tipo_op, axis=1)

CAT = {"OP01": "C01", "OP02": "C01", "OP04": "C01", "OP05": "C01",
       "OP03": "C03", "OP06": "C03",
       "OP07": "C04", "OP08": "C04", "OP09": "C04", "OP12": "C04",
       "OP10": "C05", "OP11": "C05",
       "OP17": "C06", "OP15": "C06",
       "OP16": "C07", "OP18": "C07",
       "OP13": "C08", "OP14": "C08"}
CAT_NAME = {"C01": "C01 Compras de bienes: factura ↔ recepción / devolución",
            "C02": "C02 Compras y NC sin OC ni recepción",
            "C03": "C03 Gastos vinculados (609) y costos en destino (LC)",
            "C04": "C04 Subcontratación (SBC)",
            "C05": "C05 Producción propia y desmontajes (Unbuild)",
            "C06": "C06 Gastos imputados a 61/2x y reclasificaciones manuales",
            "C07": "C07 Ajustes de inventario y revaloraciones",
            "C08": "C08 Provisión 'compras no ingresadas' (cuadre manual)"}
bi["Cat"] = bi["Tipo de operación"].str[:4].map(CAT)
# facturas / NC sin OC → C02
m_sinoc = bi["Tipo de operación"].str[:4].isin(["OP01", "OP02"]) & bi["OC principal"].isna()
bi.loc[m_sinoc, "Cat"] = "C02"
bi["Categoría de conciliación"] = bi["Cat"].map(CAT_NAME)

# reclasificación CG/2026/07/0139 (recepciones de merchandising abonadas a 71) pertenece a C01
bi.loc[bi.Asiento == "CG/2026/07/0139", "Cat"] = "C01"
bi.loc[bi.Asiento == "CG/2026/07/0139", "Categoría de conciliación"] = CAT_NAME["C01"]

# ---------------------------------------------------------------- contexto por asiento
piv = bi.pivot_table(index="ID asiento", columns="Clase", values="Saldo", aggfunc="sum", fill_value=0)
for c in ["60", "609", "61C", "61K", "2x"]:
    if c not in piv:
        piv[c] = 0.0
# 'otros' sólo es fiable con todas las líneas 60/61/2x del asiento: recalcular con bases completas
full = pd.concat([d60.assign(Clase=np.where(d60.Cuenta.str.startswith("609"), "609", "60"))[["ID asiento", "Clase", "Saldo (D-H)"]],
                  d61.assign(Clase=np.where(d61.Subcuenta.isin(["COMPRA", "ÚNICA"]), "61C", "61K"))[["ID asiento", "Clase", "Saldo (D-H)"]],
                  d2.assign(Clase="2x")[["ID asiento", "Clase", "Saldo (D-H)"]]])
pf = full.pivot_table(index="ID asiento", columns="Clase", values="Saldo (D-H)", aggfunc="sum", fill_value=0)
pf["Otras"] = -pf.sum(axis=1)
pf = pf.round(2)
bi = bi.join(pf.add_prefix("Asiento ").rename(columns={"Asiento Otras": "Asiento otras ctas"}), on="ID asiento")

# ---------------------------------------------------------------- conciliación por clave OC|Producto (C01)
c01 = bi[bi.Cat == "C01"]
k60 = c01[(c01.Clase == "60") & c01["Tipo de operación"].str[:4].isin(["OP01", "OP02"])]
k60i = c01[(c01.Clase == "60") & c01["Tipo de operación"].str[:4].isin(["OP05", "OP16"])]
k61 = c01[(c01.Clase == "61C") & c01["Tipo de operación"].str[:4].isin(["OP04", "OP05"])]
k2 = c01[(c01.Clase == "2x") & c01["Tipo de operación"].str[:4].isin(["OP04", "OP05"])]
agg = lambda d, p: d.groupby("Clave").agg(**{f"{p}": ("Saldo", "sum"), f"Q{p}": ("Cantidad", "sum"),
                                             f"Mes{p}min": ("Mes", "min"), f"Mes{p}max": ("Mes", "max"),
                                             f"Fam{p}": ("Familia", lambda s: " / ".join(sorted(set(s))))})
kt = agg(k60, "F").join(agg(k61, "R"), how="outer").join(agg(k2, "X"), how="outer")
kt[["F", "R", "X", "QF", "QR", "QX"]] = kt[["F", "R", "X", "QF", "QR", "QX"]].fillna(0)
kt["R"] = -kt["R"]  # destino (haber 61) en positivo
kt = kt.join(k60i.groupby("Clave")["Saldo"].sum().rename("I60"))
kt["I60"] = kt["I60"].fillna(0)
kt["D1"] = (kt.F - kt.R + kt.I60).round(2)   # tras reclasificar a 61 lo imputado a 60 (AC-11)
kt["D2"] = (kt.R - kt.I60 - kt.X).round(2)
info = c01[c01.Clave.notna()].groupby("Clave").agg(OC=("OC principal", "first"), Producto=("Producto", "first"),
                                                   Proveedor=("Socio", lambda s: s.dropna().iloc[0] if s.notna().any() else ""))
kt = kt.join(info)
merch_fix = set(kt[(kt.F.abs() > TOL) & (kt.R.abs() <= TOL) & (kt.X.abs() > TOL) & ((kt.F - kt.X).abs() <= TOL)].index)
reclas_0139 = -bi.loc[bi.Asiento == "CG/2026/07/0139", "Saldo"].sum()


def kstatus(r):
    famm = (r.FamF if isinstance(r.FamF, str) else "") and isinstance(r.FamR, str) and r.FamF != r.FamR
    if r.name in merch_fix:
        return ("Conciliado (corregido)", "H11", "Recepción abonada a cta 71 por configuración de categoría; "
                "regularizada con CG/2026/07/0139 (Debe 71111101 / Haber 61211101).")
    if abs(r.D1) <= TOL and abs(r.D2) <= TOL and abs(r.I60) > TOL:
        return ("Error – requiere asiento", "H10", "Devolución/ajuste de almacén imputado a 60 en lugar de 61; cuadra tras la reclasificación AC-11.")
    if abs(r.D1) <= TOL and abs(r.D2) <= TOL:
        if famm:
            return ("Error – requiere asiento", "H15", f"Importes cuadran, pero la factura se registró en {r.FamF} y la recepción en {r.FamR}.")
        if pd.notna(r.MesFmax) and pd.notna(r.MesRmax) and r.MesFmax != r.MesRmax:
            return ("Conciliado", "", "Cuadra en el acumulado; la factura y la recepción caen en meses distintos (desfase temporal resuelto).")
        return ("Conciliado", "", "Factura = recepción = ingreso a existencias.")
    if abs(r.D1) > TOL and abs(r.R) <= TOL and abs(r.X) <= TOL:
        return ("Diferencia temporal (corte)", "H12", "Bien facturado sin recepción en almacén al 30/09/2026 (existencia por recibir).")
    if abs(r.D1) > TOL and abs(r.F) <= TOL:
        if r.D1 > 0:
            return ("Diferencia temporal (corte)", "H13", "Devolución a proveedor sin nota de crédito registrada al 30/09/2026 (NC por recibir).")
        return ("Diferencia temporal (corte)", "H13", "Bien recibido sin factura registrada al 30/09/2026 (factura por recibir).")
    if abs(r.D1) > TOL:
        same_q = abs(r.QF - r.QR) <= max(0.01, 0.001 * abs(r.QR))
        if same_q or (abs(r.R) > TOL and abs(r.D1) <= 0.02 * abs(r.F)):
            return ("Error – requiere asiento", "H14", "Misma cantidad facturada y recibida con distinto valor: diferencia de precio/TC no trasladada al inventario.")
        return ("Diferencia temporal (corte)", "H13" if r.D1 < 0 else "H12",
                "Cantidad facturada ≠ cantidad recibida (entrega/factura parcial o distinta UdM): saldo pendiente de factura o de recepción.")
    if abs(r.D2) > TOL:
        return ("Observación", "H16", "61 = 60 pero el ingreso a 2x difiere (contrapartida de la recepción en otra cuenta).")
    return ("Conciliado", "", "")


st = kt.apply(kstatus, axis=1, result_type="expand")
kt[["Estado", "Código", "Explicación"]] = st
kt = kt.reset_index().rename(columns={"index": "Clave"})
kt = kt.sort_values(["Estado", "D1"], key=lambda s: s if s.name != "D1" else -s.abs()).reset_index(drop=True)
kt["ID Clave"] = np.arange(1, len(kt) + 1)
bi = bi.merge(kt[["Clave", "ID Clave", "Estado", "Código", "Explicación"]].rename(
    columns={"Estado": "kEstado", "Código": "kCódigo", "Explicación": "kExpl"}), on="Clave", how="left")
bi.loc[bi.Cat != "C01", ["ID Clave", "kEstado", "kCódigo", "kExpl"]] = np.nan

# ---------------------------------------------------------------- subcontratación por OC
sb = bi[bi.Cat == "C04"]
sbc_ing61 = sb[(sb["Tipo de operación"].str.startswith("OP08")) & (sb.Clase == "61C")].groupby("OC principal")["Saldo"].sum()
sbc_ing2 = sb[(sb["Tipo de operación"].str.startswith("OP08")) & (sb.Clase == "2x")].groupby("OC principal")["Saldo"].sum()
sbc_comp = sb[(sb["Tipo de operación"].str.startswith("OP09")) & (sb.Clase == "61K")].groupby("OC principal")["Saldo"].sum()
sbc_comp_m = sb[(sb["Tipo de operación"].str.startswith("OP09")) & (sb.Clase == "61K")].groupby("OC principal")["Mes"].max()
sbc_fac = sb[sb.Clase.isin(["60", "609"])].groupby("OC principal")["Saldo"].sum()
sbc_prov = sb.groupby("OC principal")["Socio"].agg(lambda s: s.dropna().iloc[0] if s.notna().any() else "")
sbc_prod = sb[sb["Tipo de operación"].str.startswith("OP08")].groupby("OC principal")["Producto"].agg(lambda s: s.dropna().iloc[0] if s.notna().any() else "")
so = pd.concat([sbc_prov, sbc_prod, sbc_ing2, sbc_ing61, sbc_comp, sbc_fac, sbc_comp_m], axis=1,
               keys=["Subcontratista", "Producto recibido", "Ingreso 2x", "Ingreso 61C", "Componentes 61K", "Factura 60", "Mes último consumo"]).fillna(
    {"Ingreso 2x": 0, "Ingreso 61C": 0, "Componentes 61K": 0, "Factura 60": 0})
so.index.name = "OC"
so = so[(so[["Ingreso 2x", "Ingreso 61C", "Componentes 61K", "Factura 60"]].abs().sum(axis=1) > 0)].reset_index()


def sstat(r):
    serv = r["Ingreso 2x"] - r["Componentes 61K"]
    d = r["Factura 60"] - serv
    if abs(r["Ingreso 2x"]) <= TOL and abs(r["Factura 60"]) > TOL:
        return "Diferencia temporal (corte)", "H05", "Servicio facturado sin ingreso del producto subcontratado al 30/09."
    if abs(r["Factura 60"]) <= TOL and abs(r["Ingreso 2x"]) > TOL:
        return "Diferencia temporal (corte)", "H05", "Producto subcontratado ingresado sin factura del servicio (servicio por facturar)."
    if abs(d) <= max(TOL, 0.02 * abs(r["Factura 60"])):
        return "Conciliado", "", "Servicio facturado = valor ingresado − componentes enviados."
    return "Observación", "H05", "Servicio facturado ≠ (valor ingresado − componentes): revisar costo de componentes o precio del servicio."


so[["Estado", "Código", "Explicación"]] = so.apply(sstat, axis=1, result_type="expand")

# pendiente de reclasificar SBC por familia
sbc_cons_fam = bi[(bi["Tipo de operación"].str.startswith("OP09")) & (bi.Clase == "61K")].groupby("Familia")["Saldo"].sum()
rc = bi[(bi["Tipo de operación"].str.startswith("OP12"))]
rc_cons_fam = -rc[rc.Clase == "61K"].groupby("Familia")["Saldo"].sum()
rc_compra_fam = rc[rc.Clase == "61C"].groupby("Familia")["Saldo"].sum()
sbc_pend = pd.concat([sbc_cons_fam, rc_cons_fam, rc_compra_fam], axis=1, keys=["Consumo SBC", "Reclasificado (Haber 61K)", "Reclasificado a 61C"]).fillna(0)
sbc_pend["Reclasificado a 71"] = sbc_pend["Reclasificado (Haber 61K)"] - sbc_pend["Reclasificado a 61C"]
sbc_pend["Pendiente"] = (sbc_pend["Consumo SBC"] - sbc_pend["Reclasificado (Haber 61K)"]).round(2)
sbc_pend["% a 61C"] = np.where(sbc_pend["Reclasificado (Haber 61K)"].abs() > 0, sbc_pend["Reclasificado a 61C"] / sbc_pend["Reclasificado (Haber 61K)"], 1.0)
sbc_month = bi[(bi["Tipo de operación"].str.startswith("OP09")) & (bi.Clase == "61K")].pivot_table(index="Mes", columns="Familia", values="Saldo", aggfunc="sum", fill_value=0)

# ---------------------------------------------------------------- 609 vs LC por OC de servicio
v = bi[bi.Cat == "C03"]
inv609 = v[v.Clase.isin(["609", "60"]) & v["Tipo de operación"].str[:4].isin(["OP03", "OP02"])].groupby("OC principal", dropna=False)["Saldo"].sum()
lc = v[v["Tipo de operación"].str.startswith("OP06")]
lc61 = lc[lc.Clase.isin(["61C", "61K"])].groupby("OC secundaria")["Saldo"].sum()
lc609 = lc[lc.Clase.isin(["609", "60"])].groupby("OC secundaria")["Saldo"].sum()
lc2 = lc[lc.Clase == "2x"].groupby("OC secundaria")["Saldo"].sum()
prov609 = v[v.Clase.isin(["609", "60"])].groupby("OC principal", dropna=False)["Socio"].agg(lambda s: s.dropna().iloc[0] if s.notna().any() else "")
glosa609 = v[v.Clase.isin(["609", "60"])].groupby("OC principal", dropna=False)["Glosa"].agg(lambda s: str(s.dropna().iloc[0])[:80] if s.notna().any() else "")
cta609 = v[v.Clase.isin(["609", "60"])].groupby("OC principal", dropna=False)["Cuenta"].agg(lambda s: " / ".join(sorted(set(s))))
mes609 = v[v.Clase.isin(["609", "60"])].groupby("OC principal", dropna=False)["Mes"].max()
for _s in (prov609, glosa609, cta609, mes609):
    _s.index = _s.index.fillna("(sin OC)") if hasattr(_s.index, "fillna") else _s.index
inv609.index = inv609.index.fillna("(sin OC)")
lt = pd.concat([prov609, glosa609, cta609, mes609, inv609, lc61, lc609, lc2], axis=1,
               keys=["Proveedor", "Glosa", "Cuenta 60/609", "Mes", "Factura 609", "LC abonado a 61", "LC abonado a 609", "LC cargado a 2x"]).fillna(
    {"Factura 609": 0, "LC abonado a 61": 0, "LC abonado a 609": 0, "LC cargado a 2x": 0})
lt.index.name = "OC servicio"
lt = lt.reset_index()
lt["OC servicio"] = lt["OC servicio"].fillna("(sin OC)")


def lstat(r):
    cap = -(r["LC abonado a 61"] + r["LC abonado a 609"])
    if r["OC servicio"] == "(sin OC)" and abs(r["Factura 609"]) > TOL:
        return "Observación", "H08", "Factura/NC 609 sin OC: no puede vincularse a un costo en destino."
    if abs(r["Factura 609"]) <= TOL and abs(r["LC cargado a 2x"]) > TOL:
        return "Observación", "H07", "Costo en destino capitalizado sin factura 609 vinculada (abonado a 61/otras cuentas)."
    if abs(r["LC abonado a 609"]) > TOL:
        return "Error – requiere asiento", "H06", "El costo en destino se abonó a la 609 en lugar de la 61 (destino)."
    if abs(cap) <= TOL and abs(r["Factura 609"]) > TOL:
        return "Error – requiere asiento", "H08", "Gasto vinculado (609) sin costo en destino: no se capitalizó al inventario."
    if abs(r["Factura 609"] - cap) <= max(TOL, 0.01 * abs(r["Factura 609"])):
        return "Conciliado", "", "Factura 609 = costo en destino capitalizado."
    if r["Factura 609"] > cap:
        return "Observación", "H08", "Capitalización parcial: parte del 609 no fue trasladada al inventario (LC parcial o porción de unidades ya consumidas)."
    return "Observación", "H07", "Costo en destino mayor a la factura 609."


lt[["Estado", "Código", "Explicación"]] = lt.apply(lstat, axis=1, result_type="expand")
lt = lt[(lt[["Factura 609", "LC abonado a 61", "LC abonado a 609", "LC cargado a 2x"]].abs().sum(axis=1) > 0)]

# ---------------------------------------------------------------- códigos por línea
HALL = {}


def setcode(mask, code, estado, expl):
    bi.loc[mask, ["Código hallazgo", "Estado conciliación", "Explicación"]] = [code, estado, expl]


bi["Código hallazgo"] = ""
bi["Estado conciliación"] = "Conciliado"
bi["Explicación"] = ""
op = bi["Tipo de operación"].str[:4]

# C01 – desde la clave
m = bi.Cat.eq("C01") & bi["ID Clave"].notna()
bi.loc[m, "Estado conciliación"] = bi.loc[m, "kEstado"]
bi.loc[m, "Código hallazgo"] = bi.loc[m, "kCódigo"].fillna("")
bi.loc[m, "Explicación"] = bi.loc[m, "kExpl"]
setcode(bi.Cat.eq("C01") & bi["ID Clave"].isna(), "H09", "Observación",
        "Devolución/retorno a proveedor sin OC vinculada: verificar la NC del proveedor (no se puede conciliar por OC).")
setcode(bi.Asiento.eq("CG/2026/07/0139"), "H11", "Conciliado (corregido)",
        "Reclasificación manual existente: recepciones de merchandising abonadas a 71 trasladadas a 61211101.")
# devoluciones imputadas a 60
setcode((bi.Fuente == "60") & op.eq("OP05"), "H10", "Error – requiere asiento",
        "Devolución a proveedor (movimiento de almacén) imputada a cuenta 60: el ajuste de valor debe ir a 61 (destino).")
setcode((bi.Fuente == "60") & op.eq("OP16"), "H10", "Error – requiere asiento",
        "Revaloración manual de costo imputada a cuenta 60 (no es compra): debe ir a 61.")
# C02
setcode(bi.Cat.eq("C02") & (bi.Clase == "60"), "H09", "Observación",
        "Factura/NC de compra sin OC: no tiene recepción vinculada en Odoo; no participa del destino automático 61/2x.")
# C03
lcmap = lt.set_index("OC servicio")[["Estado", "Código", "Explicación"]]
m = bi.Cat.eq("C03") & bi.Clase.isin(["609", "60"]) & op.eq("OP03")
k = bi.loc[m, "OC principal"].fillna("(sin OC)")
bi.loc[m, ["Estado conciliación", "Código hallazgo", "Explicación"]] = lcmap.reindex(k).values
m = bi.Cat.eq("C03") & op.eq("OP06")
k = bi.loc[m, "OC secundaria"].fillna("(sin OC)")
vals = lcmap.reindex(k)
vals = vals.fillna({"Estado": "Observación", "Código": "H07", "Explicación": "Costo en destino sin OC de servicio vinculada."})
bi.loc[m, ["Estado conciliación", "Código hallazgo", "Explicación"]] = vals.values
setcode(bi.Cat.eq("C03") & op.eq("OP06") & bi.Clase.isin(["609", "60"]), "H06", "Error – requiere asiento",
        "Costo en destino abonado a la 609 (gasto vinculado) en lugar de 61: la 609 queda anulada y no refleja el costo de compra.")
setcode(bi.Cat.eq("C03") & op.eq("OP06") & bi["Asiento 61C"].abs().le(TOL) & bi["Asiento 609"].abs().le(TOL) & bi["Asiento 60"].abs().le(TOL),
        "H07", "Observación",
        "Costo en destino (fletes / costos indirectos de cacao) capitalizado en 2x con contrapartida fuera de 60/61: no pasa por compras.")
# C04
m = bi.Cat.eq("C04")
setcode(m & op.eq("OP09") & bi.Mes.ge(6), "H04", "Error – requiere asiento",
        "Consumo de componentes enviados a subcontratación no reclasificado de 61-Consumo a 61-Compra/71 (reclasificación sólo cubrió Ene-May).")
setcode(m & op.eq("OP09") & bi.Mes.lt(6), "", "Conciliado (corregido)",
        "Consumo SBC Ene-May revertido con CG/2026/07/0097 (ver residuo por familia en hoja Conc_Subcontratacion).")
setcode(m & op.eq("OP12"), "", "Conciliado (corregido)", "Reclasificación manual CG/2026/07/0097 de consumos SBC Ene-May.")
somap = so.set_index("OC")[["Estado", "Código", "Explicación"]]
for o in ("OP07", "OP08"):
    mm = m & op.eq(o) & bi["OC principal"].isin(somap.index)
    bi.loc[mm, ["Estado conciliación", "Código hallazgo", "Explicación"]] = somap.reindex(bi.loc[mm, "OC principal"]).values
setcode(m & op.eq("OP07") & bi["OC principal"].isna(), "H05", "Observación",
        "Factura de servicio de subcontratación (armado/selección) sin OC: no se vincula con el ingreso del producto.")
setcode(m & op.eq("OP08") & bi.Clase.eq("61C") & ~bi["OC principal"].isin(somap.index), "H05", "Observación",
        "Ingreso de producto subcontratado sin OC vinculada.")
# C05
setcode(bi.Cat.eq("C05") & op.eq("OP10") & bi.Clase.eq("61C"), "H03", "Error – requiere asiento",
        "Ingreso de producción propia abonado a 61-Compra: la producción no es compra, debe abonarse a 71 (producción almacenada).")
setcode(bi.Cat.eq("C05") & op.eq("OP11") & bi.Clase.isin(["61C", "61K"]), "H02", "Error – requiere asiento",
        "Desmontaje: componentes devueltos abonados a 61-Compra; la reclasificación CG/2026/07/0143 llevó también el consumo a 61-Compra. Neto debe ir a 61-Consumo.")
# C06
setcode(bi.Cat.eq("C06") & op.eq("OP17") & bi.Clase.eq("61C"), "H01", "Conciliado (corregido)",
        "Factura de gasto/servicio (alquiler, comisión, movilidad) imputada a 61-Compra; ya reclasificada con CG/2026/07/0101-0102.")
setcode(bi.Cat.eq("C06") & op.eq("OP17") & bi.Clase.eq("61K"), "H01", "Error – requiere asiento",
        "Factura de gasto/activo menor imputada a 61-Consumo y no reclasificada.")
done_rec = bi.Asiento.isin(["CG/2026/07/0101", "CG/2026/07/0102"])
setcode(bi.Cat.eq("C06") & done_rec, "H01", "Conciliado (corregido)", "Reclasificación manual existente de gastos imputados a 61.")
setcode(bi.Cat.eq("C06") & op.eq("OP17") & bi.Clase.eq("2x"), "H17", "Error – requiere asiento",
        "Factura de compra (repuestos) cargada directo a 2x sin pasar por 60/61: la compra no figura en la 60.")
# C07
setcode(bi.Cat.eq("C07") & bi.Clase.isin(["61C"]) & bi.Saldo.abs().gt(0.005), "H18", "Observación",
        "Ajuste de inventario / revaloración registrado en 61-Compra: no es compra; debería ir a 61-Consumo o 69/65/75.")
# C08
setcode(bi.Cat.eq("C08"), "H19", "Error – requiere asiento",
        "Provisión mensual de 'compras no ingresadas' calculada como cuadre (60 − 61) por familia, con contrapartida en 2x: no representa existencias por recibir reales.")
setcode(bi.Asiento.eq("CG/2026/07/0207"), "H20", "Error – requiere asiento",
        "Provisión de julio no extornada en agosto (rompe la mecánica provisión/extorno); su efecto permanece en 61 y 2x.")
# 61 vs 2x familia distinta
fm = d61[d61["¿Coincide familia?"] == "No"]["ID asiento"].unique()
mm = bi["ID asiento"].isin(fm) & bi.Clase.isin(["61C"]) & op.isin(["OP06"])
# sólo marcar si la familia del asiento 2x difiere de la del 61
bi.loc[bi.Cat.isna(), "Cat"] = "C07"
bi["Categoría de conciliación"] = bi["Cat"].map(CAT_NAME)
# líneas 2x/61K en asientos de C01 que son contrapartida: estado ya heredado
bi["Estado conciliación"] = bi["Estado conciliación"].fillna("Conciliado")
bi["Código hallazgo"] = bi["Código hallazgo"].fillna("")
bi["Explicación"] = bi["Explicación"].fillna("")

# ---------------------------------------------------------------- familia 61 vs 2x por asiento (LC)
lcfam = d61[(d61["¿Coincide familia?"] == "No")].copy()
_lk = lcfam[lcfam.Motivo == "Costo en destino (Landed cost)"].assign(Cuenta=lambda x: x["Cuenta"].astype(str), Saldo=lambda x: x["Saldo (D-H)"].round(2))
_lk = _lk[_lk["Cta contrapartida principal"].astype(str).str[:2].isin(["20", "24", "25", "26"])][["ID asiento", "Cuenta", "Saldo"]].drop_duplicates()
_lk["_m"] = 1
bi = bi.merge(_lk, on=["ID asiento", "Cuenta", "Saldo"], how="left")
setcode(bi["_m"].eq(1) & bi.Clase.eq("61C"), "H16", "Error – requiere asiento",
        "Costo en destino abonado a una 61 de familia distinta a la de la existencia que recibe el costo.")
bi = bi.drop(columns="_m")


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


bi["Familia conc."] = bi["Cuenta"].map(fam_conc)

# ---------------------------------------------------------------- guardar
bi = bi.sort_values(["Cat", "Mes", "ID asiento", "Fuente"]).reset_index(drop=True)
bi["ID línea"] = np.arange(1, len(bi) + 1)
out = dict(bi=bi, kt=kt, so=so, lt=lt, sbc_pend=sbc_pend, sbc_month=sbc_month, lcfam=lcfam, CAT_NAME=CAT_NAME,
           reclas_0139=reclas_0139)
pickle.dump(out, open(f"{W}/model.pkl", "wb"))

# ---------------------------------------------------------------- diagnóstico
pd.set_option("display.width", 250); pd.set_option("display.max_rows", 200); pd.set_option("display.max_columns", 30)
print(len(bi), "líneas en base integrada")
print(bi.pivot_table(index="Categoría de conciliación", columns="Clase", values="Saldo", aggfunc="sum", fill_value=0).round(0))
print(bi.groupby(["Estado conciliación", "Código hallazgo"])["Saldo"].agg(["count", "sum"]).round(0))
print(kt.groupby(["Estado", "Código"])[["F", "R", "X", "D1", "D2"]].agg(["count", "sum"]).round(0))
print(so.groupby(["Estado"])[["Ingreso 2x", "Componentes 61K", "Factura 60"]].agg(["count", "sum"]).round(0))
print(sbc_pend.round(0))
print(lt.groupby(["Estado", "Código"])[["Factura 609", "LC abonado a 61", "LC abonado a 609", "LC cargado a 2x"]].agg(["count", "sum"]).round(0))
