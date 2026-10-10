"""Asientos de corrección propuestos + puentes antes/después."""
import pandas as pd, numpy as np, pickle, sys

W = sys.argv[1]
M = pickle.load(open(f"{W}/model.pkl", "rb"))
bi, kt, so, lt, sbc_pend = M["bi"], M["kt"], M["so"], M["lt"], M["sbc_pend"]
d61 = pd.read_pickle(f"{W}/Detalle_61.pkl")
d2 = pd.read_pickle(f"{W}/Base_Cta2.pkl")
TOL = 1.0
FECHA = pd.Timestamp("2026-09-30")

ACC = {"Mercaderías": dict(c60="60111101", c61="61211101", k61="61211101", c2="20111101", c28="28111101"),
       "Materias primas": dict(c60="60211101", c61="61231102", k61="61231101", c2="24111101", c28="28411101"),
       "Suministros": dict(c60="60321101", c61="61321102", k61="61321101", c2="25241101", c28="28511101"),
       "Repuestos": dict(c60="60331101", c61="61331102", k61="61331102", c2="25311102", c28="28511101"),
       "Envases": dict(c60="60411101", c61="61411102", k61="61411101", c2="26111101", c28="28611101"),
       "Embalajes": dict(c60="60421101", c61="61421102", k61="61421101", c2="26211101", c28="28611101")}
NAMES = {"60111101": "Compras – Mercaderías", "60211101": "Compras – Materias primas", "60321101": "Compras – Suministros",
         "60331101": "Compras – Repuestos (cuenta sugerida)", "60411101": "Compras – Envases", "60421101": "Compras – Embalajes",
         "60911101": "Costos vinculados – Mercaderías (transporte)", "60919101": "Costos vinculados – Mercaderías (otros)",
         "60921101": "Costos vinculados – Materias primas (transporte)", "60925102": "Costos vinculados – Materias primas (otros)",
         "60931101": "Costos vinculados – Materiales auxiliares/suministros", "60941101": "Costos vinculados – Envases y embalajes (transporte)",
         "60945101": "Costos vinculados – Envases y embalajes (otros)",
         "61211101": "Variación de existencias – Mercaderías", "61231101": "Variación MP – Consumo",
         "61231102": "Variación MP – Compra (destino)", "61321101": "Variación suministros – Consumo",
         "61321102": "Variación suministros – Compra (destino)", "61331102": "Variación repuestos – Compra (destino)",
         "61411101": "Variación envases – Consumo", "61411102": "Variación envases – Compra (destino)",
         "61421101": "Variación embalajes – Consumo", "61421102": "Variación embalajes – Compra (destino)",
         "20111101": "Mercaderías", "20911101": "Mercaderías en tránsito", "24111101": "Materias primas",
         "25241101": "Suministros – otros", "25311102": "Repuestos", "26111101": "Envases",
         "26211101": "Embalajes – envolturas y etiquetas", "28111101": "Mercaderías por recibir (cuenta sugerida)",
         "28411101": "Materias primas por recibir (cuenta sugerida)", "28511101": "Suministros/repuestos por recibir (cuenta sugerida)",
         "28611101": "Envases y embalajes por recibir (cuenta sugerida)", "42111101": "Facturas por recibir – proveedores (cuenta sugerida)",
         "71311102": "Variación producción almacenada – productos en proceso", "65xxxxxx": "Gasto por naturaleza (definir: 65/63)"}


def clase_of(acc):
    if acc.startswith("609"):
        return "609"
    if acc.startswith("60"):
        return "60"
    if acc.startswith("61"):
        return "61K" if acc in ("61231101", "61321101", "61411101", "61421101") else "61C"
    if acc[:1] == "2":
        return "2x"
    return "Otras"


rows = []


def add(ac, hall, cat, acc, debe, haber, glosa, ref, fam="", nota="", mes=9):
    if round(debe, 2) == 0 and round(haber, 2) == 0:
        return
    rows.append(dict(AC=ac, Fecha=FECHA, Hallazgo=hall, Cat=cat, Cuenta=acc, Nombre=NAMES.get(acc, ""), Clase=clase_of(acc),
                     Familia=fam, Debe=round(debe, 2), Haber=round(haber, 2), Glosa=glosa, Referencia=ref, Nota=nota))


def pair(ac, hall, cat, acc_d, acc_h, amt, glosa, ref, fam="", nota=""):
    """amt>0: Debe acc_d / Haber acc_h ; amt<0 invierte."""
    amt = round(amt, 2)
    if abs(amt) < 0.005:
        return
    if amt < 0:
        acc_d, acc_h, amt = acc_h, acc_d, -amt
    add(ac, hall, cat, acc_d, amt, 0, glosa, ref, fam, nota)
    add(ac, hall, cat, acc_h, 0, amt, glosa, ref, fam, nota)


# ---------------- AC-01 Reversión provisiones de cuadre (Sep + Jul no extornada)
for asi, hall in [("CG/2026/09/0195", "H19"), ("CG/2026/07/0207", "H20")]:
    l61 = d61[d61.Asiento == asi]
    l2 = d2[d2.Asiento == asi]
    for _, r in pd.concat([l61.rename(columns={"Cuenta 61": "Cuenta"}), l2]).iterrows():
        acc = str(r.Cuenta).replace(".0", "")
        add("AC-01", hall, "C08", acc, r.Haber, r.Debe, f"Reversión provisión de cuadre {asi} (no representa compras no ingresadas)", asi,
            "", "Sustituir por el corte real de AC-02/AC-03.")

# ---------------- key-level
c01 = bi[(bi.Cat == "C01") & bi["ID Clave"].notna()]


def main_acc(df, cls):
    x = df[df.Clase == cls]
    if x.empty:
        return None
    return x.groupby("Cuenta")["Saldo"].sum().abs().idxmax()


kinfo = {}
for idc, g in c01.groupby("ID Clave"):
    kinfo[idc] = dict(a60=main_acc(g, "60"), a61=main_acc(g, "61C"), a2=main_acc(g, "2x"))
FAM_OF = {}
for f, d in ACC.items():
    for v in d.values():
        FAM_OF.setdefault(v, f)
FAM_OF.update({"20911101": "Mercaderías"})

for _, k in kt.iterrows():
    inf = kinfo.get(k["ID Clave"], {})
    a60, a61, a2 = inf.get("a60"), inf.get("a61"), inf.get("a2")
    fam = FAM_OF.get(a60 or a61 or a2 or "", "")
    fa = ACC.get(fam, ACC["Materias primas"])
    ref = f"{k.OC} | {k.Producto}"
    if k.Código == "H12":
        pair("AC-02", "H12", "C01", fa["c28"], a61 or fa["c61"], k.D1,
             "Corte: bien facturado no recibido al 30/09 → existencias por recibir", ref, fam,
             "Extornar en octubre o al registrar la recepción en Odoo.")
    elif k.Código == "H13":
        pair("AC-03", "H13", "C01", a60 or fa["c60"], "42111101", -k.D1,
             "Corte: bien recibido no facturado al 30/09 → factura por recibir", ref, fam,
             "Extornar al registrar la factura del proveedor.")
    elif k.Código == "H14":
        pair("AC-04", "H14", "C01", a2 or fa["c2"], a61 or fa["c61"], k.D1,
             "Diferencia de precio/TC factura vs valorización de la recepción", ref, fam,
             "Si el bien ya fue consumido, cargar 61-Consumo en lugar de 2x.")
    elif k.Código == "H15":
        f60 = FAM_OF.get(a60, ""); f61 = FAM_OF.get(a61, "")
        amt = k.F
        pair("AC-05", "H15", "C01", ACC[f61]["c60"], a60, amt,
             f"Reclasificación de familia: factura en {f60}, recepción en {f61}", ref, f61,
             "Corregir la categoría del producto / cuenta de compra en Odoo.")

# ---------------- AC-06 SBC pendiente
for fam, r in sbc_pend.iterrows():
    pend = r["Pendiente"]
    if abs(pend) < TOL:
        continue
    pct = r["Reclasificado a 61C"] / r["Reclasificado (Haber 61K)"] if r["Reclasificado (Haber 61K)"] else 1.0
    fa = ACC[fam]
    a61c = round(pend * pct, 2)
    a71 = round(pend - a61c, 2)
    add("AC-06", "H04", "C04", fa["c61"], a61c, 0, "Reversión de consumos de componentes enviados a subcontratación pendiente (Jun-Set y residuo Ene-May)",
        "1101/SBC – Consumo MP enviada a subcontratación", fam, "Misma lógica que CG/2026/07/0097.")
    if a71:
        add("AC-06", "H04", "C04", "71311102", a71, 0, "Parte de componentes SBC de productos en proceso (misma proporción que CG/2026/07/0097)",
            "1101/SBC", fam)
    add("AC-06", "H04", "C04", fa["k61"], 0, pend, "Reversión de consumos SBC pendiente", "1101/SBC", fam)

# ---------------- AC-07 Desmontajes: neto de 61C a 61K
ub = bi[(bi["Tipo de operación"].str.startswith("OP11")) & (bi.Clase == "61C")]
for fam, amt in ub.groupby("Familia")["Saldo"].sum().items():
    if abs(amt) < TOL or fam == "Mercaderías":
        continue
    fa = ACC[fam]
    # saldo deudor en 61C (amt>0) → Haber 61C / Debe 61K
    pair("AC-07", "H02", "C05", fa["k61"], fa["c61"], amt,
         "Desmontajes (Unbuild): trasladar efecto neto de 61-Compra a 61-Consumo", "UB/ + CG/2026/07/0143", fam,
         "El desmontaje es transformación interna, no compra.")

# ---------------- AC-08 Producción propia en 61C → 71
pp = bi[(bi["Tipo de operación"].str.startswith("OP10")) & (bi.Clase == "61C")]
for fam, amt in pp.groupby("Familia")["Saldo"].sum().items():
    fa = ACC.get(fam)
    if fa is None or abs(amt) < TOL:
        continue
    pair("AC-08", "H03", "C05", fa["c61"], "71311102", -amt,
         "Ingreso de producción propia abonado a 61-Compra → producción almacenada (71)", "Órdenes de fabricación (OP)", fam,
         "Revisar cuenta de producción de la categoría en Odoo.")

# ---------------- AC-09 LC abonado a 609 → 61
lc609 = bi[(bi["Tipo de operación"].str.startswith("OP06")) & (bi.Clase.isin(["609", "60"]))]
for (acc, oc), g in lc609.groupby(["Cuenta", "OC secundaria"]):
    amt = g.Saldo.sum()
    # familia del bien que recibió el costo: por 2x del asiento
    ids = g["ID asiento"].unique()
    x = bi[bi["ID asiento"].isin(ids) & (bi.Clase == "2x")]
    fam = FAM_OF.get(x.groupby("Cuenta")["Saldo"].sum().abs().idxmax(), "Envases") if not x.empty else "Envases"
    pair("AC-09", "H06", "C03", acc, ACC[fam]["c61"], -amt,
         "Costo en destino abonado a 609: trasladar el abono a 61 (destino)", f"{oc} | {', '.join(g.Asiento.unique()[:3])}", fam,
         "Configurar la cuenta de abono del LC = cuenta puente/61, no la 609.")

# ---------------- AC-10 609 sin capitalizar
F609 = {"6091": "Mercaderías", "6092": "Materias primas", "6093": "Suministros", "6094": "Envases"}
err = lt[(lt.Código == "H08") & (lt.Estado == "Error – requiere asiento")]
for _, r in err.iterrows():
    fam = F609.get(str(r["Cuenta 60/609"])[:4], "Envases")
    fa = ACC[fam]
    pair("AC-10", "H08", "C03", fa["c2"], fa["c61"], r["Factura 609"],
         "Gasto vinculado (609) no capitalizado: destino a existencias", f"{r['OC servicio']} | {r['Proveedor']}", fam,
         "Preferible: crear el Costo en destino en Odoo. Si el bien ya fue consumido, cargar 61-Consumo. Si no es costo de compra, reclasificar a 63.")

# ---------------- AC-11 factura cargada directo a 2x
f2 = bi[(bi["Tipo de operación"].str.startswith("OP17")) & (bi.Clase == "2x")]
for asi, g in f2.groupby("Asiento"):
    amt = g.Saldo.sum()
    pair("AC-11", "H17", "C06", "60331101", "61331102", amt,
         "Compra de repuestos cargada directo a 2x: registrar compra (60) y su destino (61)", f"{asi} | {g.Socio.iloc[0]}", "Repuestos",
         "Asignar a la categoría del producto la cuenta de compra 60 y usar recepción de almacén.")

# ---------------- AC-12 devoluciones / revaloraciones en 60
x60 = bi[(bi.Fuente == "60") & (bi["Código hallazgo"] == "H10")]
for (acc, mes, top), g in x60.groupby(["Cuenta", "Mes", "Tipo de operación"]):
    fam = FAM_OF.get(acc, "")
    amt = g.Saldo.sum()
    pair("AC-12", "H10", "C07" if top.startswith("OP16") else "C01", ACC[fam]["k61" if top.startswith("OP16") else "c61"], acc, amt,
         f"Movimiento de inventario (devolución/revaloración) imputado a 60 – mes {mes}: trasladar a 61", f"{len(g)} líneas – mes {mes}", fam,
         "Revisar cuentas de valorización de la categoría (entrada/salida de stock).")

# ---------------- AC-13 gastos en 61 pendientes
h01 = bi[(bi["Código hallazgo"] == "H01") & (bi["Estado conciliación"] == "Error – requiere asiento") & bi.Clase.isin(["61C", "61K"])]
h01 = h01[~h01["ID asiento"].isin(bi[(bi.Asiento.isin(["CG/2026/07/0102", "CG/2026/07/0101"]))]["ID asiento"])]
done = d61[d61.Asiento.isin(["CG/2026/07/0101", "CG/2026/07/0102"])]["Glosa línea"].str.upper().tolist()
for _, r in h01.iterrows():
    g = str(r.Glosa).upper()
    if any(g[-30:] in dd for dd in done) or r.Clase == "61C":
        continue
    pair("AC-13", "H01", "C06", "65xxxxxx", r.Cuenta, r.Saldo, f"Gasto imputado a 61: {str(r.Glosa)[:60]}", f"{r.Asiento} | {r.Socio}", "",
         "Cuenta de gasto según naturaleza (activo menor/suministro).")

# ---------------- AC-14 familia 61 ≠ familia 2x en LC
lf = M["lcfam"]
lf = lf[lf.Motivo == "Costo en destino (Landed cost)"]
MAP2 = {"20": "Mercaderías", "24": "Materias primas", "252": "Suministros", "261": "Envases", "262": "Embalajes"}
for _, r in lf.iterrows():
    a2 = str(r["Cta contrapartida principal"]).replace(".0", "")
    fam2 = next((v for k, v in MAP2.items() if a2.startswith(k)), None)
    if fam2 is None:
        continue
    acc = str(r["Cuenta"]).replace(".0", "")
    pair("AC-14", "H16", "C03", acc, ACC[fam2]["c61"], -r["Saldo (D-H)"],
         f"Costo en destino: 61 de familia distinta a la existencia ({a2})", f"{r.Asiento} | {r.Producto}", fam2)

# ---------------- AC-15 ajustes en 61C
h18 = bi[bi["Código hallazgo"] == "H18"]
for (acc, fam), g in h18.groupby(["Cuenta", "Familia"]):
    if fam not in ACC:
        continue
    kacc = "61321101" if fam == "Repuestos" else ACC[fam]["k61"]
    pair("AC-15", "H18", "C07", kacc, acc, g.Saldo.sum(),
         "Ajustes de inventario/revaloración registrados en 61-Compra → 61-Consumo", f"{len(g)} líneas", fam)

ac = pd.DataFrame(rows)
FAM_OF.update({"60911101": "Gastos vinculados (609)", "60919101": "Gastos vinculados (609)", "60921101": "Gastos vinculados (609)",
               "60925102": "Gastos vinculados (609)", "60931101": "Gastos vinculados (609)", "60941101": "Gastos vinculados (609)",
               "60945101": "Gastos vinculados (609)", "61231101": "Materias primas", "61321101": "Suministros",
               "61411101": "Envases", "61421101": "Embalajes", "60331101": "Repuestos", "61331102": "Repuestos"})
ac["Familia"] = [FAM_OF.get(a, f) if (not f or a.startswith("609") or a.startswith("6")) else f for a, f in zip(ac.Cuenta, ac.Familia)]
ac["Saldo"] = ac.Debe - ac.Haber
ac["Línea"] = np.arange(1, len(ac) + 1)

# ---------------- puentes
bridge = bi.pivot_table(index="Cat", columns="Clase", values="Saldo", aggfunc="sum", fill_value=0)
adj = ac.pivot_table(index="Cat", columns="Clase", values="Saldo", aggfunc="sum", fill_value=0)
for df in (bridge, adj):
    for c in ["60", "609", "61C", "61K", "2x", "Otras"]:
        if c not in df:
            df[c] = 0.0
after = bridge.add(adj, fill_value=0)


def show(df, t):
    d = df.copy()
    d["Compras 60+609"] = d["60"] + d["609"]
    d["Destino 61C"] = -d["61C"]
    d["Dif 60-61"] = d["Compras 60+609"] - d["Destino 61C"]
    d["Dif 61-2x"] = d["Destino 61C"] - d["2x"]
    print(t); print(d[["Compras 60+609", "Destino 61C", "2x", "Dif 60-61", "Dif 61-2x"]].round(0)); print(d[["Compras 60+609", "Destino 61C", "2x", "Dif 60-61", "Dif 61-2x"]].sum().round(0))


pd.set_option("display.width", 250)
show(bridge, "ANTES"); show(after, "DESPUÉS")
chk = ac.groupby("AC")[["Debe", "Haber"]].sum().round(2)
print(chk.assign(d=chk.Debe - chk.Haber))
print(ac.groupby(["AC", "Hallazgo"]).size())
M["ac"] = ac
pickle.dump(M, open(f"{W}/model.pkl", "wb"))
