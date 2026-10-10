"""Convierte las hojas del .xlsb de Odoo a pickles de pandas. Uso: python prep.py <archivo.xlsb> <carpeta_trabajo>"""
import sys, pandas as pd
from pyxlsb import open_workbook

src, W = sys.argv[1], sys.argv[2]
with open_workbook(src) as wb:
    for name in wb.sheets:
        with wb.get_sheet(name) as sh:
            rows = [[c.v for c in r] for r in sh.rows()]
        n = max(len(r) for r in rows)
        hdr = rows[0] + [None] * (n - len(rows[0]))
        hdr = [h if h not in (None, "") else f"col{i}" for i, h in enumerate(hdr)]
        df = pd.DataFrame([r + [None] * (n - len(r)) for r in rows[1:]], columns=hdr).dropna(axis=1, how="all")
        df.to_pickle(f"{W}/{name}.pkl")
