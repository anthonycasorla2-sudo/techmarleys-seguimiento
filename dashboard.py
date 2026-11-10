"""Genera dashboard_datos.js con todo lo que muestra dashboard.html.

Uso:
    python dashboard.py
"""

import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
DB = CARPETA / "techmarleys.db"
SALIDA = CARPETA / "dashboard_datos.js"
TIENDA = "https://techmarleys.com/"


def generar(db=DB, salida=SALIDA, excels=None):
    """excels: {"AAAA-MM": "ruta/relativa.xlsx"} para mostrar el enlace al Excel del mes."""
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    columnas = {r["name"] for r in con.execute("PRAGMA table_info(productos)")}
    con_valor = "valor" in {r["name"] for r in con.execute("PRAGMA table_info(capturas)")}

    productos = [{
        "id": r["id"], "nombre": r["nombre"], "marca": r["marca"], "modelo": r["modelo"],
        "condicion": r["condicion"], "categoria": r["categoria"], "subcategoria": r["subcategoria"],
        "sucursal": r["sucursal"], "stock": r["stock"] or 0, "precio": r["precio_final"] or 0,
        "foto": TIENDA + r["foto"] if "foto" in columnas and r["foto"] else "",
        "activo": r["activo"], "desde": r["primera_vez"],
    } for r in con.execute("SELECT * FROM productos")]

    eventos = [dict(r) for r in con.execute("""
        SELECT fecha_hora, producto_id AS id, nombre, tipo, stock_antes, stock_despues,
               unidades, precio_antes, precio, monto
        FROM eventos ORDER BY fecha_hora, id""")]

    # Unidades y valor del inventario en cada captura, para la evolución del stock.
    valor = "c.valor" if con_valor else "NULL"
    capturas = [list(r) for r in con.execute(f"""
        SELECT c.fecha_hora, c.total_unidades, c.total_productos,
               COALESCE({valor}, (SELECT ROUND(SUM(h.stock * h.precio_final), 2)
                                  FROM stock_historial h WHERE h.captura_id = c.id))
        FROM capturas c WHERE c.ok = 1 ORDER BY c.fecha_hora""")]

    ultima = con.execute("SELECT fecha_hora, ok, error FROM capturas ORDER BY id DESC LIMIT 1").fetchone()
    con.close()

    datos = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tienda": TIENDA,
        "ultima_captura": dict(ultima) if ultima else None,
        "productos": productos,
        "eventos": eventos,
        "capturas": capturas,
        "excels": excels or {},
    }
    # Se escribe a un archivo temporal y luego se reemplaza, para que el
    # dashboard nunca lea un archivo a medio escribir.
    temporal = salida.with_suffix(".tmp")
    temporal.write_text("window.DATOS = " + json.dumps(datos, ensure_ascii=False, separators=(",", ":")) + ";\n",
                        encoding="utf-8")
    os.replace(temporal, salida)
    return salida


if __name__ == "__main__":
    print(generar())
