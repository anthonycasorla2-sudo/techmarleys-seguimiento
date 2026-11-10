"""Rastreador de inventario de techmarleys.com.

Cada ejecución descarga el catálogo completo de la tienda, lo compara con la
captura anterior y anota en techmarleys.db cada cambio:

  VENTA              el stock bajó (unidades vendidas x precio = monto)
  VENTA (AGOTADO)    el stock bajó y quedó en 0
  INGRESO            el stock subió (llegó mercancía)
  NUEVO              apareció un producto que no existía
  CAMBIO DE PRECIO   cambió el precio final
  RETIRADO           el producto ya no aparece en la web
  REAPARECIDO        un producto retirado volvió a aparecer

Después de cada captura se regenera el Excel del mes en la carpeta reportes/.

Uso:
    python rastreador.py               toma una captura ahora
    python rastreador.py --programado  igual, pero solo entre 10 am y 8 pm
"""

import json
import logging
import sqlite3
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
DB = CARPETA / "techmarleys.db"
LOG = CARPETA / "rastreador.log"

API = "https://techmarleys.com/Dashboard/productos_api.php"
POR_PAGINA = 48  # el servidor no devuelve más de 48 por página
MAX_PAGINAS = 200
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) TechMarleysSeguimiento/1.0"
HORA_INICIO, HORA_FIN = 10, 20  # horario de las capturas programadas

log = logging.getLogger("rastreador")

ESQUEMA = """
CREATE TABLE IF NOT EXISTS productos (
    id INTEGER PRIMARY KEY,
    nombre TEXT, marca TEXT, modelo TEXT, condicion TEXT,
    categoria TEXT, subcategoria TEXT, sucursal TEXT,
    stock INTEGER, precio_venta REAL, precio_oferta REAL, precio_final REAL,
    disponible INTEGER, activo INTEGER,
    primera_vez TEXT, ultima_vez TEXT, foto TEXT
);
CREATE TABLE IF NOT EXISTS capturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_hora TEXT, ok INTEGER, error TEXT,
    total_productos INTEGER, total_unidades INTEGER, valor REAL
);
CREATE TABLE IF NOT EXISTS stock_historial (
    captura_id INTEGER, producto_id INTEGER, stock INTEGER, precio_final REAL,
    PRIMARY KEY (captura_id, producto_id)
);
CREATE TABLE IF NOT EXISTS eventos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    captura_id INTEGER, fecha_hora TEXT, fecha TEXT, hora TEXT,
    producto_id INTEGER, nombre TEXT, marca TEXT, modelo TEXT, condicion TEXT,
    categoria TEXT, subcategoria TEXT, sucursal TEXT,
    tipo TEXT, stock_antes INTEGER, stock_despues INTEGER, unidades INTEGER,
    precio_antes REAL, precio REAL, monto REAL, detalle TEXT
);
CREATE INDEX IF NOT EXISTS ix_eventos_fecha ON eventos(fecha);
CREATE INDEX IF NOT EXISTS ix_eventos_tipo ON eventos(tipo);
"""


def conectar():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.executescript(ESQUEMA)
    if "foto" not in {r["name"] for r in con.execute("PRAGMA table_info(productos)")}:
        con.execute("ALTER TABLE productos ADD COLUMN foto TEXT")
    if "valor" not in {r["name"] for r in con.execute("PRAGMA table_info(capturas)")}:
        con.execute("ALTER TABLE capturas ADD COLUMN valor REAL")
        con.execute("""UPDATE capturas SET valor = (SELECT ROUND(SUM(h.stock * h.precio_final), 2)
                       FROM stock_historial h WHERE h.captura_id = capturas.id) WHERE ok = 1""")
        con.commit()
    return con


def numero(valor):
    try:
        return float(valor) if valor not in (None, "") else None
    except (TypeError, ValueError):
        return None


def descargar_pagina(pagina):
    params = urllib.parse.urlencode({
        "listar_publicados": 1, "page": pagina, "limit": POR_PAGINA,
        "q": "", "cat": 0, "sucursal": 0, "filter": "all", "sort": "newest",
    })
    peticion = urllib.request.Request(f"{API}?{params}", headers={"User-Agent": USER_AGENT})
    for intento in range(3):
        try:
            with urllib.request.urlopen(peticion, timeout=40) as r:
                datos = json.loads(r.read().decode("utf-8"))
            if not datos.get("ok"):
                raise RuntimeError(f"la API respondió ok=false en la página {pagina}")
            return datos
        except Exception:
            if intento == 2:
                raise
            time.sleep(5 * (intento + 1))


def descargar_catalogo():
    """Devuelve {id: producto} con todas las páginas del catálogo."""
    productos = {}
    for pagina in range(1, MAX_PAGINAS + 1):
        datos = descargar_pagina(pagina)
        for p in datos.get("productos", []):
            productos[int(p["id"])] = {
                "id": int(p["id"]),
                "nombre": (p.get("nombre") or "").strip(),
                "marca": p.get("marca") or "",
                "modelo": p.get("modelo") or "",
                "condicion": p.get("condicion") or "",
                "categoria": p.get("categoria") or "",
                "subcategoria": p.get("subcategoria") or "",
                "sucursal": p.get("sucursal") or p.get("ubicacion") or "",
                "stock": int(p.get("stock_disponible") or 0),
                "precio_venta": numero(p.get("precio_venta")),
                "precio_oferta": numero(p.get("precio_oferta")),
                "precio_final": numero(p.get("precio_final")) or numero(p.get("precio_venta")) or 0.0,
                "disponible": 1 if p.get("disponible") else 0,
                "foto": p.get("foto") or "",
            }
        if not datos.get("has_more"):
            break
        time.sleep(0.5)
    return productos


def capturar():
    ahora = datetime.now()
    marca_tiempo = ahora.strftime("%Y-%m-%d %H:%M:%S")
    con = conectar()

    try:
        productos = descargar_catalogo()
        previos = {r["id"]: dict(r) for r in con.execute("SELECT * FROM productos")}
        activos_previos = sum(1 for r in previos.values() if r["activo"])
        # Si la web devuelve muy pocos productos de golpe, es una falla de la
        # página y no queremos marcar cientos de productos como "retirados".
        if not productos or (activos_previos and len(productos) < activos_previos * 0.5):
            raise RuntimeError(
                f"la web devolvió {len(productos)} productos (antes había {activos_previos}); se ignora esta captura")
    except Exception as e:
        con.execute("INSERT INTO capturas (fecha_hora, ok, error) VALUES (?, 0, ?)", (marca_tiempo, str(e)))
        con.commit()
        con.close()
        log.error("Captura fallida: %s", e)
        return False

    es_primera = con.execute("SELECT 1 FROM capturas WHERE ok = 1 LIMIT 1").fetchone() is None
    total_unidades = sum(p["stock"] for p in productos.values())
    total_valor = round(sum(p["stock"] * p["precio_final"] for p in productos.values()), 2)
    captura_id = con.execute(
        "INSERT INTO capturas (fecha_hora, ok, total_productos, total_unidades, valor) VALUES (?, 1, ?, ?, ?)",
        (marca_tiempo, len(productos), total_unidades, total_valor)).lastrowid

    eventos = []

    def evento(tipo, p, stock_antes=None, stock_despues=None, unidades=None,
               precio_antes=None, precio=None, monto=None, detalle=""):
        eventos.append((
            captura_id, marca_tiempo, ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M"),
            p["id"], p["nombre"], p["marca"], p["modelo"], p["condicion"],
            p["categoria"], p["subcategoria"], p["sucursal"],
            tipo, stock_antes, stock_despues, unidades, precio_antes, precio,
            round(monto, 2) if monto is not None else None, detalle))

    for pid, p in productos.items():
        antes = previos.get(pid)
        stock, precio = p["stock"], p["precio_final"]

        if antes is None:
            if not es_primera:
                evento("NUEVO", p, 0, stock, stock, None, precio, stock * precio,
                       "Producto nuevo en la web" + ("" if stock else " (sin stock)"))
        elif not antes["activo"]:
            evento("REAPARECIDO", p, 0, stock, stock, antes["precio_final"], precio, stock * precio,
                   "Volvió a aparecer en la web")
        else:
            diferencia = stock - (antes["stock"] or 0)
            precio_previo = antes["precio_final"] or 0.0
            if diferencia < 0:
                vendidas = -diferencia
                evento("VENTA (AGOTADO)" if stock <= 0 else "VENTA", p, antes["stock"], stock, vendidas,
                       precio_previo, precio_previo, vendidas * precio_previo,
                       "Se agotó" if stock <= 0 else f"Quedan {stock}")
            elif diferencia > 0:
                evento("INGRESO", p, antes["stock"], stock, diferencia, precio_previo, precio, diferencia * precio,
                       "Reposición: estaba agotado" if (antes["stock"] or 0) <= 0 else "Entró más stock")
            if abs(precio - precio_previo) > 0.005:
                cambio = f"{(precio - precio_previo) / precio_previo:+.1%}" if precio_previo else ""
                evento("CAMBIO DE PRECIO", p, antes["stock"], stock, None, precio_previo, precio,
                       None, f"{precio_previo:,.2f} -> {precio:,.2f} {cambio}".strip())

        con.execute("""
            INSERT INTO productos (id, nombre, marca, modelo, condicion, categoria, subcategoria, sucursal,
                                   stock, precio_venta, precio_oferta, precio_final, disponible, activo,
                                   primera_vez, ultima_vez, foto)
            VALUES (:id, :nombre, :marca, :modelo, :condicion, :categoria, :subcategoria, :sucursal,
                    :stock, :precio_venta, :precio_oferta, :precio_final, :disponible, 1, :ahora, :ahora, :foto)
            ON CONFLICT(id) DO UPDATE SET
                nombre = excluded.nombre, marca = excluded.marca, modelo = excluded.modelo,
                condicion = excluded.condicion, categoria = excluded.categoria,
                subcategoria = excluded.subcategoria, sucursal = excluded.sucursal,
                stock = excluded.stock, precio_venta = excluded.precio_venta,
                precio_oferta = excluded.precio_oferta, precio_final = excluded.precio_final,
                disponible = excluded.disponible, activo = 1, ultima_vez = excluded.ultima_vez,
                foto = excluded.foto
        """, {**p, "ahora": marca_tiempo})
        con.execute("INSERT INTO stock_historial VALUES (?, ?, ?, ?)", (captura_id, pid, stock, precio))

    for pid, antes in previos.items():
        if antes["activo"] and pid not in productos:
            quedaban = antes["stock"] or 0
            precio = antes["precio_final"] or 0.0
            evento("RETIRADO", antes, quedaban, 0, quedaban, precio, precio, quedaban * precio,
                   "Ya no aparece en la web (posible venta o lo quitaron)")
            con.execute("UPDATE productos SET activo = 0, stock = 0 WHERE id = ?", (pid,))

    con.executemany(f"INSERT INTO eventos VALUES (NULL{', ?' * 20})", eventos)
    con.commit()
    con.close()

    if es_primera:
        log.info("Primera captura: %d productos, %d unidades en stock (punto de partida, sin eventos)",
                 len(productos), total_unidades)
    else:
        resumen = {}
        for e in eventos:
            resumen[e[12]] = resumen.get(e[12], 0) + 1
        texto = ", ".join(f"{t}: {n}" for t, n in sorted(resumen.items())) or "sin cambios"
        log.info("Captura OK: %d productos, %d unidades. %s", len(productos), total_unidades, texto)
    return True


def main():
    salidas = [logging.FileHandler(LOG, encoding="utf-8")]
    if sys.stderr:  # con pythonw (tarea programada) no hay consola
        salidas.append(logging.StreamHandler())
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=salidas)

    if "--programado" in sys.argv and not (HORA_INICIO <= datetime.now().hour <= HORA_FIN):
        log.info("Fuera del horario %d:00-%d:00, no se toma captura", HORA_INICIO, HORA_FIN)
        return

    capturar()

    try:
        import dashboard
        dashboard.generar()
        log.info("Dashboard actualizado")
    except Exception as e:
        log.exception("Error generando el dashboard: %s", e)

    try:
        import reporte
        ruta = reporte.generar()
        if ruta:
            log.info("Reporte actualizado: %s", ruta)
    except PermissionError:
        log.warning("No se pudo actualizar el Excel (¿está abierto?). Se actualizará en la próxima captura.")
    except Exception as e:
        log.exception("Error generando el reporte: %s", e)


if __name__ == "__main__":
    main()
