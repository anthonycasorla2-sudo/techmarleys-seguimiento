"""Versión en la nube: GitHub Actions la ejecuta cada hora.

1. Reconstruye techmarleys.db desde datos/techmarleys.sql (se guarda como texto
   para que git solo almacene lo que cambia en cada captura).
2. Toma la captura con rastreador.py.
3. Vuelve a escribir datos/techmarleys.sql.
4. Arma la carpeta sitio/ (dashboard + Excel de cada mes) que se publica en GitHub Pages.

Uso:
    python nube.py               captura y arma el sitio
    python nube.py --programado  igual, pero solo captura entre 10 am y 8 pm
"""

import logging
import shutil
import sqlite3
import sys
from datetime import datetime, timedelta
from pathlib import Path

import dashboard
import rastreador
import reporte

CARPETA = Path(__file__).resolve().parent
DUMP = CARPETA / "datos" / "techmarleys.sql"
SITIO = CARPETA / "sitio"
WEB = CARPETA / "web"
DIAS_STOCK_DETALLE = 14  # el stock de cada producto en cada captura solo se guarda 2 semanas

# Base de trabajo propia, para no tocar techmarleys.db si esto se ejecuta en la PC.
rastreador.DB = reporte.DB = dashboard.DB = CARPETA / "nube.db"
reporte.SALIDA = SITIO.parent / "reportes_nube"

log = logging.getLogger("nube")


def cargar():
    rastreador.DB.unlink(missing_ok=True)
    if DUMP.exists():
        con = sqlite3.connect(rastreador.DB)
        con.executescript(DUMP.read_text(encoding="utf-8"))
        con.close()


def guardar():
    con = rastreador.conectar()  # aplica las migraciones pendientes
    limite = (datetime.now() - timedelta(days=DIAS_STOCK_DETALLE)).strftime("%Y-%m-%d %H:%M:%S")
    con.execute("DELETE FROM stock_historial WHERE captura_id IN (SELECT id FROM capturas WHERE fecha_hora < ?)",
                (limite,))
    con.commit()
    DUMP.parent.mkdir(exist_ok=True)
    DUMP.write_text("\n".join(con.iterdump()) + "\n", encoding="utf-8")
    con.close()


def armar_sitio():
    shutil.rmtree(SITIO, ignore_errors=True)
    (SITIO / "reportes").mkdir(parents=True)
    con = sqlite3.connect(rastreador.DB)
    meses = [r[0] for r in con.execute(
        "SELECT DISTINCT substr(fecha_hora, 1, 7) FROM capturas WHERE ok = 1 ORDER BY 1")]
    con.close()
    excels = {}
    for mes in meses:
        ruta = reporte.generar(mes)
        shutil.copy(ruta, SITIO / "reportes" / ruta.name)
        excels[mes] = f"reportes/{ruta.name}"
    dashboard.generar(rastreador.DB, SITIO / "dashboard_datos.js", excels)
    shutil.copy(CARPETA / "dashboard.html", SITIO / "index.html")
    for archivo in WEB.iterdir():
        shutil.copy(archivo, SITIO / archivo.name)


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", stream=sys.stdout)
    cargar()
    hora = datetime.now().hour
    if "--programado" in sys.argv and not (rastreador.HORA_INICIO <= hora <= rastreador.HORA_FIN):
        log.info("Fuera del horario %d:00-%d:00, solo se vuelve a publicar el sitio",
                 rastreador.HORA_INICIO, rastreador.HORA_FIN)
    else:
        rastreador.capturar()
    guardar()
    armar_sitio()
    log.info("Sitio listo en %s", SITIO)


if __name__ == "__main__":
    main()
