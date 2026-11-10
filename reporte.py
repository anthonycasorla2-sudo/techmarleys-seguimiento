"""Genera el Excel mensual de seguimiento de techmarleys.com a partir de techmarleys.db.

Uso:
    python reporte.py            mes actual
    python reporte.py 2026-11    otro mes
"""

import sqlite3
import sys
from datetime import date, datetime
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
# openpyxl va instalado en lib/ para que la tarea programada lo encuentre siempre.
sys.path.insert(0, str(CARPETA / "lib"))

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DB = CARPETA / "techmarleys.db"
SALIDA = CARPETA / "reportes"
URL_PRODUCTO = "https://techmarleys.com/producto.php?id={}"

DINERO = '"$"#,##0.00'
ENTERO = "#,##0"
PORCENTAJE = "0.0%"
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]

VENTAS = "tipo IN ('VENTA', 'VENTA (AGOTADO)')"
INGRESOS = "tipo IN ('INGRESO', 'NUEVO', 'REAPARECIDO')"

FUENTE_ENCABEZADO = Font(bold=True, color="FFFFFF")
RELLENO_ENCABEZADO = PatternFill("solid", fgColor="7F13EC")
FUENTE_TITULO = Font(bold=True, size=16, color="7F13EC")
FUENTE_SECCION = Font(bold=True, size=12, color="7F13EC")


def dia_semana(fecha):
    return DIAS[date.fromisoformat(fecha).weekday()]


def encabezado(ws, fila, columna, texto):
    celda = ws.cell(fila, columna, texto)
    celda.font = FUENTE_ENCABEZADO
    celda.fill = RELLENO_ENCABEZADO
    celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def hoja(wb, titulo, columnas, filas, enlace_id=False):
    """columnas: lista de (encabezado, ancho, formato). Con enlace_id, la
    columna 'Producto' enlaza a la página del producto (usa la columna 'ID')."""
    ws = wb.create_sheet(titulo)
    nombres = [c[0] for c in columnas]
    for i, (texto, ancho, _) in enumerate(columnas, 1):
        encabezado(ws, 1, i, texto)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    for fila in filas:
        ws.append(list(fila))
    for i, (_, _, formato) in enumerate(columnas, 1):
        if formato:
            for (celda,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                celda.number_format = formato
    if enlace_id and "ID" in nombres and "Producto" in nombres:
        col_id, col_prod = nombres.index("ID") + 1, nombres.index("Producto") + 1
        for r in range(2, ws.max_row + 1):
            pid = ws.cell(r, col_id).value
            if pid:
                ws.cell(r, col_prod).hyperlink = URL_PRODUCTO.format(pid)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    else:
        ws.cell(2, 1, "(todavía no hay datos este mes)")
    return ws


def tabla(ws, fila, titulo, columnas, filas):
    """Escribe una tabla pequeña dentro de la hoja Resumen; devuelve la siguiente fila libre."""
    ws.cell(fila, 1, titulo).font = FUENTE_SECCION
    fila += 1
    for i, (texto, _) in enumerate(columnas, 1):
        encabezado(ws, fila, i, texto)
    for valores in filas:
        fila += 1
        for i, (valor, (_, formato)) in enumerate(zip(valores, columnas), 1):
            celda = ws.cell(fila, i, valor)
            if formato:
                celda.number_format = formato
    if not filas:
        fila += 1
        ws.cell(fila, 1, "(sin datos)")
    return fila + 2


def generar(mes=None):
    mes = mes or datetime.now().strftime("%Y-%m")
    if not DB.exists():
        return None
    con = sqlite3.connect(DB)
    q = lambda sql, *args: con.execute(sql, args).fetchall()
    uno = lambda sql, *args: con.execute(sql, args).fetchone()
    filtro_mes = "substr(fecha, 1, 7) = ?"

    wb = Workbook()
    wb.remove(wb.active)

    # ---------- Resumen ----------
    ws = wb.create_sheet("Resumen")
    año, num_mes = map(int, mes.split("-"))
    ws["A1"] = f"Seguimiento TechMarleys - {MESES[num_mes - 1]} {año}"
    ws["A1"].font = FUENTE_TITULO
    ws["A2"] = f"Actualizado: {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = Font(italic=True, color="666666")

    capturas = uno("SELECT COUNT(*), MIN(fecha_hora), MAX(fecha_hora) FROM capturas "
                   "WHERE ok = 1 AND substr(fecha_hora, 1, 7) = ?", mes)
    fallidas = uno("SELECT COUNT(*) FROM capturas WHERE ok = 0 AND substr(fecha_hora, 1, 7) = ?", mes)[0]
    v_unid, v_monto, v_prod, v_eventos = uno(
        f"SELECT COALESCE(SUM(unidades), 0), COALESCE(SUM(monto), 0), COUNT(DISTINCT producto_id), COUNT(*) "
        f"FROM eventos WHERE {VENTAS} AND {filtro_mes}", mes)
    agotados = uno(f"SELECT COUNT(*) FROM eventos WHERE tipo = 'VENTA (AGOTADO)' AND {filtro_mes}", mes)[0]
    i_unid, i_valor = uno(f"SELECT COALESCE(SUM(unidades), 0), COALESCE(SUM(monto), 0) "
                          f"FROM eventos WHERE {INGRESOS} AND {filtro_mes}", mes)
    nuevos = uno(f"SELECT COUNT(*) FROM eventos WHERE tipo = 'NUEVO' AND {filtro_mes}", mes)[0]
    precios = uno(f"SELECT COUNT(*) FROM eventos WHERE tipo = 'CAMBIO DE PRECIO' AND {filtro_mes}", mes)[0]
    r_cant, r_unid, r_monto = uno(f"SELECT COUNT(*), COALESCE(SUM(unidades), 0), COALESCE(SUM(monto), 0) "
                                  f"FROM eventos WHERE tipo = 'RETIRADO' AND {filtro_mes}", mes)
    inv_prod, inv_con_stock, inv_unid, inv_valor = uno(
        "SELECT COUNT(*), SUM(stock > 0), COALESCE(SUM(stock), 0), COALESCE(SUM(stock * precio_final), 0) "
        "FROM productos WHERE activo = 1")
    dias_con_datos = uno(f"SELECT COUNT(DISTINCT substr(fecha_hora, 1, 10)) FROM capturas "
                         f"WHERE ok = 1 AND substr(fecha_hora, 1, 7) = ?", mes)[0]

    metricas = [
        ("VENTAS (estimadas por bajadas de stock)", None, None),
        ("Unidades vendidas", v_unid, ENTERO),
        ("Monto vendido", v_monto, DINERO),
        ("Ticket promedio por unidad", (v_monto / v_unid) if v_unid else 0, DINERO),
        ("Promedio vendido por día", (v_monto / dias_con_datos) if dias_con_datos else 0, DINERO),
        ("Productos distintos vendidos", v_prod, ENTERO),
        ("Productos que se agotaron", agotados, ENTERO),
        ("", None, None),
        ("MERCANCÍA QUE ENTRÓ", None, None),
        ("Unidades ingresadas (reposiciones + nuevos)", i_unid, ENTERO),
        ("Valor de la mercancía ingresada (a precio de venta)", i_valor, DINERO),
        ("Productos nuevos publicados", nuevos, ENTERO),
        ("Cambios de precio", precios, ENTERO),
        ("Productos retirados de la web", r_cant, ENTERO),
        ("  unidades que tenían (posibles ventas)", r_unid, ENTERO),
        ("  monto posible", r_monto, DINERO),
        ("", None, None),
        ("INVENTARIO ACTUAL EN LA WEB", None, None),
        ("Productos publicados", inv_prod, ENTERO),
        ("Productos con stock", inv_con_stock or 0, ENTERO),
        ("Productos agotados", inv_prod - (inv_con_stock or 0), ENTERO),
        ("Unidades en stock", inv_unid, ENTERO),
        ("Valor del inventario (a precio de venta)", inv_valor, DINERO),
        ("", None, None),
        ("Capturas tomadas este mes", capturas[0], ENTERO),
        ("Capturas fallidas", fallidas, ENTERO),
        ("Primera captura del mes", capturas[1] or "-", None),
        ("Última captura", capturas[2] or "-", None),
    ]
    fila = 4
    for etiqueta, valor, formato in metricas:
        celda = ws.cell(fila, 1, etiqueta)
        if valor is None and etiqueta:
            celda.font = FUENTE_SECCION
        elif etiqueta:
            v = ws.cell(fila, 2, valor)
            v.font = Font(bold=True)
            if formato:
                v.number_format = formato
        fila += 1
    fila += 1

    fila = tabla(ws, fila, "Top 10 más vendidos", [("Producto", None), ("Unidades", ENTERO), ("Monto", DINERO)],
                 q(f"SELECT MAX(nombre), SUM(unidades), SUM(monto) FROM eventos WHERE {VENTAS} AND {filtro_mes} "
                   f"GROUP BY producto_id ORDER BY SUM(unidades) DESC, SUM(monto) DESC LIMIT 10", mes))
    fila = tabla(ws, fila, "Ventas por sucursal", [("Sucursal", None), ("Unidades", ENTERO), ("Monto", DINERO)],
                 q(f"SELECT sucursal, SUM(unidades), SUM(monto) FROM eventos WHERE {VENTAS} AND {filtro_mes} "
                   f"GROUP BY sucursal ORDER BY SUM(monto) DESC", mes))
    tabla(ws, fila, "Ventas por condición", [("Condición", None), ("Unidades", ENTERO), ("Monto", DINERO)],
          q(f"SELECT condicion, SUM(unidades), SUM(monto) FROM eventos WHERE {VENTAS} AND {filtro_mes} "
            f"GROUP BY condicion ORDER BY SUM(monto) DESC", mes))
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16

    # ---------- Ventas por día ----------
    por_dia = {r[0]: r[1:] for r in q(f"""
        SELECT fecha,
               COALESCE(SUM(CASE WHEN {VENTAS} THEN unidades END), 0),
               COALESCE(SUM(CASE WHEN {VENTAS} THEN monto END), 0),
               COUNT(DISTINCT CASE WHEN {VENTAS} THEN producto_id END),
               SUM(tipo = 'VENTA (AGOTADO)'),
               COALESCE(SUM(CASE WHEN {INGRESOS} THEN unidades END), 0),
               SUM(tipo = 'NUEVO'),
               SUM(tipo = 'RETIRADO')
        FROM eventos WHERE {filtro_mes} GROUP BY fecha""", mes)}
    fechas = sorted(set(por_dia) | {r[0] for r in q(
        "SELECT DISTINCT substr(fecha_hora, 1, 10) FROM capturas WHERE ok = 1 AND substr(fecha_hora, 1, 7) = ?", mes)})
    filas = [(f, dia_semana(f), *por_dia.get(f, (0, 0, 0, 0, 0, 0, 0))) for f in fechas]
    ws = hoja(wb, "Ventas por día", [
        ("Fecha", 12, None), ("Día", 11, None), ("Unidades vendidas", 11, ENTERO), ("Monto vendido", 15, DINERO),
        ("Productos distintos", 11, ENTERO), ("Agotados", 10, ENTERO), ("Unidades ingresadas", 11, ENTERO),
        ("Productos nuevos", 10, ENTERO), ("Retirados", 10, ENTERO)], filas)
    if filas:
        grafico = BarChart()
        grafico.title = "Monto vendido por día"
        grafico.y_axis.title = "US$"
        grafico.legend = None
        grafico.add_data(Reference(ws, min_col=4, min_row=1, max_row=len(filas) + 1), titles_from_data=True)
        grafico.set_categories(Reference(ws, min_col=1, min_row=2, max_row=len(filas) + 1))
        grafico.width, grafico.height = 24, 10
        ws.add_chart(grafico, "K2")

    # ---------- Ventas por producto ----------
    hoja(wb, "Ventas por producto", [
        ("ID", 7, None), ("Producto", 45, None), ("Marca", 12, None), ("Modelo", 14, None), ("Condición", 10, None),
        ("Categoría", 14, None), ("Subcategoría", 16, None), ("Sucursal", 11, None),
        ("Unidades vendidas", 10, ENTERO), ("Monto", 13, DINERO), ("Precio promedio", 12, DINERO),
        ("Veces que se agotó", 10, ENTERO), ("Días con venta", 9, ENTERO),
        ("Primera venta", 17, None), ("Última venta", 17, None), ("Stock actual", 9, ENTERO)],
        q(f"""SELECT e.producto_id, MAX(e.nombre), MAX(e.marca), MAX(e.modelo), MAX(e.condicion),
                     MAX(e.categoria), MAX(e.subcategoria), MAX(e.sucursal),
                     SUM(e.unidades), SUM(e.monto), SUM(e.monto) / SUM(e.unidades),
                     SUM(e.tipo = 'VENTA (AGOTADO)'), COUNT(DISTINCT e.fecha),
                     MIN(e.fecha_hora), MAX(e.fecha_hora), p.stock
              FROM eventos e LEFT JOIN productos p ON p.id = e.producto_id
              WHERE e.{VENTAS} AND {filtro_mes}
              GROUP BY e.producto_id ORDER BY SUM(e.unidades) DESC, SUM(e.monto) DESC""", mes),
        enlace_id=True)

    # ---------- Ventas por categoría ----------
    total = v_monto or 1
    hoja(wb, "Ventas por categoría", [
        ("Categoría", 18, None), ("Subcategoría", 24, None), ("Unidades", 10, ENTERO), ("Monto", 14, DINERO),
        ("% del monto", 10, PORCENTAJE), ("Productos distintos", 11, ENTERO)],
        [(*r[:4], r[3] / total, r[4]) for r in q(
            f"SELECT categoria, subcategoria, SUM(unidades), SUM(monto), COUNT(DISTINCT producto_id) "
            f"FROM eventos WHERE {VENTAS} AND {filtro_mes} GROUP BY categoria, subcategoria "
            f"ORDER BY SUM(monto) DESC", mes)])

    # ---------- Ventas por marca ----------
    hoja(wb, "Ventas por marca", [
        ("Marca", 18, None), ("Unidades", 10, ENTERO), ("Monto", 14, DINERO),
        ("% del monto", 10, PORCENTAJE), ("Productos distintos", 11, ENTERO)],
        [(r[0], r[1], r[2], r[2] / total, r[3]) for r in q(
            f"SELECT marca, SUM(unidades), SUM(monto), COUNT(DISTINCT producto_id) "
            f"FROM eventos WHERE {VENTAS} AND {filtro_mes} GROUP BY marca ORDER BY SUM(monto) DESC", mes)])

    # ---------- Detalle de ventas ----------
    hoja(wb, "Detalle de ventas", [
        ("Fecha", 11, None), ("Día", 10, None), ("Hora", 7, None), ("ID", 7, None), ("Producto", 45, None),
        ("Marca", 12, None), ("Modelo", 14, None), ("Condición", 10, None), ("Categoría", 14, None),
        ("Subcategoría", 16, None), ("Sucursal", 11, None), ("Stock antes", 8, ENTERO),
        ("Stock después", 8, ENTERO), ("Unidades", 9, ENTERO), ("Precio", 12, DINERO), ("Monto", 13, DINERO),
        ("¿Se agotó?", 9, None)],
        [(r[0], dia_semana(r[0]), *r[1:-1], "Sí" if r[-1] == "VENTA (AGOTADO)" else "No") for r in q(
            f"SELECT fecha, hora, producto_id, nombre, marca, modelo, condicion, categoria, subcategoria, sucursal, "
            f"stock_antes, stock_despues, unidades, precio, monto, tipo FROM eventos "
            f"WHERE {VENTAS} AND {filtro_mes} ORDER BY fecha_hora DESC, monto DESC", mes)],
        enlace_id=True)

    # ---------- Agotados ----------
    hoja(wb, "Agotados", [
        ("Fecha", 11, None), ("Día", 10, None), ("Hora", 7, None), ("ID", 7, None), ("Producto", 45, None),
        ("Marca", 12, None), ("Condición", 10, None), ("Categoría", 14, None), ("Subcategoría", 16, None),
        ("Sucursal", 11, None), ("Unidades vendidas", 9, ENTERO), ("Precio", 12, DINERO), ("Monto", 13, DINERO),
        ("Estado hoy", 16, None)],
        [(r[0], dia_semana(r[0]), *r[1:-1], "Repuesto" if (r[-1] or 0) > 0 else "Sigue agotado") for r in q(
            f"SELECT e.fecha, e.hora, e.producto_id, e.nombre, e.marca, e.condicion, e.categoria, e.subcategoria, "
            f"e.sucursal, e.unidades, e.precio, e.monto, p.stock FROM eventos e "
            f"LEFT JOIN productos p ON p.id = e.producto_id "
            f"WHERE e.tipo = 'VENTA (AGOTADO)' AND {filtro_mes} ORDER BY e.fecha_hora DESC", mes)],
        enlace_id=True)

    # ---------- Mercancía ingresada ----------
    hoja(wb, "Mercancía ingresada", [
        ("Fecha", 11, None), ("Hora", 7, None), ("Tipo", 13, None), ("ID", 7, None), ("Producto", 45, None),
        ("Marca", 12, None), ("Modelo", 14, None), ("Condición", 10, None), ("Categoría", 14, None),
        ("Subcategoría", 16, None), ("Sucursal", 11, None), ("Stock antes", 8, ENTERO),
        ("Stock después", 8, ENTERO), ("Unidades", 9, ENTERO), ("Precio", 12, DINERO), ("Valor", 13, DINERO),
        ("Detalle", 28, None)],
        q(f"SELECT fecha, hora, tipo, producto_id, nombre, marca, modelo, condicion, categoria, subcategoria, "
          f"sucursal, stock_antes, stock_despues, unidades, precio, monto, detalle FROM eventos "
          f"WHERE {INGRESOS} AND {filtro_mes} ORDER BY fecha_hora DESC", mes),
        enlace_id=True)

    # ---------- Cambios de precio ----------
    hoja(wb, "Cambios de precio", [
        ("Fecha", 11, None), ("Hora", 7, None), ("ID", 7, None), ("Producto", 45, None), ("Marca", 12, None),
        ("Categoría", 14, None), ("Sucursal", 11, None), ("Precio antes", 12, DINERO), ("Precio nuevo", 12, DINERO),
        ("Diferencia", 12, DINERO), ("Cambio %", 9, PORCENTAJE)],
        [(*r, r[-1] - r[-2], (r[-1] - r[-2]) / r[-2] if r[-2] else None) for r in q(
            f"SELECT fecha, hora, producto_id, nombre, marca, categoria, sucursal, precio_antes, precio "
            f"FROM eventos WHERE tipo = 'CAMBIO DE PRECIO' AND {filtro_mes} ORDER BY fecha_hora DESC", mes)],
        enlace_id=True)

    # ---------- Retirados ----------
    hoja(wb, "Retirados de la web", [
        ("Fecha", 11, None), ("Hora", 7, None), ("ID", 7, None), ("Producto", 45, None), ("Marca", 12, None),
        ("Condición", 10, None), ("Categoría", 14, None), ("Sucursal", 11, None),
        ("Stock que tenía", 9, ENTERO), ("Precio", 12, DINERO), ("Monto posible", 13, DINERO)],
        q(f"SELECT fecha, hora, producto_id, nombre, marca, condicion, categoria, sucursal, unidades, precio, monto "
          f"FROM eventos WHERE tipo = 'RETIRADO' AND {filtro_mes} ORDER BY fecha_hora DESC", mes),
        enlace_id=True)

    # ---------- Inventario actual ----------
    hoja(wb, "Inventario actual", [
        ("ID", 7, None), ("Producto", 45, None), ("Marca", 12, None), ("Modelo", 14, None), ("Condición", 10, None),
        ("Categoría", 14, None), ("Subcategoría", 16, None), ("Sucursal", 11, None), ("Stock", 7, ENTERO),
        ("Precio", 12, DINERO), ("Valor en stock", 13, DINERO), ("Estado", 11, None), ("Visto desde", 17, None)],
        [(*r[:10], r[8] * r[9], "Disponible" if r[8] > 0 else "Agotado", r[10]) for r in q(
            "SELECT id, nombre, marca, modelo, condicion, categoria, subcategoria, sucursal, stock, precio_final, "
            "primera_vez FROM productos WHERE activo = 1 ORDER BY categoria, subcategoria, nombre")],
        enlace_id=True)

    con.close()
    SALIDA.mkdir(exist_ok=True)
    ruta = SALIDA / f"reporte_{mes}.xlsx"
    wb.save(ruta)
    return ruta


if __name__ == "__main__":
    print(generar(sys.argv[1] if len(sys.argv) > 1 else None))
