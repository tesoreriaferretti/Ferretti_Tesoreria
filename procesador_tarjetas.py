"""Logica de negocio: integra reportes de Culqi y Mercado Pago en FP Tarjetas 2026 / FP Surco Tarjetas 2026."""
from copy import copy
from io import BytesIO
import re

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
import pandas as pd


# Comercios permitidos por empresa
COMERCIOS_POR_EMPRESA = {
    "FP_STORES": ["FP STORES"],
    "FP_SURCO": ["FP STORES SURCO", "FP SURCO"],
}

# Mapeo de hojas Culqi: (Producto, Moneda) -> nombre de pestaña
MAPA_CULQI = {
    ("POS", "PEN"): "Culqi POS MN",
    ("POS", "USD"): "Culqi POS ME",
    ("PAYMENT_LINK", "PEN"): "Culqi Link MN",
    ("PAYMENT_LINK", "USD"): "Culqi Link ME",
}


def limpiar_str(val):
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def parsear_monto(val):
    if val is None or (not isinstance(val, (str, int, float)) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^\d.\-]", "", str(val).replace(",", ""))
    try:
        return float(s)
    except ValueError:
        return None


def pertenece_a_empresa(comercio_raw, empresa_objetivo):
    c = limpiar_str(comercio_raw).upper()
    if not c:
        return False
    if empresa_objetivo == "FP_STORES":
        return "FP STORES" in c and "SURCO" not in c
    if empresa_objetivo == "FP_SURCO":
        return "SURCO" in c
    return False


def clonar_estilo_fila(ws, fila_modelo, fila_destino, max_col):
    """Copia alto de fila, fuente, alineacion, bordes y formato numerico de una fila modelo."""
    if fila_modelo is None or fila_modelo == fila_destino:
        return

    # 1) Copiar alto exacto de fila
    alto = ws.row_dimensions[fila_modelo].height
    if alto is not None:
        ws.row_dimensions[fila_destino].height = alto

    # 2) Copiar estilos celda por celda
    for col_idx in range(1, max_col + 1):
        src = ws.cell(fila_modelo, col_idx)
        dst = ws.cell(fila_destino, col_idx)
        if src.has_style:
            dst.font = copy(src.font)
            dst.alignment = copy(src.alignment)
            dst.border = copy(src.border)
            dst.number_format = src.number_format
            # No copiamos el relleno manual en D:AK porque ahi manda el formato condicional
            if col_idx < 4:
                dst.fill = copy(src.fill)


def asegurar_formato_condicional_culqi(ws):
    """Asegura jerarquia de colores (Plomo > Amarillo > Celeste) en D5:AK5000."""
    rango = "D5:AK5000"
    ws.conditional_formatting._cf_rules.clear()

    fill_plomo = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_celeste = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    regla_plomo = FormulaRule(
        formula=['OR(ISNUMBER(SEARCH("devuelta",$F5)),ISNUMBER(SEARCH("extorn",$F5)),ISNUMBER(SEARCH("anulad",$F5)))'],
        stopIfTrue=True,
        fill=fill_plomo,
    )
    regla_amarillo = FormulaRule(
        formula=['AND(TRIM($A5)<>"",TRIM($B5)="",TRIM($C5)="")'],
        stopIfTrue=True,
        fill=fill_amarillo,
    )
    regla_celeste = FormulaRule(
        formula=['AND(TRIM($A5)<>"",OR(TRIM($B5)<>"",TRIM($C5)<>""))'],
        stopIfTrue=True,
        fill=fill_celeste,
    )

    ws.conditional_formatting.add(rango, regla_plomo)
    ws.conditional_formatting.add(rango, regla_amarillo)
    ws.conditional_formatting.add(rango, regla_celeste)


def _procesar_culqi(wb, df_raw, empresa_objetivo, resumen, simulacion):
    header_idx = None
    for idx in range(min(15, len(df_raw))):
        fila_vals = [limpiar_str(x) for x in df_raw.iloc[idx].tolist()]
        if "ID de Venta" in fila_vals and "Estado" in fila_vals:
            header_idx = idx
            break
    if header_idx is None:
        return 0

    headers = [limpiar_str(x) for x in df_raw.iloc[header_idx].tolist()]
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    nuevos_total = 0
    actualizados_total = 0

    for (prod_code, mon_code), hoja_nombre in MAPA_CULQI.items():
        if hoja_nombre not in wb.sheetnames:
            continue
        ws = wb[hoja_nombre]

        # Leer encabezados de la fila 4 del maestro (columnas D en adelante)
        cols_maestro = {}
        max_col_maestro = 37  # Columna AK = 37
        for c in range(4, ws.max_column + 1):
            h_val = limpiar_str(ws.cell(4, c).value)
            if h_val:
                cols_maestro[h_val] = c
                if c > max_col_maestro:
                    max_col_maestro = c

        col_id = cols_maestro.get("ID de Venta", 4)
        col_estado = cols_maestro.get("Estado", 6)

        # Mapear IDs existentes y detectar ultima fila con datos para usarla de modelo
        existentes = {}
        ultima_fila = 4
        for r in range(5, ws.max_row + 1):
            val_id = limpiar_str(ws.cell(r, col_id).value)
            if val_id:
                existentes[val_id] = r
                ultima_fila = r
            elif any(ws.cell(r, c).value is not None for c in range(1, max_col_maestro + 1)):
                ultima_fila = r

        fila_modelo = ultima_fila if ultima_fila >= 5 else None

        nuevos_hoja = 0
        act_hoja = 0

        for _, row in df.iterrows():
            comercio = row.get("Nombre Comercial", "")
            if not pertenece_a_empresa(comercio, empresa_objetivo):
                continue

            prod = limpiar_str(row.get("Producto", "")).upper()
            mon = limpiar_str(row.get("Moneda", "")).upper()
            if prod != prod_code or mon != mon_code:
                continue

            id_venta = limpiar_str(row.get("ID de Venta", ""))
            if not id_venta:
                continue

            estado_nuevo = limpiar_str(row.get("Estado", ""))

            if id_venta in existentes:
                r_dest = existentes[id_venta]
                estado_actual = limpiar_str(ws.cell(r_dest, col_estado).value)
                if estado_nuevo and estado_nuevo.lower() != estado_actual.lower():
                    if not simulacion:
                        for h_name, c_idx in cols_maestro.items():
                            if h_name in row.index and pd.notna(row[h_name]):
                                val = row[h_name]
                                ws.cell(r_dest, c_idx).value = val.item() if hasattr(val, "item") else val
                    act_hoja += 1
            else:
                ultima_fila += 1
                existentes[id_venta] = ultima_fila
                if not simulacion:
                    for h_name, c_idx in cols_maestro.items():
                        if h_name in row.index and pd.notna(row[h_name]):
                            val = row[h_name]
                            ws.cell(ultima_fila, c_idx).value = val.item() if hasattr(val, "item") else val
                    # Aplicar exactamente el alto de fila, alineacion, fuente y bordes anteriores
                    clonar_estilo_fila(ws, fila_modelo, ultima_fila, max_col_maestro)
                    if fila_modelo is None:
                        fila_modelo = ultima_fila
                nuevos_hoja += 1

        if not simulacion:
            asegurar_formato_condicional_culqi(ws)

        if nuevos_hoja > 0 or act_hoja > 0:
            resumen.append(f"• {hoja_nombre}: {nuevos_hoja} nuevos, {act_hoja} actualizados (aprobada → abonada).")
        nuevos_total += nuevos_hoja
        actualizados_total += act_hoja

    return nuevos_total


def _procesar_mpago(wb, df_raw, resumen, simulacion):
    if "MPago" not in wb.sheetnames:
        return 0

    header_idx = None
    for idx in range(min(15, len(df_raw))):
        fila_vals = [limpiar_str(x) for x in df_raw.iloc[idx].tolist()]
        if "Número de operación" in fila_vals or "ID de operación en Mercado Pago" in fila_vals:
            header_idx = idx
            break
    if header_idx is None:
        return 0

    headers = [limpiar_str(x) for x in df_raw.iloc[header_idx].tolist()]
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    ws = wb["MPago"]
    cols_maestro = {}
    max_col_mp = 1
    for c in range(1, ws.max_column + 1):
        h_val = limpiar_str(ws.cell(1, c).value)
        if h_val:
            cols_maestro[h_val] = c
            max_col_mp = c

    col_op = cols_maestro.get("Número de operación") or cols_maestro.get("ID de operación en Mercado Pago") or 4

    existentes = set()
    ultima_fila = 1
    for r in range(2, ws.max_row + 1):
        val_op = limpiar_str(ws.cell(r, col_op).value)
        if val_op:
            existentes.add(val_op)
            ultima_fila = r
        elif any(ws.cell(r, c).value is not None for c in range(1, max_col_mp + 1)):
            ultima_fila = r

    fila_modelo = ultima_fila if ultima_fila >= 2 else None
    nuevos_mp = 0

    for _, row in df.iterrows():
        op_id = limpiar_str(
            row.get("Número de operación") if "Número de operación" in row.index else row.get("ID de operación en Mercado Pago", "")
        )
        if not op_id or op_id in existentes:
            continue

        existentes.add(op_id)
        ultima_fila += 1
        if not simulacion:
            for h_name, c_idx in cols_maestro.items():
                if h_name in row.index and pd.notna(row[h_name]):
                    val = row[h_name]
                    ws.cell(ultima_fila, c_idx).value = val.item() if hasattr(val, "item") else val
            # Clonar exactamente el alto de fila, fuente, alineacion y bordes anteriores
            clonar_estilo_fila(ws, fila_modelo, ultima_fila, max_col_mp)
            if fila_modelo is None:
                fila_modelo = ultima_fila
        nuevos_mp += 1

    if nuevos_mp > 0:
        resumen.append(f"• MPago: {nuevos_mp} operaciones nuevas agregadas.")

    return nuevos_mp


def procesar_tarjetas(master_bytes, insumos, empresa_objetivo="FP_STORES", simulacion=False):
    wb = openpyxl.load_workbook(BytesIO(master_bytes))
    resumen = []
    total_nuevos = 0

    for nombre_arch, contenido in insumos:
        xl = pd.ExcelFile(BytesIO(contenido))
        for sheet in xl.sheet_names:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            if df_raw.empty:
                continue

            muestra = " ".join(
                limpiar_str(x) for x in df_raw.iloc[:12].values.flatten() if pd.notna(x)
            )
            if "ID de Venta" in muestra and "Nombre Comercial" in muestra:
                total_nuevos += _procesar_culqi(wb, df_raw, empresa_objetivo, resumen, simulacion)
            elif "Mercado Pago" in muestra or "Número de operación" in muestra:
                total_nuevos += _procesar_mpago(wb, df_raw, resumen, simulacion)

    if not resumen:
        resumen.append("No se encontraron transacciones nuevas para esta empresa en los archivos subidos.")

    if simulacion:
        return None, total_nuevos, resumen

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), total_nuevos, resumen
