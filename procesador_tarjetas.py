"""Logica de negocio: integra reportes de Culqi y Mercado Pago en FP Tarjetas 2026 / FP Surco Tarjetas 2026."""
from copy import copy
from io import BytesIO
import re
import unicodedata

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
import pandas as pd


def limpiar_str(val):
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def normalizar_clave_col(texto):
    s = limpiar_str(texto).lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


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


def pertenece_a_empresa(empresa_val, comercio_val, empresa_objetivo):
    emp = limpiar_str(empresa_val).upper()
    com = limpiar_str(comercio_val).upper()
    texto = f"{emp} {com}"
    if not texto.strip():
        return False
    if empresa_objetivo == "FP_STORES":
        return ("FP STORES" in texto or "FERRETTI" in texto) and "SURCO" not in texto
    if empresa_objetivo == "FP_SURCO":
        return "SURCO" in texto
    return False


def resolver_hoja_culqi(comercio_val, moneda_val):
    com = limpiar_str(comercio_val).upper()
    mon = limpiar_str(moneda_val).upper()
    es_link = "LINK" in com
    if mon == "PEN":
        return "Culqi Link MN" if es_link else "Culqi POS MN"
    if mon == "USD":
        return "Culqi Link ME" if es_link else "Culqi POS ME"
    return None


def clonar_estilo_fila(ws, fila_modelo, fila_destino, max_col, proteger_relleno_hasta_col=3):
    """Copia alto de fila, fuente, alineacion, bordes y formato numerico de la ultima fila previa."""
    if fila_modelo is None or fila_modelo == fila_destino:
        return

    alto = ws.row_dimensions[fila_modelo].height
    if alto is not None:
        ws.row_dimensions[fila_destino].height = alto

    for col_idx in range(1, max_col + 1):
        src = ws.cell(fila_modelo, col_idx)
        dst = ws.cell(fila_destino, col_idx)
        if src.has_style:
            dst.font = copy(src.font)
            dst.alignment = copy(src.alignment)
            dst.border = copy(src.border)
            dst.number_format = src.number_format
            if col_idx <= proteger_relleno_hasta_col:
                dst.fill = copy(src.fill)


def asegurar_formato_condicional_culqi(ws):
    """
    Aplica las 3 reglas exactas en D5:AK5000 (Plomo > Amarillo > Celeste):
      1. Plomo: AK (Estado) es rechazada, anulada, pendiente, devuelta o extornada.
      2. Amarillo: AZ (Nro. Pedido) tiene informacion real (ignora vacio y '-').
      3. Celeste: BF (Observaciones) tiene informacion.
    """
    rango = "D5:AK5000"
    ws.conditional_formatting._cf_rules.clear()

    fill_plomo = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_celeste = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    regla_plomo = FormulaRule(
        formula=[
            'OR(ISNUMBER(SEARCH("rechazad",$AK5)),ISNUMBER(SEARCH("anulad",$AK5)),'
            'ISNUMBER(SEARCH("pendient",$AK5)),ISNUMBER(SEARCH("devuelt",$AK5)),'
            'ISNUMBER(SEARCH("extorn",$AK5)))'
        ],
        stopIfTrue=True,
        fill=fill_plomo,
    )
    regla_amarillo = FormulaRule(
        formula=['AND(TRIM($AZ5)<>"",TRIM($AZ5)<>"-")'],
        stopIfTrue=True,
        fill=fill_amarillo,
    )
    regla_celeste = FormulaRule(
        formula=['TRIM($BF5)<>""'],
        stopIfTrue=True,
        fill=fill_celeste,
    )

    ws.conditional_formatting.add(rango, regla_plomo)
    ws.conditional_formatting.add(rango, regla_amarillo)
    ws.conditional_formatting.add(rango, regla_celeste)


def asegurar_formato_condicional_mpago(ws):
    """
    Aplica las reglas de color en MPago tomando como referencia la columna B:
      1. Celeste (A2:J5000): B = 'Retiro de dinero a cuenta bancaria'
      2. Anaranjado (A2:E5000): B = 'Anulación parcial de costo de Mercado Pago',
         'Devolución parcial de cobro', 'Anulación de costo de Mercado Pago' o 'Devolución de cobro'
      3. Amarillo (A2:E5000): B = 'Cobro' con dato en F (Nro. Pedido)
      4. Amarillo (A2:E5000): B = 'Costo de Mercado Pago' cuyo ID en D coincide con un 'Cobro' ya pintado
    """
    ws.conditional_formatting._cf_rules.clear()

    fill_celeste = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    fill_naranja = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # 1. Celeste hasta columna J (A2:J5000)
    regla_celeste = FormulaRule(
        formula=['TRIM($B2)="Retiro de dinero a cuenta bancaria"'],
        stopIfTrue=True,
        fill=fill_celeste,
    )
    ws.conditional_formatting.add("A2:J5000", regla_celeste)

    # 2. Anaranjado y Amarillo en A2:E5000
    regla_naranja = FormulaRule(
        formula=[
            'OR(TRIM($B2)="Anulación parcial de costo de Mercado Pago",'
            'TRIM($B2)="Devolución parcial de cobro",'
            'TRIM($B2)="Anulación de costo de Mercado Pago",'
            'TRIM($B2)="Devolución de cobro")'
        ],
        stopIfTrue=True,
        fill=fill_naranja,
    )
    regla_cobro_amarillo = FormulaRule(
        formula=['AND(TRIM($B2)="Cobro",TRIM($F2)<>"",TRIM($F2)<>"-")'],
        stopIfTrue=True,
        fill=fill_amarillo,
    )
    regla_costo_amarillo = FormulaRule(
        formula=[
            'AND(TRIM($B2)="Costo de Mercado Pago",TRIM($D2)<>"",'
            'COUNTIFS($B$2:$B$5000,"Cobro",$D$2:$D$5000,$D2,$F$2:$F$5000,"<>",$F$2:$F$5000,"<>-")>0)'
        ],
        stopIfTrue=True,
        fill=fill_amarillo,
    )

    ws.conditional_formatting.add("A2:E5000", regla_naranja)
    ws.conditional_formatting.add("A2:E5000", regla_cobro_amarillo)
    ws.conditional_formatting.add("A2:E5000", regla_costo_amarillo)


def _limpiar_guion_celda(val_clean):
    """Convierte cualquier guion aislado '-' del reporte de Culqi en celda vacia."""
    if isinstance(val_clean, str) and val_clean.strip() == "-":
        return None
    return val_clean


def _procesar_culqi(wb, df_raw, empresa_objetivo, resumen, simulacion):
    header_idx = None
    for idx in range(min(15, len(df_raw))):
        fila_vals = [normalizar_clave_col(x) for x in df_raw.iloc[idx].tolist()]
        if ("id venta" in fila_vals or "id de venta" in fila_vals) and "estado" in fila_vals:
            header_idx = idx
            break
    if header_idx is None:
        return 0

    headers = [limpiar_str(x) for x in df_raw.iloc[header_idx].tolist()]
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    mapa_cols_df = {normalizar_clave_col(c): c for c in headers if c}
    col_df_empresa = mapa_cols_df.get("empresa", "Empresa")
    col_df_comercio = mapa_cols_df.get("comercio") or mapa_cols_df.get("nombre comercial") or "Comercio"
    col_df_moneda = mapa_cols_df.get("moneda", "Moneda")
    col_df_id_venta = mapa_cols_df.get("id venta") or mapa_cols_df.get("id de venta") or "ID Venta"
    col_df_estado = mapa_cols_df.get("estado", "Estado")

    info_hojas = {}
    for h_nom in ["Culqi POS MN", "Culqi POS ME", "Culqi Link MN", "Culqi Link ME"]:
        if h_nom not in wb.sheetnames:
            continue
        ws = wb[h_nom]
        cols_maestro = {}
        max_col_maestro = max(ws.max_column, 58)
        col_id = 4
        col_est = 37

        for c in range(1, ws.max_column + 1):
            h_val = limpiar_str(ws.cell(4, c).value)
            if h_val:
                cols_maestro[normalizar_clave_col(h_val)] = c
                if normalizar_clave_col(h_val) in ("id venta", "id de venta"):
                    col_id = c
                elif normalizar_clave_col(h_val) == "estado":
                    col_est = c

        existentes = {}
        ultima_fila = 4
        for r in range(5, ws.max_row + 1):
            val_id = limpiar_str(ws.cell(r, col_id).value)
            if val_id:
                existentes[val_id] = r
                ultima_fila = r

        info_hojas[h_nom] = {
            "ws": ws,
            "cols_maestro": cols_maestro,
            "max_col": max_col_maestro,
            "col_id": col_id,
            "col_est": col_est,
            "existentes": existentes,
            "ultima_fila": ultima_fila,
            "fila_modelo": ultima_fila if ultima_fila >= 5 else None,
            "nuevos": 0,
            "actualizados": 0,
        }

    for _, row in df.iterrows():
        emp_val = row.get(col_df_empresa, "") if col_df_empresa in row.index else ""
        com_val = row.get(col_df_comercio, "") if col_df_comercio in row.index else ""
        if not pertenece_a_empresa(emp_val, com_val, empresa_objetivo):
            continue

        mon_val = row.get(col_df_moneda, "") if col_df_moneda in row.index else ""
        hoja_dest = resolver_hoja_culqi(com_val, mon_val)
        if not hoja_dest or hoja_dest not in info_hojas:
            continue

        h_data = info_hojas[hoja_dest]
        ws = h_data["ws"]
        id_venta = limpiar_str(row.get(col_df_id_venta, ""))
        if not id_venta:
            continue

        estado_nuevo = limpiar_str(row.get(col_df_estado, ""))

        if id_venta in h_data["existentes"]:
            r_dest = h_data["existentes"][id_venta]
            estado_actual = limpiar_str(ws.cell(r_dest, h_data["col_est"]).value)
            if estado_nuevo and estado_nuevo.lower() != estado_actual.lower():
                if not simulacion:
                    for norm_h, c_idx in h_data["cols_maestro"].items():
                        col_df_match = mapa_cols_df.get(norm_h)
                        if col_df_match and col_df_match in row.index and pd.notna(row[col_df_match]):
                            val = row[col_df_match]
                            val_clean = val.item() if hasattr(val, "item") else val
                            if norm_h == "venta final":
                                m_num = parsear_monto(val_clean)
                                if m_num is not None:
                                    val_clean = m_num
                            val_clean = _limpiar_guion_celda(val_clean)
                            ws.cell(r_dest, c_idx).value = val_clean
                h_data["actualizados"] += 1
        else:
            h_data["ultima_fila"] += 1
            r_dest = h_data["ultima_fila"]
            h_data["existentes"][id_venta] = r_dest
            if not simulacion:
                for norm_h, c_idx in h_data["cols_maestro"].items():
                    col_df_match = mapa_cols_df.get(norm_h)
                    if col_df_match and col_df_match in row.index and pd.notna(row[col_df_match]):
                        val = row[col_df_match]
                        val_clean = val.item() if hasattr(val, "item") else val
                        if norm_h == "venta final":
                            m_num = parsear_monto(val_clean)
                            if m_num is not None:
                                val_clean = m_num
                        val_clean = _limpiar_guion_celda(val_clean)
                        ws.cell(r_dest, c_idx).value = val_clean

                clonar_estilo_fila(
                    ws,
                    h_data["fila_modelo"],
                    r_dest,
                    h_data["max_col"],
                    proteger_relleno_hasta_col=3,
                )
                if h_data["fila_modelo"] is None:
                    h_data["fila_modelo"] = r_dest
            h_data["nuevos"] += 1

    nuevos_total = 0
    for h_nom, h_data in info_hojas.items():
        if not simulacion:
            asegurar_formato_condicional_culqi(h_data["ws"])
        if h_data["nuevos"] > 0 or h_data["actualizados"] > 0:
            resumen.append(
                f"• {h_nom}: {h_data['nuevos']} nuevos, {h_data['actualizados']} actualizados (aprobada → abonada)."
            )
        nuevos_total += h_data["nuevos"]

    return nuevos_total


def _procesar_mpago(wb, df_raw, resumen, simulacion):
    if "MPago" not in wb.sheetnames:
        return 0

    header_idx = None
    for idx in range(min(15, len(df_raw))):
        fila_vals = [normalizar_clave_col(x) for x in df_raw.iloc[idx].tolist()]
        if "numero de movimiento" in fila_vals and "tipo de operacion" in fila_vals:
            header_idx = idx
            break
    if header_idx is None:
        return 0

    headers = [limpiar_str(x) for x in df_raw.iloc[header_idx].tolist()]
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    mapa_cols_df = {normalizar_clave_col(c): c for c in headers if c}
    col_df_mov = mapa_cols_df.get("numero de movimiento", "Número de Movimiento")

    ws = wb["MPago"]
    cols_maestro = {}
    max_col_mp = max(ws.max_column, 12)
    col_mov_maestro = 3  # Columna C: Número de Movimiento

    for c in range(1, ws.max_column + 1):
        h_val = limpiar_str(ws.cell(1, c).value)
        if h_val:
            norm_h = normalizar_clave_col(h_val)
            cols_maestro[norm_h] = c
            if norm_h == "numero de movimiento":
                col_mov_maestro = c

    existentes = set()
    ultima_fila = 1
    for r in range(2, ws.max_row + 1):
        val_mov = limpiar_str(ws.cell(r, col_mov_maestro).value)
        if val_mov:
            existentes.add(val_mov)
            ultima_fila = r

    fila_modelo = ultima_fila if ultima_fila >= 2 else None
    nuevos_mp = 0

    for _, row in df.iterrows():
        mov_id = limpiar_str(row.get(col_df_mov, ""))
        if not mov_id or mov_id in existentes:
            continue

        existentes.add(mov_id)
        ultima_fila += 1
        if not simulacion:
            for norm_h, c_idx in cols_maestro.items():
                col_df_match = mapa_cols_df.get(norm_h)
                if col_df_match and col_df_match in row.index and pd.notna(row[col_df_match]):
                    val = row[col_df_match]
                    val_clean = val.item() if hasattr(val, "item") else val
                    if norm_h == "importe":
                        m_num = parsear_monto(val_clean)
                        if m_num is not None:
                            val_clean = m_num
                    ws.cell(ultima_fila, c_idx).value = val_clean

            clonar_estilo_fila(
                ws,
                fila_modelo,
                ultima_fila,
                max_col_mp,
                proteger_relleno_hasta_col=0,
            )
            if fila_modelo is None:
                fila_modelo = ultima_fila
        nuevos_mp += 1

    if not simulacion:
        asegurar_formato_condicional_mpago(ws)

    if nuevos_mp > 0:
        resumen.append(f"• MPago: {nuevos_mp} movimientos nuevos agregados.")

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
                normalizar_clave_col(x) for x in df_raw.iloc[:12].values.flatten() if pd.notna(x)
            )
            if ("id venta" in muestra or "id de venta" in muestra) and ("comercio" in muestra or "nombre comercial" in muestra):
                total_nuevos += _procesar_culqi(wb, df_raw, empresa_objetivo, resumen, simulacion)
            elif "numero de movimiento" in muestra and "operacion relacionada" in muestra:
                total_nuevos += _procesar_mpago(wb, df_raw, resumen, simulacion)

    if not resumen:
        resumen.append("No se encontraron transacciones nuevas para esta empresa en los archivos subidos.")

    if simulacion:
        return None, total_nuevos, resumen

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), total_nuevos, resumen
