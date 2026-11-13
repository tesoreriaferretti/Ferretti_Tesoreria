"""Logica de negocio: integra los reportes del BCP al archivo maestro por empresa."""
from datetime import date, datetime
from io import BytesIO
import re

import openpyxl
import pandas as pd


# Mapeo de las 19 cuentas BCP -> 9 Pestañas de Empresas
CUENTAS_A_HOJA = {
    # 1. EDIFICADORA PARQUE ANCON SAC
    "1912022124094": "EDIFICADORA_ANCON",   # CTA CTE PEN
    "1919899731190": "EDIFICADORA_ANCON",   # CTA CTE USD

    # 2. BLACKSTONE SERVICIOS INMOBILIARIOS SAC
    "1912201877080": "BLACKSTONE",          # CTA CTE PEN

    # 3. AGRICOLA SUNFIELDS SAC
    "1912148597098": "AGRICOLA_SUNFIELDS",  # CTA CTE PEN
    "1912124183103": "AGRICOLA_SUNFIELDS",  # CTA CTE USD
    "19340255699120": "AGRICOLA_SUNFIELDS", # CTA AHORRO USD

    # 4. MELAMINEX SAC
    "1910171089070": "MELAMINEX",           # CTA CTE PEN
    "1931362484173": "MELAMINEX",           # CTA CTE USD
    "19111824379115": "MELAMINEX",          # CTA AHORRO USD

    # 5. FP STORES SURCO SAC
    "1947416839063": "FP_SURCO",            # CTA CTE PEN
    "1918034009114": "FP_SURCO",            # CTA CTE USD

    # 6. FP STORES SAC (Soporta 'FP_STORES' o 'Movimientos')
    "1932540464051": "FP_STORES",           # CTA CTE PEN
    "1932538503153": "FP_STORES",           # CTA CTE USD

    # 7. ARQUIA SAC
    "1911502397026": "ARQUIA",              # CTA CTE PEN
    "1911488895151": "ARQUIA",              # CTA CTE USD

    # 8. ARKSPACE SAC
    "1912113769000": "ARKSPACE",            # CTA CTE PEN
    "1912112574140": "ARKSPACE",            # CTA CTE USD

    # 9. STORING SAC
    "1939406484088": "STORING",             # CTA CTE PEN
    "1939408786150": "STORING",             # CTA CTE USD
}


def parsear_monto(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^\d.\-]", "", str(val).replace(",", ""))
    try:
        return float(s)
    except ValueError:
        return None


def limpiar_texto(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(val)).strip()


def limpiar_op_para_clave(val):
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return ""
    return str(val).split(".")[0].strip().lstrip("0")


def formatear_op_8_digitos(val):
    if val is None or (not isinstance(val, str) and pd.isna(val)) or str(val).strip() == "":
        return ""
    s = str(val).split(".")[0].strip().lstrip("0")
    return s.zfill(8) if s else "00000000"


def normalizar_cuenta(val):
    if val is None:
        return ""
    s = str(val).split(".")[0].strip()
    return re.sub(r"[^0-9]", "", s)


def clave_unica(cuenta, num_op, monto, fecha_str="", desc=""):
    c = normalizar_cuenta(cuenta)
    o = limpiar_op_para_clave(num_op)
    try:
        m = f"{float(monto):.2f}"
    except (ValueError, TypeError):
        m = str(monto).strip()
    if "REGULARIZAC" in str(desc).upper():
        return f"{c}_{str(fecha_str).strip()}_{o}_{m}_ITF"
    return f"{c}_{o}_{m}"


def parsear_fecha_real(val):
    if val is None or (not isinstance(val, (str, datetime, date)) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    dt = pd.to_datetime(str(val).strip(), errors="coerce", dayfirst=True)
    return dt.date() if pd.notna(dt) else None


def _fecha_a_texto(fecha_val):
    if hasattr(fecha_val, "strftime"):
        return fecha_val.strftime("%d/%m/%Y")
    return str(fecha_val or "").strip()


def resolver_hoja_destino(cuenta_norm, wb_sheetnames):
    hoja = CUENTAS_A_HOJA.get(cuenta_norm)
    if not hoja:
        for k_cta, v_hoja in CUENTAS_A_HOJA.items():
            if k_cta.lstrip("0") == cuenta_norm.lstrip("0") and cuenta_norm != "":
                hoja = v_hoja
                break
    if hoja == "FP_STORES" and "FP_STORES" not in wb_sheetnames and "Movimientos" in wb_sheetnames:
        return "Movimientos"
    return hoja


def procesar(master_bytes, insumos, simulacion=False):
    wb = openpyxl.load_workbook(BytesIO(master_bytes), keep_vba=True)

    hojas_objetivo = set(CUENTAS_A_HOJA.values())
    if "Movimientos" in wb.sheetnames:
        hojas_objetivo.add("Movimientos")

    hojas_activas = [h for h in wb.sheetnames if h in hojas_objetivo]
    datos_hojas = {}
    rechazados = []
    dup_antiguos = 0
    egresos_antiguos = 0

    # 1) Leer pestañas del maestro
    for nombre_hoja in hojas_activas:
        ws = wb[nombre_hoja]
        registros = {}
        filas = []

        for r in range(3, ws.max_row + 1):
            m = ws.cell(r, 5).value
            if m is None:
                continue
            monto = parsear_monto(m)
            if monto is None:
                rechazados.append(f"[{nombre_hoja}] fila {r}: monto no legible ({m!r}), se conservo sin cambios")
                fila_raw = [ws.cell(r, c).value for c in range(1, 12)]
                fila_raw[1] = parsear_fecha_real(fila_raw[1]) or fila_raw[1]
                filas.append(fila_raw)
                continue
            if monto <= 0:
                egresos_antiguos += 1
                continue

            obs = str(ws.cell(r, 11).value or "").strip()
            if "DUPLICADO" in obs.upper():
                dup_antiguos += 1
                continue

            fecha_obj = parsear_fecha_real(ws.cell(r, 2).value)
            fecha_str = _fecha_a_texto(fecha_obj if fecha_obj else ws.cell(r, 2).value)
            desc = limpiar_texto(ws.cell(r, 4).value)
            cuenta = normalizar_cuenta(ws.cell(r, 9).value)
            num_op = ws.cell(r, 7).value

            k = clave_unica(cuenta, num_op, monto, fecha_str, desc)
            if k in registros:
                dup_antiguos += 1
                continue

            op8 = formatear_op_8_digitos(num_op)
            sucursal = limpiar_texto(ws.cell(r, 6).value)
            fila = [ws.cell(r, c).value for c in range(1, 12)]
            fila[0] = f"{fecha_str}-{cuenta}-{op8}-{sucursal}-{desc}-"
            fila[1] = fecha_obj if fecha_obj else ws.cell(r, 2).value
            fila[3] = desc
            fila[6] = op8
            registros[k] = len(filas)
            filas.append(fila)

        datos_hojas[nombre_hoja] = {"registros": registros, "filas": filas}

    # 2) Procesar reportes del banco
    nuevos = 0
    actualizados = 0
    resumen_por_empresa = {}

    for nombre, contenido in insumos:
        try:
            xl = pd.ExcelFile(BytesIO(contenido))
        except Exception as e:
            rechazados.append(f"{nombre}: no se pudo abrir ({e})")
            raise ValueError(f"No se pudo abrir el archivo '{nombre}'. Revisa que sea un Excel valido.")

        for sheet in xl.sheet_names:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            if len(df_raw) < 5:
                continue

            cuenta = normalizar_cuenta(df_raw.iloc[0, 1])
            hoja_dest = resolver_hoja_destino(cuenta, wb.sheetnames)

            if not hoja_dest or hoja_dest not in datos_hojas:
                rechazados.append(
                    f"OMITIDO {nombre} / {sheet}: la cuenta '{cuenta}' no tiene una pestaña configurada en el maestro."
                )
                continue

            registros = datos_hojas[hoja_dest]["registros"]
            filas = datos_hojas[hoja_dest]["filas"]
            resumen_por_empresa.setdefault(hoja_dest, {"nuevos": 0, "actualizados": 0})

            moneda = limpiar_texto(df_raw.iloc[1, 1])
            cols = [str(c).strip() for c in df_raw.iloc[4].tolist()]
            df_data = pd.read_excel(xl, sheet_name=sheet, skiprows=5, header=None)
            vistos_hoja = set()
            es_hist = any("Operación - Número" in c for c in cols)

            for i, fila in df_data.iterrows():
                num_fila = i + 7
                monto = parsear_monto(fila[3])
                if monto is None:
                    if not all(pd.isna(x) for x in fila.tolist()):
                        rechazados.append(f"{nombre} / {sheet} fila {num_fila}: monto no legible ({fila[3]!r})")
                    continue
                if monto <= 0:
                    continue

                fecha_dt = pd.to_datetime(fila[0], errors="coerce", dayfirst=True)
                if pd.isna(fecha_dt):
                    rechazados.append(
                        f"{nombre} / {sheet} fila {num_fila}: fecha no valida ({fila[0]!r}), se agrego sin fecha"
                    )
                    fecha_obj = None
                    fecha_str = ""
                else:
                    fecha_obj = fecha_dt.date()
                    fecha_str = fecha_obj.strftime("%d/%m/%Y")

                desc = limpiar_texto(fila[2])
                sucursal = limpiar_texto(fila[4])
                num_op_raw = fila[5]
                if es_hist and len(fila) > 7:
                    usuario = limpiar_texto(fila[7])
                else:
                    usuario = limpiar_texto(fila[6]) if len(fila) > 6 else ""

                k = clave_unica(cuenta, num_op_raw, monto, fecha_str, desc)
                op8 = formatear_op_8_digitos(num_op_raw)

                if op8 == "":
                    rechazados.append(
                        f"AVISO [{hoja_dest}] {nombre} fila {num_fila}: deposito sin numero de operacion ({monto:.2f})."
                    )
                if k in vistos_hoja:
                    rechazados.append(
                        f"AVISO [{hoja_dest}] {nombre} fila {num_fila}: misma cuenta, operacion y monto repetidos ({monto:.2f})."
                    )
                vistos_hoja.add(k)

                if k in registros:
                    if es_hist:
                        idx = registros[k]
                        if fecha_obj is not None:
                            filas[idx][1] = fecha_obj
                        f_txt = _fecha_a_texto(filas[idx][1])
                        filas[idx][3] = desc
                        filas[idx][6] = op8
                        filas[idx][7] = usuario
                        filas[idx][0] = f"{f_txt}-{cuenta}-{op8}-{filas[idx][5]}-{desc}-"
                        actualizados += 1
                        resumen_por_empresa[hoja_dest]["actualizados"] += 1
                else:
                    concat = f"{fecha_str}-{cuenta}-{op8}-{sucursal}-{desc}-"
                    registros[k] = len(filas)
                    filas.append([
                        concat,
                        fecha_obj,
                        None,
                        desc,
                        monto,
                        sucursal,
                        op8,
                        usuario,
                        cuenta,
                        moneda,
                        "",
                    ])
                    nuevos += 1
                    resumen_por_empresa[hoja_dest]["nuevos"] += 1

    total_global = sum(len(d["filas"]) for d in datos_hojas.values())
    stats = {
        "nuevos_ingresos": nuevos,
        "actualizados_con_historico": actualizados,
        "duplicados_antiguos_eliminados": dup_antiguos,
        "egresos_antiguos_eliminados": egresos_antiguos,
        "filas_con_observaciones": len(rechazados),
        "total_depositos": total_global,
        "detalle_por_empresa": resumen_por_empresa,
    }

    if simulacion:
        return None, stats, rechazados

    # 3) Ordenar cronológicamente por Columna B y reescribir
    for nombre_hoja, contenido_hoja in datos_hojas.items():
        ws = wb[nombre_hoja]
        filas_hoja = contenido_hoja["filas"]

        filas_ordenadas = [
            f for _, f in sorted(
                enumerate(filas_hoja),
                key=lambda par: (
                    par[1][1] if isinstance(par[1][1], date) else date.max,
                    par[0],
                ),
            )
        ]

        for r in range(3, ws.max_row + 1):
            for c in range(1, 12):
                ws.cell(r, c).value = None

        for r_idx, fila in enumerate(filas_ordenadas, start=3):
            for c_idx, val in enumerate(fila, start=1):
                celda = ws.cell(r_idx, c_idx)
                celda.value = val
                if c_idx == 2 and isinstance(val, (date, datetime)):
                    celda.number_format = "dd/mm/yyyy"
                elif c_idx == 7:
                    celda.number_format = "@"

        for col in ["C", "F", "H", "I"]:
            ws.column_dimensions[col].hidden = True

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), stats, rechazados
