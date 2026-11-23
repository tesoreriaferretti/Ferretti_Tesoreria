"""Logica de negocio: integra reportes del BCP en BCP_Ferretti_2026.xlsm y BCP_Grupo_Ferretti_2026.xlsm."""
from datetime import date, datetime
from io import BytesIO
import re

import openpyxl
import pandas as pd


# Mapeo de las 21 cuentas BCP -> Pestañas exactas de tus dos libros maestros
CUENTAS_A_HOJA = {
    # --- LIBRO 1: BCP_Ferretti_2026.xlsm ---
    # FP STORES SAC
    "1932540464051": ("FP Stores", "193-2540464-0-51 - FP STORES SAC"),
    "1932538503153": ("FP Stores", "193-2538503-1-53 - FP STORES SAC"),

    # FP STORES SURCO SAC
    "1947416839063": ("FP Surco", "194-7416839-0-63 - FP STORES SURCO SAC"),
    "1918034009114": ("FP Surco", "191-8034009-1-14 - FP STORES SURCO SAC"),

    # --- LIBRO 2: BCP_Grupo_Ferretti_2026.xlsm ---
    # ARQUIA SAC
    "1911502397026": ("Arquia", "191-1502397-0-26 - ARQUIA SAC"),
    "1911488895151": ("Arquia", "191-1488895-1-51 - ARQUIA SAC"),

    # ARKSPACE SAC
    "1912113769000": ("Arkspace", "191-2113769-0-00 - ARKSPACE SAC"),
    "1912112574140": ("Arkspace", "191-2112574-1-40 - ARKSPACE SAC"),

    # AGRICOLA SUNFIELDS SAC
    "1912148597098": ("Agricola", "191-2148597-0-98 - AGRICOLA SUNFIELDS SAC"),
    "1912124183103": ("Agricola", "191-2124183-1-03 - AGRICOLA SUNFIELDS SAC"),
    "19340255699120": ("Agricola", "193-40255699-1-20 - AGRICOLA SUNFIELDS SAC"),

    # STORING SAC (4 cuentas: 2 anteriores + 2 nuevas)
    "1939406484088": ("Storing", "193-9406484-0-88 - STORING SAC"),
    "1939408786150": ("Storing", "193-9408786-1-50 - STORING SAC"),
    "1918028671085": ("Storing", "191-8028671-0-85 - STORING S.A.C."),
    "1918028670185": ("Storing", "191-8028670-1-85 - STORING S.A.C."),

    # MELAMINEX SAC
    "1910171089070": ("Melaminex", "191-0171089-0-70 - MELAMINEX SAC"),
    "1931362484173": ("Melaminex", "193-1362484-1-73 - MELAMINEX SAC"),
    "19111824379115": ("Melaminex", "191-11824379-1-15 - MELAMINEX SAC"),

    # EDIFICADORA PARQUE ANCON SAC (EPA)
    "1912022124094": ("EPA", "191-2022124-0-94 - EDIFICADORA PARQUE ANCON SAC"),
    "1919899731190": ("EPA", "191-9899731-1-90 - EDIFICADORA PARQUE ANCON SAC"),

    # BLACKSTONE SERVICIOS INMOBILIARIOS SAC
    "1912201877080": ("Blackstone", "191-2201877-0-80 - BLACKSTONE SERVICIOS INMOBILIARIOS SAC"),
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


def extraer_solo_digitos_cuenta(val):
    if val is None:
        return ""
    parte_cta = str(val).split(" - ")[0].split(".")[0].strip()
    return re.sub(r"[^0-9]", "", parte_cta)


def clave_unica(cuenta_digitos, num_op, monto, fecha_str="", desc=""):
    o = limpiar_op_para_clave(num_op)
    try:
        m = f"{float(monto):.2f}"
    except (ValueError, TypeError):
        m = str(monto).strip()
    if "REGULARIZAC" in str(desc).upper():
        return f"{cuenta_digitos}_{str(fecha_str).strip()}_{o}_{m}_ITF"
    return f"{cuenta_digitos}_{o}_{m}"


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


def resolver_destino(cuenta_digitos):
    if cuenta_digitos in CUENTAS_A_HOJA:
        return CUENTAS_A_HOJA[cuenta_digitos]
    for k_cta, info in CUENTAS_A_HOJA.items():
        if k_cta.lstrip("0") == cuenta_digitos.lstrip("0") and cuenta_digitos != "":
            return info
    return None, None


def procesar(master_bytes, insumos, simulacion=False):
    wb = openpyxl.load_workbook(BytesIO(master_bytes), keep_vba=True)

    hojas_configuradas = {v[0] for v in CUENTAS_A_HOJA.values()}
    hojas_activas = [h for h in wb.sheetnames if h in hojas_configuradas]

    datos_hojas = {}
    rechazados = []
    dup_antiguos = 0
    egresos_antiguos = 0

    # 1) Leer y depurar cada pestaña activa del maestro (Columnas A-J)
    for nombre_hoja in hojas_activas:
        ws = wb[nombre_hoja]
        registros = {}
        filas = []

        for r in range(3, ws.max_row + 1):
            m = ws.cell(r, 4).value  # Columna D (4): Monto
            if m is None:
                continue
            monto = parsear_monto(m)
            if monto is None:
                rechazados.append(f"[{nombre_hoja}] fila {r}: monto no legible ({m!r}), se conservo sin cambios")
                fila_raw = [ws.cell(r, c).value for c in range(1, 11)]
                fila_raw[0] = parsear_fecha_real(fila_raw[0]) or fila_raw[0]
                filas.append(fila_raw)
                continue
            if monto <= 0:
                egresos_antiguos += 1
                continue

            obs = str(ws.cell(r, 10).value or "").strip()  # Columna J (10): OBSERVACIONES
            if "DUPLICADO" in obs.upper():
                dup_antiguos += 1
                continue

            fecha_obj = parsear_fecha_real(ws.cell(r, 1).value)  # Columna A (1): Fecha
            fecha_str = _fecha_a_texto(fecha_obj if fecha_obj else ws.cell(r, 1).value)
            desc = limpiar_texto(ws.cell(r, 3).value)            # Columna C (3): Descripción
            num_op = ws.cell(r, 6).value                          # Columna F (6): Nº operación
            cta_raw = ws.cell(r, 8).value                         # Columna H (8): Numero de Cuenta
            cta_digitos = extraer_solo_digitos_cuenta(cta_raw)

            k = clave_unica(cta_digitos, num_op, monto, fecha_str, desc)
            if k in registros:
                dup_antiguos += 1
                continue

            op8 = formatear_op_8_digitos(num_op)
            fila = [ws.cell(r, c).value for c in range(1, 11)]
            fila[0] = fecha_obj if fecha_obj else ws.cell(r, 1).value
            fila[2] = desc
            fila[5] = op8
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

            cta_digitos = extraer_solo_digitos_cuenta(df_raw.iloc[0, 1])
            hoja_dest, etiqueta_cta = resolver_destino(cta_digitos)

            if not hoja_dest or hoja_dest not in datos_hojas:
                rechazados.append(
                    f"OMITIDO {nombre} / {sheet}: la cuenta '{cta_digitos}' no pertenece a este archivo maestro."
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

                k = clave_unica(cta_digitos, num_op_raw, monto, fecha_str, desc)
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
                            filas[idx][0] = fecha_obj
                        filas[idx][2] = desc
                        filas[idx][5] = op8
                        filas[idx][6] = usuario
                        actualizados += 1
                        resumen_por_empresa[hoja_dest]["actualizados"] += 1
                else:
                    registros[k] = len(filas)
                    filas.append([
                        fecha_obj,      # Col A (1): Fecha
                        None,           # Col B (2): Fecha valuta
                        desc,           # Col C (3): Descripción operación
                        monto,          # Col D (4): Monto
                        sucursal,       # Col E (5): Sucursal - agencia
                        op8,            # Col F (6): Nº operación
                        usuario,        # Col G (7): Usuario
                        etiqueta_cta,   # Col H (8): Numero de Cuenta
                        moneda,         # Col I (9): Moneda
                        "",             # Col J (10): OBSERVACIONES
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

    # 3) Ordenar cronológicamente por Columna A (Fecha) y reescribir
    for nombre_hoja, contenido_hoja in datos_hojas.items():
        ws = wb[nombre_hoja]
        filas_hoja = contenido_hoja["filas"]

        filas_ordenadas = [
            f for _, f in sorted(
                enumerate(filas_hoja),
                key=lambda par: (
                    par[1][0] if isinstance(par[1][0], date) else date.max,
                    par[0],
                ),
            )
        ]

        for r in range(3, ws.max_row + 1):
            for c in range(1, 11):
                ws.cell(r, c).value = None

        for r_idx, fila in enumerate(filas_ordenadas, start=3):
            for c_idx, val in enumerate(fila, start=1):
                celda = ws.cell(r_idx, c_idx)
                celda.value = val
                if c_idx == 1 and isinstance(val, (date, datetime)):
                    celda.number_format = "dd/mm/yyyy"
                elif c_idx == 6:
                    celda.number_format = "@"

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), stats, rechazados
