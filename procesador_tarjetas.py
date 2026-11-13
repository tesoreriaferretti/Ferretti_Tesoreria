"""Logica de negocio: integra reportes de Culqi y Mercado Pago por archivo maestro de empresa."""
from io import BytesIO
import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
import pandas as pd

FILL_PLOMO = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
FILL_CELESTE = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
FILL_AMARILLO = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

HOJAS_CULQI = ["Culqi Link MN", "Culqi Link ME", "Culqi POS MN", "Culqi POS ME"]


def detectar_tipo_reporte(file_bytes):
    try:
        df = pd.read_excel(BytesIO(file_bytes), nrows=5)
        cols = [str(c).strip() for c in df.columns]
        if "ID Venta" in cols and "Producto" in cols:
            return "CULQI"
        if "Número de Movimiento" in cols and "Tipo de Operación" in cols:
            return "MERCADO_PAGO"
    except Exception:
        pass
    return "DESCONOCIDO"


def asegurar_formato_condicional_culqi(ws):
    rango = "D5:AK5000"

    # 1. Plomo: Estado rechazada o anulada (Columna AK)
    rule_plomo = FormulaRule(
        formula=['OR($AK5="rechazada",$AK5="anulada")'],
        fill=FILL_PLOMO,
    )

    # 2. Amarillo (Predomina sobre Celeste): Estado aprobada/abonada + Cotización/Pedido en AZ
    rule_amarillo = FormulaRule(
        formula=['AND(OR($AK5="aprobada",$AK5="abonada"),TRIM($AZ5)<>"",NOT(ISBLANK($AZ5)))'],
        fill=FILL_AMARILLO,
    )

    # 3. Celeste: Estado aprobada/abonada + Observaciones en BF + AZ vacía
    rule_celeste = FormulaRule(
        formula=[
            'AND(OR($AK5="aprobada",$AK5="abonada"),TRIM($BF5)<>"",NOT(ISBLANK($BF5)),OR(ISBLANK($AZ5),TRIM($AZ5)=""))'
        ],
        fill=FILL_CELESTE,
    )

    ws.conditional_formatting.add(rango, rule_plomo)
    ws.conditional_formatting.add(rango, rule_amarillo)
    ws.conditional_formatting.add(rango, rule_celeste)


def pertenece_a_empresa(empresa_fila, empresa_objetivo):
    emp = str(empresa_fila).strip().upper()
    if not emp or emp == "NAN":
        return True
    es_surco = "SURCO" in emp
    if empresa_objetivo == "FP_SURCO":
        return es_surco
    return not es_surco


def procesar_tarjetas(wb_bytes, insumos, empresa_objetivo="FP_STORES", simulacion=False):
    wb = openpyxl.load_workbook(BytesIO(wb_bytes), data_only=False)

    mapa_culqi = {}
    for sh_name in HOJAS_CULQI:
        if sh_name in wb.sheetnames:
            ws = wb[sh_name]
            mapa_culqi[sh_name] = {}
            for r in range(5, ws.max_row + 1):
                val = ws.cell(row=r, column=4).value  # Columna D: ID Venta
                if val:
                    mapa_culqi[sh_name][str(val).strip()] = r

    ids_mpago = set()
    if "MPago" in wb.sheetnames:
        ws_mp = wb["MPago"]
        for r in range(2, ws_mp.max_row + 1):
            val = ws_mp.cell(row=r, column=3).value  # Columna C: Número de Movimiento
            if val:
                ids_mpago.add(str(val).strip())

    resumen = []
    total_nuevos = 0

    for nombre_archivo, file_bytes in insumos:
        tipo = detectar_tipo_reporte(file_bytes)

        if tipo == "CULQI":
            df = pd.read_excel(BytesIO(file_bytes))
            cols_culqi = list(df.columns[:51])
            cont_nuevos = 0
            cont_act = 0
            cont_omitidas_emp = 0

            for _, fila in df.iterrows():
                id_venta = str(fila.get("ID Venta", "")).strip()
                if not id_venta or id_venta == "nan":
                    continue

                if not pertenece_a_empresa(fila.get("Empresa", ""), empresa_objetivo):
                    cont_omitidas_emp += 1
                    continue

                producto = str(fila.get("Producto", "")).strip().lower()
                moneda = str(fila.get("Moneda", "")).strip().upper()

                if "link" in producto:
                    hoja_dest = "Culqi Link MN" if moneda == "PEN" else "Culqi Link ME"
                else:
                    hoja_dest = "Culqi POS MN" if moneda == "PEN" else "Culqi POS ME"

                if hoja_dest not in wb.sheetnames:
                    continue

                ws = wb[hoja_dest]

                if id_venta in mapa_culqi[hoja_dest]:
                    fila_existente = mapa_culqi[hoja_dest][id_venta]
                    estado_actual = str(ws.cell(row=fila_existente, column=37).value or "").strip().lower()
                    estado_nuevo = str(fila.get("Estado", "")).strip().lower()

                    if estado_nuevo and estado_actual != estado_nuevo:
                        for col_idx, col_name in enumerate(cols_culqi, start=1):
                            val = fila.get(col_name)
                            ws.cell(row=fila_existente, column=col_idx, value=val if pd.notna(val) else None)
                        cont_act += 1
                else:
                    nueva_fila = ws.max_row + 1
                    for col_idx, col_name in enumerate(cols_culqi, start=1):
                        val = fila.get(col_name)
                        ws.cell(row=nueva_fila, column=col_idx, value=val if pd.notna(val) else None)
                    mapa_culqi[hoja_dest][id_venta] = nueva_fila
                    cont_nuevos += 1

            for sh_name in mapa_culqi:
                asegurar_formato_condicional_culqi(wb[sh_name])

            msg = f"**{nombre_archivo}** (Culqi): {cont_nuevos} nuevas, {cont_act} actualizadas."
            if cont_omitidas_emp > 0:
                msg += f" ({cont_omitidas_emp} omitidas por corresponder a otra razón social)."
            resumen.append(msg)
            total_nuevos += cont_nuevos

        elif tipo == "MERCADO_PAGO":
            df = pd.read_excel(BytesIO(file_bytes))
            cols_mp = ["Fecha de Pago", "Tipo de Operación", "Número de Movimiento", "Operación Relacionada", "Importe"]
            cont_nuevos = 0

            if "MPago" in wb.sheetnames:
                ws = wb["MPago"]
                for _, fila in df.iterrows():
                    num_mov = str(fila.get("Número de Movimiento", "")).strip()
                    if num_mov and num_mov != "nan" and num_mov not in ids_mpago:
                        nueva_fila = ws.max_row + 1
                        for col_idx, col_name in enumerate(cols_mp, start=1):
                            val = fila.get(col_name)
                            ws.cell(row=nueva_fila, column=col_idx, value=val if pd.notna(val) else None)
                        ids_mpago.add(num_mov)
                        cont_nuevos += 1

            resumen.append(f"**{nombre_archivo}** (Mercado Pago): {cont_nuevos} movimientos nuevos.")
            total_nuevos += cont_nuevos

        else:
            resumen.append(f"⚠️ **{nombre_archivo}**: Formato no reconocido.")

    if simulacion:
        return None, total_nuevos, resumen

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), total_nuevos, resumen
