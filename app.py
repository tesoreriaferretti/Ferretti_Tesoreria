import pandas as pd
import streamlit as st
from drive_utils import actualizar, copiar, descargar, subir_nuevo
from procesador import procesar
from procesador_tarjetas import procesar_tarjetas

st.set_page_config(page_title="Tesorería Grupo Ferretti", page_icon="🏦", layout="wide")

modulo = st.sidebar.radio(
    "Módulo de Tesorería:",
    [
        "🏪 Bancos Retail (FP Stores / Surco)",
        "🏢 Bancos Grupo (7 Empresas)",
        "💳 Tarjetas (Culqi / Mercado Pago)",
    ],
)

d = st.secrets["drive"]


def render_modulo_bancos(clave_secret, nombre_archivo_excel, titulo_pagina, prefix_key):
    st.title(titulo_pagina)
    st.write("1. Sube los reportes del banco (históricos o diarios). 2. Revisa la simulación. 3. Confirma.")
    st.info(f"Antes de procesar, cierra el archivo '{nombre_archivo_excel}' en Excel.")

    target_id = d.get(clave_secret, "")
    if not target_id:
        st.warning(f"Falta configurar '{clave_secret}' en los Secrets de Streamlit.")

    archivos = st.file_uploader(
        "Reportes del BCP (.xls / .xlsx)",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key=f"up_{prefix_key}",
    )

    if st.button("1) Ver simulación (no guarda nada)", key=f"sim_{prefix_key}"):
        if not archivos or not target_id:
            st.warning("Selecciona al menos un archivo y verifica los Secrets.")
        else:
            with st.spinner(f"Descargando {nombre_archivo_excel} y analizando cuentas..."):
                try:
                    master_bytes = descargar(target_id)
                    insumos = [(a.name, a.getvalue()) for a in archivos]
                    _, stats, rechazados = procesar(master_bytes, insumos, simulacion=True)

                    st.session_state[f"insumos_{prefix_key}"] = insumos
                    st.session_state[f"stats_{prefix_key}"] = stats
                    st.session_state[f"rechazados_{prefix_key}"] = rechazados
                except Exception as e:
                    st.error(f"Error en simulación: {e}")

    if f"stats_{prefix_key}" in st.session_state:
        stats = st.session_state[f"stats_{prefix_key}"]
        rechazados = st.session_state[f"rechazados_{prefix_key}"]

        st.success("Simulación completada.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Nuevos ingresos", stats["nuevos_ingresos"])
        c2.metric("Actualizados c/ histórico", stats["actualizados_con_historico"])
        c3.metric("Total registros en libro", stats["total_depositos"])
        c4.metric("Observaciones / Omitidos", stats["filas_con_observaciones"])

        if stats.get("detalle_por_empresa"):
            st.subheader("Detalle por empresa")
            st.table(pd.DataFrame.from_dict(stats["detalle_por_empresa"], orient="index"))

        if rechazados:
            with st.expander(f"Ver {len(rechazados)} observaciones o cuentas omitidas"):
                for r in rechazados:
                    st.write(f"- {r}")

        if st.button("2) Confirmar y guardar en Google Drive", type="primary", key=f"save_{prefix_key}"):
            with st.spinner("Generando respaldo, ordenando fechas y guardando en Drive..."):
                try:
                    master_bytes = descargar(target_id)
                    nuevo_xlsm, _, _ = procesar(
                        master_bytes, st.session_state[f"insumos_{prefix_key}"], simulacion=False
                    )

                    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                    copiar(target_id, f"Backup_{nombre_archivo_excel}_{ts}.xlsm", d["backups_id"])
                    actualizar(
                        target_id,
                        nuevo_xlsm,
                        "application/vnd.ms-excel.sheet.macroEnabled.12",
                    )
                    for nombre_arch, cont_arch in st.session_state[f"insumos_{prefix_key}"]:
                        subir_nuevo(f"{ts}_{nombre_arch}", cont_arch, d["procesados_id"])

                    st.success(f"¡Archivo '{nombre_archivo_excel}' actualizado y ordenado por fecha en Drive!")
                    del st.session_state[f"stats_{prefix_key}"]
                    del st.session_state[f"insumos_{prefix_key}"]
                except Exception as e:
                    st.error(f"Error al guardar: {e}")


# ==============================================================================
# RUTEO POR MÓDULO
# ==============================================================================
if modulo == "🏪 Bancos Retail (FP Stores / Surco)":
    render_modulo_bancos(
        clave_secret="bcp_retail_id",
        nombre_archivo_excel="BCP_Ferretti_2026.xlsm",
        titulo_pagina="🏪 Bancos BCP — FP Stores & FP Surco",
        prefix_key="retail",
    )

elif modulo == "🏢 Bancos Grupo (7 Empresas)":
    render_modulo_bancos(
        clave_secret="bcp_grupo_id",
        nombre_archivo_excel="BCP_Grupo_Ferretti_2026.xlsm",
        titulo_pagina="🏢 Bancos BCP — Grupo Ferretti (7 Empresas)",
        prefix_key="grupo",
    )

elif modulo == "💳 Tarjetas (Culqi / Mercado Pago)":
    st.title("💳 Procesar pagos con Tarjeta (Culqi y Mercado Pago)")

    empresa_sel = st.selectbox(
        "Empresa a actualizar:",
        ["FP STORES SAC", "FP STORES SURCO SAC"],
        key="emp_tarjetas",
    )

    if empresa_sel == "FP STORES SAC":
        tarjetas_id = d.get("tarjetas_master_id", "")
        emp_code = "FP_STORES"
        nombre_excel = "FP Tarjetas 2026.xlsx"
    else:
        tarjetas_id = d.get("tarjetas_surco_id", "")
        emp_code = "FP_SURCO"
        nombre_excel = "FP Surco Tarjetas 2026.xlsx"

    st.info(f"Antes de procesar, cierra el archivo '{nombre_excel}' en Excel.")

    if not tarjetas_id:
        st.warning(f"Falta configurar el ID maestro para {empresa_sel} en los Secrets.")

    archivos_tarj = st.file_uploader(
        "Reportes de Culqi o Mercado Pago (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="up_tarj",
    )

    if st.button("1) Ver simulación Tarjetas", key="sim_tarj"):
        if not archivos_tarj or not tarjetas_id:
            st.warning("Verifica que hayas subido archivos y que la empresa tenga un maestro activo.")
        else:
            with st.spinner("Analizando operaciones de tarjetas..."):
                try:
                    master_bytes = descargar(tarjetas_id)
                    insumos_t = [(a.name, a.getvalue()) for a in archivos_tarj]
                    _, total_nuevos, resumen = procesar_tarjetas(
                        master_bytes, insumos_t, empresa_objetivo=emp_code, simulacion=True
                    )

                    st.session_state["tarj_insumos"] = insumos_t
                    st.session_state["tarj_total"] = total_nuevos
                    st.session_state["tarj_resumen"] = resumen
                    st.session_state["tarj_emp_code"] = emp_code
                    st.session_state["tarj_target_id"] = tarjetas_id
                    st.session_state["tarj_nombre_excel"] = nombre_excel
                except Exception as e:
                    st.error(f"Error en simulación: {e}")

    if "tarj_resumen" in st.session_state:
        st.success(f"Simulación lista: {st.session_state['tarj_total']} transacciones nuevas detectadas.")
        for linea in st.session_state["tarj_resumen"]:
            st.write(linea)

        if st.button("2) Confirmar y guardar en Drive", type="primary", key="save_tarj"):
            with st.spinner("Guardando base de tarjetas y respaldando..."):
                try:
                    target_id = st.session_state["tarj_target_id"]
                    master_bytes = descargar(target_id)
                    nuevo_xlsx, _, _ = procesar_tarjetas(
                        master_bytes,
                        st.session_state["tarj_insumos"],
                        empresa_objetivo=st.session_state["tarj_emp_code"],
                        simulacion=False,
                    )

                    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                    copiar(target_id, f"Backup_{st.session_state['tarj_nombre_excel']}_{ts}.xlsx", d["backups_id"])
                    actualizar(
                        target_id,
                        nuevo_xlsx,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    for nombre_arch, cont_arch in st.session_state["tarj_insumos"]:
                        subir_nuevo(f"{ts}_{nombre_arch}", cont_arch, d["procesados_id"])

                    st.success(f"¡Base '{st.session_state['tarj_nombre_excel']}' actualizada correctamente en Google Drive!")
                    del st.session_state["tarj_resumen"]
                    del st.session_state["tarj_insumos"]
                except Exception as e:
                    st.error(f"Error al guardar: {e}")
