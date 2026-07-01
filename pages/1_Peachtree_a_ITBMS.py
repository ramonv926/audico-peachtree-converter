"""
Pestaña: Peachtree -> Informe 43 (ITBMS)
========================================

Toma el export crudo de Peachtree de la cuenta 219 "CTA POR PAGAR I.T.B.M.S"
y genera el archivo del Informe 43 de compras listo para la DGI.

Pagina multipagina de Streamlit: aparece en la barra lateral junto al
convertidor EC -> Peachtree (app.py). No modifica app.py.

GRILLA EDITABLE con validacion obligatoria: contabilidad completa los campos
faltantes en el navegador. El boton de descarga del Informe 43 permanece
BLOQUEADO hasta que TODAS las filas esten completas.
"""

import io
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from itbms_convert import (
    detect_periods, run_itbms_conversion, finalize_records, missing_fields,
    vendor_master_delta, write_informe43,
)

MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
            7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}


def _label_period(ym, n):
    try:
        return f"{MESES_ES[int(ym[4:6])]} {ym[:4]} - {n} filas"
    except (ValueError, KeyError):
        return f"{ym} - {n} filas"


st.set_page_config(page_title="Peachtree -> ITBMS", page_icon="\U0001F9FE", layout="wide")
st.title("\U0001F9FE Peachtree -> Informe 43 (ITBMS)")
st.caption("Convierte el Mayor de la cuenta ITBMS de Peachtree en el archivo del Informe 43 listo para la DGI")

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "itbms_vendors.json"
ss = st.session_state

# ---------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Como usar")
    st.markdown(
        """
        1. **Sube el Mayor (.xlsx)** - export de la cuenta **219 CTA POR PAGAR I.T.B.M.S**
        2. Elige el **mes** y dale a **Generar**
        3. **Completa en la grilla** TODO lo que falte (RUC, DV, Tipo, Factura...)
        4. Cuando el aviso este en **verde**, descarga el **Informe 43**
        5. Descarga el **maestro actualizado** y mandaselo a Ramon para que lo guarde
        """
    )
    st.divider()
    with st.expander("Ver maestro de proveedores"):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg = json.load(f)
            st.caption(f"{len(cfg.get('vendors', []))} proveedores en el maestro.")
            st.dataframe(
                pd.DataFrame([{"Nombre": v["nombre"], "RUC": v["ruc"], "DV": v.get("dv", ""),
                               "Tipo": v.get("tipo", ""), "Concepto": v.get("concepto", 1)}
                              for v in cfg.get("vendors", [])]),
                hide_index=True, use_container_width=True,
            )
        except Exception as e:
            st.error(f"No se pudo leer el maestro: {e}")

# ---------------------------------------------------------- 1. upload
st.subheader("1. Sube el Mayor de Peachtree")
gl_file = st.file_uploader(
    "Mayor de la cuenta ITBMS (.xlsx)", type=["xlsx"],
    help="El archivo tal cual lo exporta Peachtree (hoja 'General Ledger' de la cuenta 219).",
    key="gl_upload",
)

# ---------------------------------------------------------- 2. month
st.subheader("2. Elige el mes a presentar")
period = None
if gl_file is not None:
    try:
        months = detect_periods(io.BytesIO(gl_file.getvalue()))
    except Exception as e:
        months = {}
        st.error(f"No se pudo leer el archivo para detectar meses: {e}")
    if months:
        ordered = sorted(months.items())
        default_ym = max(months, key=months.get)
        labels = {ym: _label_period(ym, n) for ym, n in ordered}
        period = st.selectbox(
            "Mes del informe (detectado en el archivo)",
            options=[ym for ym, _ in ordered],
            index=[ym for ym, _ in ordered].index(default_ym),
            format_func=lambda ym: labels[ym],
            help="Solo se incluiran las compras de este mes. El Informe 43 es estrictamente mensual.",
        )
        if len(ordered) > 1:
            st.warning("El archivo contiene **mas de un mes**. Se incluira solo el mes elegido.")
else:
    st.info("Sube el Mayor arriba para detectar el mes automaticamente.")

# ---------------------------------------------------------- 3. generate
st.subheader("3. Generar")
if st.button("Generar Informe 43", type="primary", use_container_width=True):
    if gl_file is None:
        st.error("Sube el Mayor antes de generar.")
        st.stop()
    if not CONFIG_PATH.exists():
        st.error(f"Falta el maestro de proveedores: {CONFIG_PATH}")
        st.stop()
    with st.spinner("Procesando el Mayor..."):
        try:
            with tempfile.TemporaryDirectory() as td:
                gp = Path(td) / gl_file.name
                gp.write_bytes(gl_file.getbuffer())
                result = run_itbms_conversion(str(gp), str(CONFIG_PATH),
                                              out_dir=str(Path(td) / "o"), period=period)
        except Exception as e:
            st.error(f"Error al procesar: {e}")
            st.exception(e)
            st.stop()
    _df = pd.DataFrame(result["editor_rows"]).drop(columns=["estado"], errors="ignore")
    _decl = result["declarant_ruc"]
    _df["falta"] = ["  ".join(missing_fields(r, _decl)) or "OK"
                    for r in _df.to_dict("records")]
    ss["itbms_sig"] = f"{gl_file.name}:{period}"
    ss["itbms_initial"] = _df
    ss["itbms_period"] = result["period"]
    ss["itbms_decl"] = _decl
    ss["itbms_warnings"] = result["warnings"]

# ---------------------------------------------------------- 4. editable grid + gate
if "itbms_initial" in ss:
    st.divider()
    st.subheader("4. Completa TODO lo que falte")
    st.caption(
        "Todos los campos son **obligatorios**: Tipo, RUC/Cedula, DV, Nombre, Factura y Concepto. "
        "**Fecha** e **ITBMS** vienen del Mayor y no se editan (asi el informe siempre cuadra). "
        "La descarga se habilita cuando no falte nada."
    )

    initial = ss["itbms_initial"]
    decl = ss["itbms_decl"]
    period = ss["itbms_period"]

    edited = st.data_editor(
        initial,
        key=f"editor_{ss['itbms_sig']}",
        use_container_width=True,
        num_rows="fixed",
        height=460,
        column_order=["falta", "fecha", "nombre", "factura", "tipo", "ruc", "dv", "concepto", "monto", "itbms"],
        column_config={
            "falta": st.column_config.TextColumn("\u26a0\ufe0f Falta", disabled=True, width="medium",
                                                 help="Campos vacios en esta fila. 'OK' = fila completa."),
            "fecha": st.column_config.TextColumn("Fecha", disabled=True, width="small"),
            "nombre": st.column_config.TextColumn("Nombre / Razon social", width="large", required=True),
            "factura": st.column_config.TextColumn("Factura", width="small", required=True),
            "tipo": st.column_config.SelectboxColumn("Tipo", options=["J", "N", "E"], width="small", required=True),
            "ruc": st.column_config.TextColumn("RUC / Cedula", required=True),
            "dv": st.column_config.TextColumn("DV", width="small", required=True),
            "concepto": st.column_config.SelectboxColumn("Concepto", options=[1, 2, 3, 4, 5, 6, 7],
                                                         width="small", required=True),
            "monto": st.column_config.NumberColumn("Monto", disabled=True, format="$%.2f"),
            "itbms": st.column_config.NumberColumn("ITBMS", disabled=True, format="$%.2f"),
        },
    )

    # ---- validacion EN VIVO sobre lo editado (siempre actual) ----
    records = edited.to_dict("records")
    # refrescar la columna "Falta" (deshabilitada) sin tocar lo que el usuario edito
    ss["itbms_initial"] = ss["itbms_initial"].assign(
        falta=["  ".join(missing_fields(r, decl)) or "OK" for r in records]
    )
    lines, pending = finalize_records(records, declarant_ruc=decl)
    total_itbms = round(sum(l["itbms"] for l in lines), 2)

    pend_rows = []
    for rec in records:
        miss = missing_fields(rec, decl)
        if miss:
            pend_rows.append({
                "fecha": rec.get("fecha", ""), "nombre": rec.get("nombre", ""),
                "factura": rec.get("factura", ""), "itbms": rec.get("itbms", 0),
                "falta": ", ".join(miss),
            })

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas totales", len(lines))
    c2.metric("Por completar", pending,
              delta=None if pending == 0 else f"{pending} pendientes", delta_color="inverse")
    c3.metric("ITBMS total", f"${total_itbms:,.2f}")

    # ---- semaforo de exportacion ----
    ready = (pending == 0)
    if ready:
        st.success("\u2705 La grilla esta COMPLETA. Lista para exportar el Informe 43.")
    else:
        st.error(f"\u26d4 Faltan campos en **{pending} fila(s)**. Completa todo antes de exportar. "
                 "Abajo se listan exactamente que campos faltan en cada fila.")
        st.dataframe(
            pd.DataFrame(pend_rows),
            hide_index=True, use_container_width=True,
            column_config={
                "fecha": "Fecha", "nombre": "Nombre", "factura": "Factura",
                "itbms": st.column_config.NumberColumn("ITBMS", format="$%.2f"),
                "falta": "Campos que faltan",
            },
        )

    # ---- archivos (la descarga del informe se bloquea si no esta listo) ----
    xlsx_bytes = b""
    if ready:
        with tempfile.TemporaryDirectory() as td:
            out_xlsx = write_informe43(lines, Path(td) / f"Informe43_{decl}_{period}.xlsx")
            xlsx_bytes = out_xlsx.read_bytes()
    cfg_delta, added = vendor_master_delta(lines, CONFIG_PATH)
    master_bytes = json.dumps(cfg_delta, ensure_ascii=False, indent=2).encode("utf-8")

    st.divider()
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            f"Descargar Informe 43 ({period})" if ready else "Descargar Informe 43 (completa la grilla primero)",
            data=xlsx_bytes,
            file_name=f"Informe43_{decl}_{period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True, disabled=not ready,
        )
    with d2:
        st.download_button(
            f"Maestro actualizado (+{added} nuevos)", data=master_bytes,
            file_name="itbms_vendors.json", mime="application/json",
            use_container_width=True, disabled=(added == 0),
            help="Mandale este archivo a Ramon para que lo guarde en el repo. "
                 "El proximo mes estos proveedores ya saldran resueltos.",
        )
    st.caption("La fecha va como **texto** AAAAMMDD y el monto como formula, igual que la plantilla aprobada.")

    if ss.get("itbms_warnings"):
        with st.expander("Que se omitio del Mayor (ventas, retenciones, pago al fisco)"):
            for w in ss["itbms_warnings"]:
                st.info(w)

    if st.button("Empezar de nuevo (otro archivo)"):
        for k in ("itbms_initial", "itbms_sig", "itbms_period", "itbms_decl", "itbms_warnings"):
            ss.pop(k, None)
        st.rerun()
