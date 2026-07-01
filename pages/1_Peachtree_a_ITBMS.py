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
    vendor_master_delta, write_informe43, load_declarants,
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
    _self = result["self_rucs"]
    _df["falta"] = ["  ".join(missing_fields(r, _self)) or "OK"
                    for r in _df.to_dict("records")]
    _df["quitar"] = False
    ss["itbms_sig"] = f"{gl_file.name}:{period}"
    ss["itbms_initial"] = _df
    ss["itbms_ver"] = 0
    ss["itbms_period"] = result["period"]
    ss["itbms_decl"] = result["declarant_ruc"]
    ss["itbms_selfrucs"] = _self
    ss["itbms_detected"] = result["declarant_detected"]
    ss["itbms_warnings"] = result["warnings"]

# ---------------------------------------------------------- 4. editable grid + gate
if "itbms_initial" in ss:
    st.divider()
    st.subheader("4. Completa TODO lo que falte")

    # empresa detectada para el nombre del archivo
    _decl_names = {}
    try:
        _decl_names = {d["ruc"]: d["nombre"] for d in load_declarants(CONFIG_PATH)}
    except Exception:
        pass
    _dname = _decl_names.get(ss["itbms_decl"], ss["itbms_decl"])
    if ss.get("itbms_detected"):
        st.info(f"Empresa detectada del Mayor: **{_dname}**  (RUC {ss['itbms_decl']}). "
                "El archivo se nombrara con este RUC.")
    else:
        st.warning(f"No se pudo detectar la empresa en el Mayor; se usara **{_dname}** "
                   f"(RUC {ss['itbms_decl']}) para el nombre del archivo. Si es otra empresa, renombra el archivo al descargarlo.")

    st.caption(
        "Todos los campos son **obligatorios**: Tipo, RUC/Cedula, DV, Nombre, Factura y Concepto. "
        "**Fecha** e **ITBMS** vienen del Mayor y no se editan (asi el informe siempre cuadra). "
        "La descarga se habilita cuando no falte nada."
    )

    initial = ss["itbms_initial"]
    decl = ss["itbms_decl"]
    selfrucs = ss.get("itbms_selfrucs", [decl] if decl else [])
    period = ss["itbms_period"]

    edited = st.data_editor(
        initial,
        key=f"editor_{ss['itbms_sig']}_{ss.get('itbms_ver', 0)}",
        use_container_width=True,
        num_rows="fixed",
        height=460,
        column_order=["quitar", "falta", "fecha", "nombre", "factura", "tipo", "ruc", "dv", "concepto", "monto", "itbms"],
        column_config={
            "quitar": st.column_config.CheckboxColumn("Quitar", width="small", default=False,
                                                      help="Marca y presiona 'Revisar grilla' para EXCLUIR esta fila del informe."),
            "falta": st.column_config.TextColumn("\u26a0\ufe0f Falta", disabled=True, width="medium",
                                                 help="Campos vacios en esta fila al momento de 'Revisar'. 'OK' = fila completa."),
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

    # ---- estado EN VIVO sobre lo editado (siempre actual) ----
    # NO se reasigna ss["itbms_initial"] aqui (eso borraria la edicion en curso).
    # 'Falta' y las filas quitadas se refrescan al presionar 'Revisar grilla'.
    records = edited.to_dict("records")
    kept = [r for r in records if not r.get("quitar")]
    removed = [r for r in records if r.get("quitar")]
    lines, pending = finalize_records(kept, declarant_rucs=selfrucs)
    grid_itbms = round(sum(l["itbms"] for l in lines), 2)
    mayor_itbms = round(sum(float(r.get("itbms") or 0) for r in records), 2)
    removed_itbms = round(sum(float(r.get("itbms") or 0) for r in removed), 2)

    QUITADA = "\U0001F6AB QUITADA - no va en el Informe"

    def _refresh_base(df, recs):
        out = df.copy()
        out["falta"] = [QUITADA if r.get("quitar") else ("  ".join(missing_fields(r, selfrucs)) or "OK")
                        for r in recs]
        return out

    # ---- acciones: Revisar grilla  +  Buscar proveedor ----
    rev_col, _ = st.columns([1, 3])
    with rev_col:
        revisar = st.button("\U0001F504 Revisar grilla", use_container_width=True,
                            help="Revisa TODAS las filas, aplica las que marcaste 'Quitar' y actualiza los totales.")
    if revisar:
        ss["itbms_initial"] = _refresh_base(edited, records)
        ss["itbms_ver"] = ss.get("itbms_ver", 0) + 1
        st.toast("\u2705 Grilla lista." if pending == 0 else f"Aun faltan {pending} fila(s).",
                 icon="\u2705" if pending == 0 else "\u26a0\ufe0f")
        st.rerun()

    with st.expander("\U0001F50E Buscar proveedor en el maestro (autocompleta Nombre / RUC / DV / Tipo)"):
        try:
            _cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
            maestro = [{"nombre": v["nombre"], "ruc": v["ruc"], "dv": v.get("dv", ""), "tipo": v.get("tipo", "")}
                       for v in _cfg.get("vendors", [])]
        except Exception:
            maestro = []
        st.caption("Puedes seguir escribiendo a mano en la grilla; esto es un atajo para no teclear el RUC/DV.")
        row_opts = list(range(len(records)))
        def _rowlabel(i):
            r = records[i]
            mark = "  \u26a0\ufe0f" if missing_fields(r, selfrucs) else ""
            return f"fila {i+1} - {(str(r.get('nombre')) or '(sin nombre)')[:34]} - ${float(r.get('itbms') or 0):.2f}{mark}"
        tgt = st.selectbox("1) Fila a completar", row_opts, format_func=_rowlabel, key="buscar_row")
        vlabels = [f'{v["nombre"]}  -  {v["ruc"]} (DV {v["dv"]})' for v in maestro]
        vsel = st.selectbox("2) Proveedor del maestro (escribe para buscar)", list(range(len(maestro))),
                            format_func=lambda i: vlabels[i], index=None,
                            placeholder="Escribe parte del nombre...", key="buscar_vendor")
        if st.button("Aplicar proveedor a la fila", disabled=(vsel is None or not maestro)):
            v = maestro[vsel]
            recs = edited.to_dict("records")
            recs[tgt].update(nombre=v["nombre"], ruc=v["ruc"], dv=v["dv"], tipo=v["tipo"])
            newdf = pd.DataFrame(recs)[list(edited.columns)]
            ss["itbms_initial"] = _refresh_base(newdf, recs)
            ss["itbms_ver"] = ss.get("itbms_ver", 0) + 1
            st.toast(f'Aplicado a la fila {tgt+1}: {v["nombre"]}', icon="\u2705")
            st.rerun()

    pend_rows = []
    for rec in kept:
        miss = missing_fields(rec, selfrucs)
        if miss:
            pend_rows.append({
                "fecha": rec.get("fecha", ""), "nombre": rec.get("nombre", ""),
                "factura": rec.get("factura", ""), "itbms": rec.get("itbms", 0),
                "falta": ", ".join(miss),
            })

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas en el informe", len(lines))
    c2.metric("Por completar", pending,
              delta=None if pending == 0 else f"{pending} pendientes", delta_color="inverse")
    c3.metric("ITBMS del informe", f"${grid_itbms:,.2f}")

    if removed:
        st.info(f"Se quitaron **{len(removed)} fila(s)** (B/. {removed_itbms:,.2f}). "
                f"ITBMS del informe: **B/. {grid_itbms:,.2f}**  |  Total del Mayor: B/. {mayor_itbms:,.2f}.")

    # ---- semaforo de exportacion ----
    ready = (pending == 0 and len(lines) > 0)
    if ready:
        st.success("\u2705 La grilla esta COMPLETA. Lista para exportar el Informe 43.")
    elif len(lines) == 0:
        st.error("\u26d4 No queda ninguna fila en el informe (quitaste todas). Nada que exportar.")
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
        for k in ("itbms_initial", "itbms_sig", "itbms_ver", "itbms_period", "itbms_decl",
                  "itbms_selfrucs", "itbms_detected", "itbms_warnings", "buscar_row", "buscar_vendor"):
            ss.pop(k, None)
        st.rerun()
