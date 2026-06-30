"""
Pestaña: Peachtree → Informe 43 (ITBMS)
=======================================

Segundo convertidor de la app. Toma el export crudo de Peachtree de la cuenta
219 "CTA POR PAGAR I.T.B.M.S" y genera el archivo del Informe 43 de compras
listo para subir a la DGI.

Es una página multipágina de Streamlit: aparece automáticamente en la barra
lateral junto al convertidor EC → Peachtree (app.py). No modifica app.py.
"""

import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from itbms_convert import run_itbms_conversion

st.set_page_config(page_title="Peachtree → ITBMS", page_icon="🧾", layout="centered")

st.title("🧾 Peachtree → Informe 43 (ITBMS)")
st.caption("Convierte el Mayor de la cuenta ITBMS de Peachtree en el archivo del Informe 43 listo para la DGI")

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "itbms_vendors.json"

# ────────────────────────────────────────────────────────── sidebar
with st.sidebar:
    st.header("ℹ️ Cómo usar")
    st.markdown(
        """
        1. **Sube el Mayor (.xlsx)** — el export de Peachtree de la cuenta
           **219 CTA POR PAGAR I.T.B.M.S** (hoja *General Ledger*)
        2. Dale a **Generar Informe 43**
        3. Revisa las filas **bloqueadas** (proveedores sin RUC)
        4. Descarga el **.xlsx** y súbelo a la DGI

        El sistema automáticamente:
        - Descarta las **ventas** (ITBMS cobrado) y deja solo las **compras**
        - Pone la **fecha como texto** `AAAAMMDD` (sin pelearte con Excel)
        - Calcula el **monto** = ITBMS ÷ 7%
        - Saca el **RUC/DV** del maestro o del detalle de los reembolsos
        """
    )
    st.divider()
    with st.expander("Ver maestro de proveedores"):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg = json.load(f)
            st.caption(f"{len(cfg.get('vendors', []))} proveedores en el maestro. "
                       "Edita `config/itbms_vendors.json` para agregar los que falten.")
            st.dataframe(
                pd.DataFrame([
                    {"Nombre": v["nombre"], "RUC": v["ruc"], "DV": v.get("dv", ""),
                     "Tipo": v.get("tipo", ""), "Concepto": v.get("concepto", 1)}
                    for v in cfg.get("vendors", [])
                ]),
                hide_index=True, use_container_width=True,
            )
        except Exception as e:
            st.error(f"No se pudo leer el maestro: {e}")

# ────────────────────────────────────────────────────────── main form
st.subheader("1. Sube el Mayor de Peachtree")
gl_file = st.file_uploader(
    "Mayor de la cuenta ITBMS (.xlsx)",
    type=["xlsx"],
    help="El archivo tal cual lo exporta Peachtree (hoja 'General Ledger' de la cuenta 219).",
    key="gl_upload",
)

st.subheader("2. Generar")
run_btn = st.button("🧾 Generar Informe 43", type="primary", use_container_width=True)

# ────────────────────────────────────────────────────────── run
if run_btn:
    if gl_file is None:
        st.error("⚠️ Por favor sube el Mayor antes de generar.")
        st.stop()
    if not CONFIG_PATH.exists():
        st.error(f"❌ Falta el maestro de proveedores: {CONFIG_PATH}")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        gl_path = tmp / gl_file.name
        gl_path.write_bytes(gl_file.getbuffer())
        out_path = tmp / "out"

        with st.spinner("Procesando el Mayor..."):
            try:
                result = run_itbms_conversion(
                    xlsx_path=str(gl_path),
                    config_path=str(CONFIG_PATH),
                    out_dir=str(out_path),
                )
            except Exception as e:
                st.error(f"❌ Error al procesar: {e}")
                st.exception(e)
                st.stop()

        n_blocked = result["rows_blocked"]
        if n_blocked:
            st.warning(
                f"✅ Se generaron **{result['rows_written']} filas**. "
                f"🔍 **{n_blocked} filas quedaron bloqueadas** (proveedor sin RUC) — "
                "revisa la pestaña abajo y complétalas antes de presentar."
            )
        else:
            st.success(f"✅ Listo! Se generaron **{result['rows_written']} filas**, todas resueltas.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Filas OK", result["rows_written"])
        c2.metric("Bloqueadas", n_blocked)
        c3.metric("ITBMS total", f"${result['totals']['itbms']:,.2f}")
        c4.metric("Monto total", f"${result['totals']['monto']:,.2f}")

        # download
        st.divider()
        st.subheader("📥 Descargar")
        out_file = Path(result["out_path"])
        st.download_button(
            label=f"⬇️ Descargar Informe 43 ({result['period']})",
            data=out_file.read_bytes(),
            file_name=out_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption("💡 La fecha ya va como **texto** `AAAAMMDD` y el monto como fórmula, "
                   "igual que la plantilla aprobada. Las filas bloqueadas vienen resaltadas en amarillo "
                   "y listadas en la hoja *Revisar* del archivo.")

        # detail tabs
        st.divider()
        tab_resumen, tab_bloqueadas, tab_omitidas = st.tabs(
            ["📋 Resumen", "🔍 Revisión manual", "⏭️ Qué se omitió"]
        )

        with tab_resumen:
            st.caption("Una fila por línea del informe. Las marcadas BLOQUEADA aún no tienen RUC.")
            df = pd.DataFrame(result["summary_rows"])
            st.dataframe(
                df, hide_index=True, use_container_width=True,
                column_config={
                    "monto": st.column_config.NumberColumn("Monto", format="$%.2f"),
                    "itbms": st.column_config.NumberColumn("ITBMS", format="$%.2f"),
                },
            )

        with tab_bloqueadas:
            if result["not_processed"]:
                st.caption(
                    "🔍 Filas que NO se pueden presentar todavía. Casi siempre es un proveedor "
                    "que falta en el maestro — agrégalo en `config/itbms_vendors.json` y vuelve a generar."
                )
                st.dataframe(pd.DataFrame(result["not_processed"]), hide_index=True, use_container_width=True)
            else:
                st.info("✅ Ninguna fila bloqueada — todos los proveedores tienen RUC.")

        with tab_omitidas:
            st.caption("Filas del Mayor que NO son compras y por eso no van al Informe 43:")
            for w in result["warnings"]:
                st.info(w)
