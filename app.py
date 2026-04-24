"""
Audico EC → Peachtree Quote Converter — Web App
================================================

Streamlit front-end around convert.py. Deploy to Streamlit Community Cloud (free)
for a browser-based tool accounting can use from any computer.

Run locally:   streamlit run app.py
"""

import io
import json
import shutil
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from convert import run_conversion

# ────────────────────────────────────────────────────────────────────────────
# Page config
# ────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Audico EC → Peachtree",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Audico EC → Peachtree")
st.caption("Convierte el archivo mensual EC en CSVs listos para subir a Peachtree/Sage 50")


# ────────────────────────────────────────────────────────────────────────────
# Sidebar: configuration & help
# ────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("ℹ️ Cómo usar")
    st.markdown(
        """
        1. **Sube el archivo EC** del mes (ejemplo: `EC_03_-_Marzo_2026.xlsx`)
        2. *(Opcional)* Sube la lista de clientes más reciente de Peachtree
        3. Dale a **Convertir**
        4. Revisa el resumen abajo
        5. Descarga el **ZIP con todos los CSVs** e impórtalos a Peachtree

        Peachtree genera automáticamente:
        - Número de cotización
        - ITBMS (7%)

        Tú solo tienes que **revisar y convertir cotización → factura** dentro de Peachtree.
        """
    )
    st.divider()

    st.header("⚙️ Opciones avanzadas")
    override_date = st.date_input(
        "Fecha de cotización",
        value=date.today(),
        help="Por defecto: hoy. Todas las cotizaciones llevarán esta fecha.",
    )

    with st.expander("Ver mapeo de hoteles"):
        try:
            config_path = Path(__file__).parent / "config" / "hotel_mapping.json"
            with open(config_path) as f:
                cfg = json.load(f)
            mapping_rows = [
                {
                    "Pestaña": tab,
                    "Peachtree ID": t["customer_id"],
                    "Nombre": t["customer_name"],
                }
                for tab, t in cfg["tabs"].items()
            ]
            st.dataframe(pd.DataFrame(mapping_rows), hide_index=True, use_container_width=True)
        except Exception as e:
            st.error(f"No se pudo leer el mapeo: {e}")


# ────────────────────────────────────────────────────────────────────────────
# Main form
# ────────────────────────────────────────────────────────────────────────────

st.subheader("1. Sube el archivo EC del mes")
ec_file = st.file_uploader(
    "Archivo EC (.xlsx)",
    type=["xlsx"],
    help="El archivo mensual con una pestaña por hotel. Ejemplo: EC_03_-_Marzo_2026.xlsx",
    key="ec_upload",
)

st.subheader("2. (Opcional) Lista de clientes actualizada")
customer_file = st.file_uploader(
    "LISTA_DE_CLIENTES_POR_ID.xlsx",
    type=["xlsx"],
    help=(
        "Sirve para verificar que los IDs de cliente en Peachtree siguen existiendo. "
        "Si no la subes, la conversión igual funciona — sólo no valida IDs."
    ),
    key="customer_upload",
)

st.subheader("3. Convertir")
run_btn = st.button("🚀 Convertir a cotizaciones Peachtree", type="primary", use_container_width=True)


# ────────────────────────────────────────────────────────────────────────────
# Conversion logic
# ────────────────────────────────────────────────────────────────────────────

def _zip_directory(src_dir: Path) -> bytes:
    """Zip a directory tree and return the bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in src_dir.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(src_dir))
    buf.seek(0)
    return buf.getvalue()


if run_btn:
    if ec_file is None:
        st.error("⚠️ Por favor sube un archivo EC antes de convertir.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        # Write uploaded files to tmp
        ec_path = tmp / ec_file.name
        ec_path.write_bytes(ec_file.getbuffer())

        customer_path = None
        if customer_file is not None:
            customer_path = tmp / customer_file.name
            customer_path.write_bytes(customer_file.getbuffer())

        out_path = tmp / "out"

        # Use the config that ships with the app
        config_path = Path(__file__).parent / "config" / "hotel_mapping.json"

        with st.spinner("Procesando cotizaciones..."):
            try:
                result = run_conversion(
                    xlsx_path=str(ec_path),
                    config_path=str(config_path),
                    out_dir=str(out_path),
                    customer_list_path=str(customer_path) if customer_path else None,
                    run_date=override_date,
                )
            except Exception as e:
                st.error(f"❌ Error al procesar: {e}")
                st.exception(e)
                st.stop()

        # ─── Results ───
        st.success(f"✅ Listo! Se generaron **{result['quotes_written']} cotizaciones**.")

        col1, col2, col3 = st.columns(3)
        col1.metric("Cotizaciones creadas", result["quotes_written"])
        col2.metric("No procesadas", len(result["not_processed"]))
        col3.metric("Advertencias", len(result["warnings"]))

        # Grand totals from summary
        if result["summary_rows"]:
            df_summary = pd.DataFrame(result["summary_rows"])
            df_summary["subtotal"] = df_summary["subtotal"].astype(float)
            df_summary["itbms"] = df_summary["itbms"].astype(float)
            df_summary["total"] = df_summary["total"].astype(float)

            st.divider()
            st.subheader("📈 Totales")
            tc1, tc2, tc3 = st.columns(3)
            tc1.metric("Subtotal", f"${df_summary['subtotal'].sum():,.2f}")
            tc2.metric("ITBMS", f"${df_summary['itbms'].sum():,.2f}")
            tc3.metric("Total general", f"${df_summary['total'].sum():,.2f}")

        # ─── Download button ───
        st.divider()
        st.subheader("📥 Descargar")
        zip_bytes = _zip_directory(out_path)
        st.download_button(
            label=f"⬇️ Descargar {result['quotes_written']} CSVs + reportes (ZIP)",
            data=zip_bytes,
            file_name=f"audico_cotizaciones_{override_date.isoformat()}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

        # ─── Detail tabs ───
        st.divider()
        tab_summary, tab_skipped, tab_warnings = st.tabs(
            ["📋 Resumen", "⏭️ No procesadas", "⚠️ Advertencias"]
        )

        with tab_summary:
            if result["summary_rows"]:
                st.caption("Una fila por cotización. Revisa esto antes de subir a Peachtree.")
                st.dataframe(
                    df_summary,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "subtotal": st.column_config.NumberColumn("Subtotal", format="$%.2f"),
                        "itbms": st.column_config.NumberColumn("ITBMS", format="$%.2f"),
                        "total": st.column_config.NumberColumn("Total", format="$%.2f"),
                    },
                )
            else:
                st.info("No se generaron cotizaciones.")

        with tab_skipped:
            if result["not_processed"]:
                st.caption(
                    "Filas amarillas que NO produjeron cotización (saltadas o sin datos). "
                    "Revisa si alguna debería procesarse manualmente."
                )
                st.dataframe(
                    pd.DataFrame(result["not_processed"]),
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.info("✅ Todas las filas amarillas fueron procesadas.")

        with tab_warnings:
            if result["warnings"]:
                st.caption("Cosas que podrían necesitar revisión manual:")
                for w in result["warnings"]:
                    st.warning(w)
            else:
                st.info("✅ Sin advertencias. Todo se procesó limpiamente.")
