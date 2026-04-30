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
        1. **Sube el archivo .xlsx** (EC mensual o Dreams Plaza — el sistema detecta el tipo automáticamente)
        2. Dale a **Convertir**
        3. Revisa el resumen abajo
        4. Descarga el **archivo único** e impórtalo a Peachtree

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

st.subheader("1. Sube el archivo")
ec_file = st.file_uploader(
    "Archivo a procesar (.xlsx)",
    type=["xlsx"],
    help=(
        "El sistema detecta automáticamente el tipo de archivo:\n"
        "• **EC mensual** (varias pestañas, una por hotel)\n"
        "• **Dreams Plaza** (una sola pestaña, cliente directo DR PROPERTY SERVICES CORP)"
    ),
    key="ec_upload",
)

st.subheader("2. Convertir")
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
        # Write uploaded file to tmp
        ec_path = tmp / ec_file.name
        ec_path.write_bytes(ec_file.getbuffer())

        out_path = tmp / "out"

        # Use the config that ships with the app
        config_path = Path(__file__).parent / "config" / "hotel_mapping.json"

        with st.spinner("Procesando cotizaciones..."):
            try:
                result = run_conversion(
                    xlsx_path=str(ec_path),
                    config_path=str(config_path),
                    out_dir=str(out_path),
                    run_date=override_date,
                )
            except Exception as e:
                st.error(f"❌ Error al procesar: {e}")
                st.exception(e)
                st.stop()

        # ─── Results ───
        # Show detected file type prominently
        ft = result.get("file_type", "EC")
        if ft == "DREAMS_PLAZA":
            st.info(
                "📄 **Tipo de archivo detectado: Dreams Plaza** — "
                "una sola pestaña, cliente fijo (DR PROPERTY SERVICES CORP), Audico al 70%."
            )
        else:
            st.info(
                "📄 **Tipo de archivo detectado: EC mensual** — "
                "múltiples hoteles, Audico al 50%."
            )

        n_review = len(result.get("review_flags", []))
        if n_review > 0:
            st.success(
                f"✅ Listo! Se generaron **{result['quotes_written']} cotizaciones**. "
                f"🔍 **{n_review} elementos requieren revisión manual** (ver pestaña abajo)."
            )
        else:
            st.success(f"✅ Listo! Se generaron **{result['quotes_written']} cotizaciones**.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Cotizaciones creadas", result["quotes_written"])
        col2.metric("No procesadas", len(result["not_processed"]))
        col3.metric("Revisión manual", n_review)
        col4.metric("Advertencias", len(result["warnings"]))

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

        # PRIMARY: combined CSV for bulk import — accounting just imports this ONE file
        combined_path = out_path / "_TODAS_LAS_COTIZACIONES.csv"
        if combined_path.exists():
            st.markdown(
                "**⭐ Recomendado: archivo único para importación masiva.** "
                "Sube este archivo a Sage 50 una sola vez y todas las cotizaciones se cargan juntas."
            )
            st.download_button(
                label=f"⬇️ Descargar TODO en un solo CSV ({result['quotes_written']} cotizaciones)",
                data=combined_path.read_bytes(),
                file_name=f"audico_TODAS_cotizaciones_{override_date.isoformat()}.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True,
            )
            st.caption(
                "💡 Cada cotización lleva un ID temporal (TMP-0001, TMP-0002, etc.) "
                "para que Sage las separe correctamente. Sage les asignará números reales al convertirlas en facturas."
            )
            st.markdown("---")

        # SECONDARY: full ZIP with individual files + audit reports
        st.markdown("**Alternativa: ZIP con archivos individuales + reportes de auditoría.**")
        zip_bytes = _zip_directory(out_path)
        st.download_button(
            label=f"⬇️ Descargar ZIP completo ({result['quotes_written']} CSVs individuales + reportes)",
            data=zip_bytes,
            file_name=f"audico_cotizaciones_{override_date.isoformat()}.zip",
            mime="application/zip",
            use_container_width=True,
        )

        # ─── Detail tabs ───
        st.divider()
        tab_summary, tab_skipped, tab_review, tab_warnings = st.tabs(
            ["📋 Resumen", "⏭️ No procesadas", "🔍 Revisión manual", "⚠️ Advertencias"]
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

        with tab_review:
            review_flags = result.get("review_flags", [])
            if review_flags:
                st.caption(
                    "🔍 Elementos detectados que requieren revisión humana — descuentos, créditos, "
                    "notas a contabilidad y otros casos ambiguos. **El sistema NO los incluye en las "
                    "cotizaciones automáticamente.** Decide caso por caso cómo manejarlos en Peachtree."
                )
                st.dataframe(
                    pd.DataFrame(review_flags),
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.info("✅ Sin elementos que requieran revisión manual.")

        with tab_warnings:
            if result["warnings"]:
                st.caption("Cosas que podrían necesitar revisión manual:")
                for w in result["warnings"]:
                    st.warning(w)
            else:
                st.info("✅ Sin advertencias. Todo se procesó limpiamente.")
