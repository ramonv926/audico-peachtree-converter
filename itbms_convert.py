"""
Peachtree General Ledger (cuenta ITBMS por pagar)  ->  Informe 43 (DGI)
=======================================================================

Segundo convertidor de la app. Toma el export crudo de Peachtree de la cuenta
219 "CTA POR PAGAR I.T.B.M.S" (hoja 'General Ledger') y produce el archivo del
Informe 43 de compras que se sube a la DGI a fin de mes.

Sigue las convenciones de convert.py: carga config desde config/*.json, devuelve
un dict de resultado estructurado y escribe los archivos de salida en out_dir.

Lógica (descubierta a partir de COMPRAS_MAYO_2026 + Informe43 de abril, que
reconcilian exactamente: 53 filas, B/.1504.36):

  FILTRO — qué filas del General Ledger son compras del Informe 43
    • Solo filas con DÉBITO. Los CRÉDITOS (Jrnl 'SJ') son el ITBMS *cobrado*
      sobre las VENTAS de la empresa -> no son compras, no van al informe.
    • Se excluyen retenciones (Reference/Trans empieza con 'RETEN'/'RETENCION'),
      el pago al fisco ('TESORO NACIONAL' / 'PAGOS DE IMPUESTOS' /
      'ITBMS POR PAGAR') y las filas resumen (Beginning/Current/Ending Balance).

  TRANSFORMACIÓN por fila
    ITBMS PAGADO  = Debit Amt
    MONTO BALBOAS = ITBMS / 0.07          (fórmula =J*14.2857142857143 como la plantilla)
    FECHA         = texto 'YYYYMMDD'       (¡texto!, no fecha — esto resuelve el dolor de cabeza)
    FACTURA       = Reference; en reembolsos se extrae 'FACT <n>' del detalle
    NOMBRE/RUC/DV = del detalle (reembolsos) o del MAESTRO config/itbms_vendors.json
    TIPO PERSONA  = inferido del formato de cédula/RUC (N / J / E)
    CONCEPTO      = del maestro / concepto_overrides; 1 por defecto
    COMPRAS B/S   = 1 (Locales) por defecto

  Las filas cuyo proveedor no está en el maestro (o reembolsos sin RUC en el
  texto) salen BLOQUEADAS y se listan en la hoja 'Revisar' + en el dict de
  resultado, igual que la pestaña CODE del sistema de pagos. Nada se inventa.
"""
from __future__ import annotations

import json
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ───────────────────────────────────────────────────────── constants
DEFAULT_ITBMS_RATE = 0.07
SUMMARY_ROWS = {"Beginning Balance", "Current Period Change", "Ending Balance"}
GL_SHEET_CANDIDATES = ("General Ledger", "General ledger", "GENERAL LEDGER")
BLOCKED_FILL = PatternFill("solid", fgColor="FFF3B0")

INFORME43_HEADERS = [
    "TIPO DE PERSONA", "RUC ", "DV ", "NOMBRE O RAZON SOCIAL", "FACTURA",
    "FECHA", "CONCEPTO", "COMPRAS DE BIENES Y SERVICIOS",
    "MONTO EN BALBOAS", "ITBMS PAGADO EN BALBOAS",
]

# ───────────────────────────────────────────────────────── patterns
RUC_RE = re.compile(r'(\d{3,}-\d{1,3}-\d{2,}|\d{6,}-\d-\d{3,}|[NEP]{1,2}-\d+-\d+)')
DV_RE = re.compile(r'DV\s*(\d{1,3})', re.I)
FACT_RE = re.compile(r'FAC?T\s*([A-Z0-9\-]+)', re.I)          # tolera el typo 'FCAT'
ITBMS_SUFFIX_RE = re.compile(r'\s*-\s*IT[MB]{2,3}S?\s*$', re.I)  # ' - ITBMS' / ' - ITMBS'


# ───────────────────────────────────────────────────────── helpers
def norm_name(s) -> str:
    """Normaliza un nombre de proveedor para emparejar: mayúsculas, sin acentos,
    sin puntuación, sin sufijos legales ni el ' - ITBMS' que pega Peachtree."""
    s = ITBMS_SUFFIX_RE.sub('', str(s or '')).strip()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()
    s = s.upper()
    s = re.sub(r'[.,]', ' ', s)
    s = re.sub(r'\b(S\s?A|INC|CORP|S\s?DE\s?RL)\b', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def infer_tipo(ruc) -> str:
    """N natural (cédula), E extranjero, J jurídico (empresa)."""
    r = (ruc or '').strip().upper()
    if r.startswith('E-'):
        return 'E'
    if r.startswith('N-') or r.startswith('PE-'):
        return 'N'
    if re.fullmatch(r'\d{1,2}-\d{1,4}-\d{1,5}', r):
        return 'N'
    return 'J'


def date_to_text(d) -> str:
    """A texto 'YYYYMMDD'. El General Ledger entrega datetime; si llegara texto,
    se intentan formatos comunes."""
    if isinstance(d, (datetime, date)):
        return d.strftime('%Y%m%d')
    s = str(d).strip()
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s[:10], fmt).strftime('%Y%m%d')
        except ValueError:
            pass
    return re.sub(r'\D', '', s)[:8]


def is_reten(ref, desc) -> bool:
    t = f"{ref or ''} {desc or ''}".upper().strip()
    return t.startswith('RETEN') or 'RETENCION' in t


def is_treasury(desc) -> bool:
    d = (desc or '').upper()
    return any(k in d for k in ('TESORO NACIONAL', 'PAGOS DE IMPUESTOS', 'ITBMS POR PAGAR'))


def is_reimbursement(ref) -> bool:
    return (ref or '').upper().startswith('REEMB')


def parse_reimbursement(desc: str):
    """De un blob de reembolso -> (nombre, ruc, dv, factura)."""
    m_ruc, m_dv, m_f = RUC_RE.search(desc), DV_RE.search(desc), FACT_RE.search(desc)
    nombre = (desc[:m_ruc.start()] if m_ruc else desc).strip()
    ruc = m_ruc.group(1) if m_ruc else ''
    dv = m_dv.group(1).zfill(2) if m_dv else ''
    fac = m_f.group(1) if m_f else ''
    return nombre, ruc, dv, fac


# ───────────────────────────────────────────────────────── vendor master
def load_vendors(config_path):
    """Carga config/itbms_vendors.json. Devuelve (master, overrides, declarant, rate).
    master: {norm_name: {ruc,dv,tipo,nombre,concepto}}. Acepta 'aliases' opcionales."""
    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)

    master = {}
    for v in cfg.get("vendors", []):
        rec = {
            "nombre": v["nombre"].strip(),
            "ruc": str(v["ruc"]).strip(),
            "dv": str(v.get("dv", "")).strip().zfill(2) if v.get("dv") else "",
            "tipo": v.get("tipo") or infer_tipo(v["ruc"]),
            "concepto": int(v.get("concepto", 1)),
        }
        master[norm_name(v["nombre"])] = rec
        for alias in v.get("aliases", []):
            master[norm_name(alias)] = rec

    overrides = {norm_name(k): int(val) for k, val in cfg.get("concepto_overrides", {}).items()}
    declarant = (cfg.get("declarant") or {}).get("ruc")
    rate = float(cfg.get("itbms_rate", DEFAULT_ITBMS_RATE))
    return master, overrides, declarant, rate


def lookup_master(name: str, master: dict):
    """Match exacto; luego match por prefijo para sobrevivir la truncación de
    nombres de Peachtree (~30 chars), p.ej. 'MEETING & SHOW TECHNOLOGIES, S'."""
    key = norm_name(name)
    if key in master:
        return master[key]
    if len(key) >= 12:
        cands = [v for k, v in master.items() if k.startswith(key) or key.startswith(k)]
        if len(cands) == 1:
            return cands[0]
    return None


# ───────────────────────────────────────────────────────── core transform
def _gl_sheet(wb):
    for name in GL_SHEET_CANDIDATES:
        if name in wb.sheetnames:
            return wb[name]
    return wb[wb.sheetnames[0]]  # fallback: primera hoja


def transform_rows(gl_rows, master, overrides, declarant_ruc, rate):
    """gl_rows: tuplas (acct_id, acct_desc, date, ref, jrnl, tdesc, debit, credit, balance).
    Devuelve (lines, stats). Cada line es un dict listo para escribir / mostrar."""
    lines = []
    stats = {"sj_skipped": 0, "reten": 0, "treasury": 0, "summary": 0}

    for r in gl_rows:
        _, _, dt, ref, jrnl, tdesc, debit, credit, _ = r
        if tdesc in SUMMARY_ROWS:
            stats["summary"] += 1; continue
        if debit is None:
            stats["sj_skipped"] += 1; continue          # crédito = ITBMS de ventas
        if is_reten(ref, tdesc):
            stats["reten"] += 1; continue
        if is_treasury(tdesc):
            stats["treasury"] += 1; continue

        itbms = round(float(debit), 2)
        ln = {
            "tipo": "", "ruc": "", "dv": "", "nombre": "", "factura": "",
            "fecha": date_to_text(dt), "concepto": 1, "compras_bs": 1,
            "monto": round(itbms / rate, 2), "itbms": itbms,
            "source": str(tdesc), "blocked": False, "reasons": [],
        }

        if is_reimbursement(ref):
            nombre, ruc, dv, fac = parse_reimbursement(str(tdesc))
            ln.update(nombre=nombre, ruc=ruc, dv=dv, factura=fac)
            if not ruc:
                ln["blocked"] = True
                ln["reasons"].append("Reembolso sin RUC en el detalle — completar manualmente")
            else:
                mm = lookup_master(nombre, master)
                if mm and not dv:
                    ln["dv"] = mm["dv"]
        else:
            nombre = ITBMS_SUFFIX_RE.sub('', str(tdesc)).strip()
            ln["nombre"] = nombre
            ln["factura"] = '' if ref is None else str(ref).strip()
            m = lookup_master(nombre, master)
            if m:
                ln.update(ruc=m["ruc"], dv=m["dv"], nombre=m["nombre"])
            else:
                ln["blocked"] = True
                ln["reasons"].append(f"Proveedor sin RUC/DV en el maestro: '{nombre}'")

        # sanity: el RUC del proveedor no puede ser el del propio declarante
        if declarant_ruc and ln["ruc"] == declarant_ruc:
            ln["blocked"] = True
            ln["reasons"].append(f"RUC = RUC del declarante ({declarant_ruc}); error de captura, revisar")

        ln["tipo"] = infer_tipo(ln["ruc"]) if ln["ruc"] else ""
        ln["concepto"] = overrides.get(norm_name(ln["nombre"]),
                                       (lookup_master(ln["nombre"], master) or {}).get("concepto", 1))
        lines.append(ln)

    return lines, stats


# ───────────────────────────────────────────────────────── output writer
def write_informe43(lines, out_path: Path):
    """Escribe el .xlsx en formato Informe 43. Columnas A-H como TEXTO (@),
    FECHA como texto 'YYYYMMDD', MONTO como fórmula =J*14.2857142857143.
    Filas bloqueadas resaltadas en amarillo + hoja 'Revisar'."""
    out = Workbook()
    ws = out.active
    ws.title = "Hoja1"
    ws["A1"] = "INFORME 43 - FORMATO A DILIGENCIAR"
    ws["A2"] = "Esta sección de encabezado no se debe modificar. "
    ws["A3"] = "Los datos de este informe deben ser registrados a partir de la línea 5 en adelante."
    for c, h in enumerate(INFORME43_HEADERS, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(bold=True)

    row = 5
    for ln in lines:
        fill = BLOCKED_FILL if ln["blocked"] else None
        text_vals = [ln["tipo"], ln["ruc"], ln["dv"], ln["nombre"], str(ln["factura"]),
                     ln["fecha"], ln["concepto"], ln["compras_bs"]]
        for c, v in enumerate(text_vals, 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.number_format = '@'                      # TEXTO — clave para la fecha
            if fill:
                cell.fill = fill
        m = ws.cell(row=row, column=9, value=f"=+J{row}*14.2857142857143")
        m.number_format = "0.00"
        j = ws.cell(row=row, column=10, value=ln["itbms"])
        j.number_format = "0.00"
        if fill:
            m.fill = fill; j.fill = fill
        row += 1

    blocked = [ln for ln in lines if ln["blocked"]]
    qa = out.create_sheet("Revisar")
    qa.append(["FECHA", "NOMBRE", "FACTURA", "ITBMS", "MOTIVO"])
    for c in range(1, 6):
        qa.cell(row=1, column=c).font = Font(bold=True)
    for ln in blocked:
        qa.append([ln["fecha"], ln["nombre"], str(ln["factura"]), ln["itbms"], "; ".join(ln["reasons"])])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(out_path)
    return out_path


# ───────────────────────────────────────────────────────── public entry
def run_itbms_conversion(xlsx_path, config_path="config/itbms_vendors.json",
                         out_dir="./out", period=None):
    """Convierte el General Ledger crudo de Peachtree al Informe 43.
    Devuelve un dict de resultado estructurado (mismo estilo que convert.py)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    master, overrides, declarant_ruc, rate = load_vendors(config_path)

    wb = load_workbook(xlsx_path, data_only=True)
    ws = _gl_sheet(wb)
    gl_rows = list(ws.iter_rows(values_only=True))[1:]   # quita encabezado
    wb.close()

    lines, stats = transform_rows(gl_rows, master, overrides, declarant_ruc, rate)

    # período: del primer FECHA si no se especifica
    if not period:
        fechas = [ln["fecha"] for ln in lines if ln["fecha"]]
        period = fechas[0][:6] if fechas else datetime.today().strftime("%Y%m")

    decl = (declarant_ruc or "DECLARANTE")
    out_name = f"Informe43_{decl}_{period}.xlsx"
    out_path = write_informe43(lines, out_dir / out_name)

    resolved = [ln for ln in lines if not ln["blocked"]]
    blocked = [ln for ln in lines if ln["blocked"]]

    summary_rows = [{
        "tipo": ln["tipo"], "ruc": ln["ruc"], "dv": ln["dv"], "nombre": ln["nombre"],
        "factura": str(ln["factura"]), "fecha": ln["fecha"], "concepto": ln["concepto"],
        "monto": ln["monto"], "itbms": ln["itbms"],
        "estado": "BLOQUEADA" if ln["blocked"] else "OK",
    } for ln in lines]

    not_processed = [{
        "fecha": ln["fecha"], "nombre": ln["nombre"], "factura": str(ln["factura"]),
        "itbms": ln["itbms"], "motivo": "; ".join(ln["reasons"]),
    } for ln in blocked]

    warnings = []
    if stats["sj_skipped"]:
        warnings.append(f"Se omitieron {stats['sj_skipped']} filas de VENTAS (Jrnl SJ / crédito) — "
                        "esas son ITBMS cobrado, no compras.")
    if stats["reten"]:
        warnings.append(f"Se omitieron {stats['reten']} retenciones (RETEN/RETENCION).")
    if stats["treasury"]:
        warnings.append(f"Se omitió {stats['treasury']} pago(s) al fisco (Tesoro Nacional).")
    if blocked:
        warnings.append(f"{len(blocked)} filas quedaron BLOQUEADAS — revisa la pestaña 'Revisión manual' "
                        "y completa el proveedor en config/itbms_vendors.json.")

    return {
        "rows_written": len(resolved),
        "rows_blocked": len(blocked),
        "summary_rows": summary_rows,
        "not_processed": not_processed,
        "review_flags": [],
        "warnings": warnings,
        "totals": {
            "itbms": round(sum(ln["itbms"] for ln in lines), 2),
            "monto": round(sum(ln["monto"] for ln in lines), 2),
        },
        "stats": stats,
        "period": period,
        "out_path": str(out_path),
        "out_dir": str(out_dir),
    }
