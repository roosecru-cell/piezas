import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Extractor Órdenes GNP", page_icon="🚗", layout="wide")

st.markdown("""
<style>
    .stButton>button { background-color: #C8102E; color: white; border-radius: 8px; font-weight: bold; }
    .result-box { background-color: #e8f5e9; border-left: 4px solid #2e7d32; padding: 12px; border-radius: 4px; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

st.title("🚗 Extractor de Órdenes de Admisión GNP")
st.caption("Sube uno o varios PDFs de órdenes GNP y descarga un Excel con los datos clave.")

# ─── Helpers ─────────────────────────────────────────────────────────────────

def clean(t):
    return " ".join(t.strip().split()) if t else ""

def next_line_after(lines, header_keywords):
    """Devuelve la línea inmediata después de la que contiene alguno de los keywords."""
    for i, line in enumerate(lines):
        if any(kw.lower() in line.lower() for kw in header_keywords):
            if i + 1 < len(lines):
                return clean(lines[i + 1])
    return None

def value_in_same_line(lines, keyword):
    """Busca keyword y devuelve lo que sigue en la misma línea después de él."""
    for line in lines:
        if keyword.lower() in line.lower():
            after = re.split(re.escape(keyword), line, flags=re.IGNORECASE, maxsplit=1)
            if len(after) > 1 and after[1].strip():
                return clean(after[1])
    return None

def extract_field_after_header(lines, header_keywords, stop_keywords=None):
    """Para campos que están en la línea siguiente al header de columnas."""
    stop_keywords = stop_keywords or []
    for i, line in enumerate(lines):
        if any(kw.lower() in line.lower() for kw in header_keywords):
            # Buscar siguiente línea no vacía
            for j in range(i + 1, min(i + 4, len(lines))):
                candidate = clean(lines[j])
                if not candidate:
                    continue
                # Si la siguiente línea es otro header, no es dato
                if any(sk.lower() in candidate.lower() for sk in stop_keywords):
                    break
                return candidate
    return None

def extract_block_after_label(lines, label_keywords, stop_keywords):
    """Extrae un bloque de texto (puede ser multilinea) después de un label."""
    collecting = False
    result = []
    for line in lines:
        stripped = clean(line)
        if not collecting:
            if any(kw.lower() in stripped.lower() for kw in label_keywords):
                collecting = True
                continue
        else:
            if any(sk.lower() in stripped.lower() for sk in stop_keywords):
                break
            if stripped:
                result.append(stripped)
    return " ".join(result) if result else None

def extract_cdr(lines):
    """Extrae el CDR de la sección Observaciones."""
    obs_lines = []
    in_obs = False
    for line in lines:
        stripped = clean(line)
        if not in_obs:
            if stripped.lower() == "observaciones":
                in_obs = True
                continue
        else:
            if any(k in stripped.lower() for k in ["piezas faltantes", "nombre y firma"]):
                break
            if stripped:
                obs_lines.append(stripped)
    obs_text = " ".join(obs_lines)
    m = re.search(r"CDR\s*:\s*(.+?)(?:Direcci[oó]n|$)", obs_text, re.IGNORECASE)
    if m:
        return clean(m.group(1))
    # Si no hay CDR: en Observaciones, buscar en todo el texto
    full = " ".join(lines)
    m2 = re.search(r"CDR\s*:\s*(.+?)(?:Direcci[oó]n|Siniestro|$)", full, re.IGNORECASE)
    return clean(m2.group(1)) if m2 else "N/A"

# ─── Extractor principal ──────────────────────────────────────────────────────

def extract_data(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = full_text.splitlines()
    data = {}

    # Detectar formato B (Información general / Autoajuste)
    formato_b = any("Información general" in l for l in lines)

    # ── No. de siniestro ──────────────────────────────────────────────────
    hdr = ["No. de siniestro Fecha de siniestro", "Número de siniestro Fecha de atención",
           "No. de siniestro", "Número de siniestro"]
    val = extract_field_after_header(lines, hdr,
          stop_keywords=["Estado", "Número de póliza", "Datos de la póliza"])
    if val:
        # El valor puede ser "0183607944 21/05/2026 ..." — tomar primer token
        data["No. Siniestro"] = val.split()[0]
    else:
        data["No. Siniestro"] = "N/A"

    # ── Fecha de atención ─────────────────────────────────────────────────
    # Está en la misma fila de datos que el siniestro (Formato A: columna 4)
    # Buscar línea con patrón fecha/hora de atención
    fecha_aten = "N/A"
    for line in lines:
        # Línea con 4-5 fechas dd/mm/aaaa: siniestro | fecha sin | hora sin | fecha aten | fecha entrega
        dates = re.findall(r'\d{2}/\d{2}/\d{4}', line)
        times = re.findall(r'\d{2}:\d{2}(?:\s*[AP]M)?', line)
        if len(dates) >= 3:
            # Formato A: [fecha_sin, fecha_aten, fecha_entrega] con horas intercaladas
            # Reconstruimos posiciones: buscar "fecha aten" = 3ra fecha
            fecha_aten = dates[2] if len(dates) >= 3 else dates[-1]
            break
        elif len(dates) == 2 and formato_b:
            # Formato B: "0183313899 15/05/2026 09:00 15/05/2026 10:11"
            fecha_aten = dates[0]
            break
    data["Fecha de Atención"] = fecha_aten

    # ── Versión ───────────────────────────────────────────────────────────
    val = extract_field_after_header(lines, ["Versión Placas", "Tipo de vehículo Versión"],
          stop_keywords=["Vehículo responsable", "Modelo Número"])
    if not val:
        for i, line in enumerate(lines):
            if re.match(r"Tipo de veh[ií]culo\s+Versi[oó]n", line, re.IGNORECASE):
                if i + 1 < len(lines):
                    v = clean(lines[i+1])
                    # quitar el prefijo de tipo (AUT, AUTOMOVIL, etc.)
                    v = re.sub(r"^(AUT|AUTOMÓVIL|AUTOMOVIL)\s+", "", v, flags=re.IGNORECASE)
                    val = v
                    break
    if val:
        # Formato A la línea es "VOLKSWAGEN CROSSGOLF L4 1.4... UPG752J 3VWVB6..."
        # La versión termina antes de las placas — tomar hasta un patrón de placa
        m = re.match(r"(.+?)\s+[A-Z]{2,3}\d{3,4}[A-Z]?\b", val)
        data["Versión"] = clean(m.group(1)) if m else val
    else:
        data["Versión"] = "N/A"

    # ── Modelo (año) ──────────────────────────────────────────────────────
    val = extract_field_after_header(lines,
          ["Clasificación Tipo de vehículo Armadora"],
          stop_keywords=["Versión", "Vehículo responsable"])
    if val:
        # línea: "AUTOMÓVIL VOLKSWAGEN VOLKSWAGEN CROSSGOLF 2017"
        m = re.search(r"\b(19|20)\d{2}\b", val)
        data["Modelo"] = m.group(0) if m else val.split()[-1]
    else:
        # Formato B: "Modelo Número de serie" → siguiente línea "2014 WBA3..."
        val2 = extract_field_after_header(lines, ["Modelo Número de serie"],
               stop_keywords=["Daños", "Declaración"])
        if val2:
            data["Modelo"] = val2.split()[0]
        else:
            data["Modelo"] = "N/A"

    # ── Placas ────────────────────────────────────────────────────────────
    # Aparece en la línea de datos de versión: buscar patrón de placa mexicana
    placas = "N/A"
    for line in lines:
        m = re.search(r'\b([A-Z]{2,3}\d{3,4}[A-Z]?)\b', line)
        if m and "Placas" not in line and "Número" not in line:
            placas = m.group(1)
            break
    data["Placas"] = placas

    # ── Nombre conductor / asegurado ──────────────────────────────────────
    val = extract_field_after_header(lines,
          ["Nombre del conductor Fecha de nacimiento"],
          stop_keywords=["Identificación", "Licencia"])
    if not val:
        val = extract_field_after_header(lines,
              ["Nombre del Asegurado Referencia", "Nombre del asegurado Cobertura"],
              stop_keywords=["Datos del vehículo", "Teléfono"])
    if val:
        # En formato A la línea es "LARISSA CORELLY... 20/09/1998 27 5651164865"
        # Tomar solo la parte del nombre (sin fecha/edad/tel)
        m = re.match(r"([A-ZÁÉÍÓÚÜÑa-záéíóúüñ ]{5,}?)(?:\s+\d{2}/\d{2}/\d{4}|\s+\d{10})", val)
        data["Nombre / Conductor"] = clean(m.group(1)) if m else val
    else:
        data["Nombre / Conductor"] = "N/A"

    # ── Teléfono ──────────────────────────────────────────────────────────
    tel = "N/A"
    for line in lines:
        # Buscar línea con nombre+tel: "NOMBRE 20/09/1998 27 5651164865, 7821759935"
        m = re.search(r'\b(\d{10}(?:,\s*\d{10})?)\s*$', line)
        if m:
            tel = m.group(1)
            break
    data["Teléfono"] = tel

    # ── Correo electrónico ────────────────────────────────────────────────
    correo = "N/A"
    email_pattern = re.compile(r'[\w.+-]+@[\w.-]+\.\w+')
    for line in lines:
        emails = email_pattern.findall(line)
        if emails:
            correo = ", ".join(emails)
            break
    data["Correo Electrónico"] = correo

    # ── Daños a consecuencia ──────────────────────────────────────────────
    val = extract_block_after_label(lines,
          ["Daños a consecuencia del siniestro", "Daños a consecuencia"],
          stop_keywords=["Observaciones", "Piezas faltantes", "Nombre y firma", "Declaración"])
    data["Daños a Consecuencia"] = val if val else "N/A"

    # ── CDR (de Observaciones) ────────────────────────────────────────────
    data["CDR"] = extract_cdr(lines)

    return data


# ─── Excel builder ────────────────────────────────────────────────────────────

COLUMNS = [
    "No. Siniestro", "Versión", "Modelo", "Placas",
    "Nombre / Conductor", "Teléfono", "Correo Electrónico",
    "Fecha de Atención", "Daños a Consecuencia", "CDR"
]

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Órdenes GNP"

    thin = Side(style="thin", color="BDBDBD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", start_color="C8102E")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30

    for row_idx, row in enumerate(rows, 2):
        fill_color = "FCE4EC" if row_idx % 2 == 0 else "FFFFFF"
        fill = PatternFill("solid", start_color=fill_color)
        for col_idx, col_name in enumerate(COLUMNS, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(col_name, "N/A"))
            cell.font = Font(name="Arial", size=9)
            cell.fill = fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[row_idx].height = 50

    widths = [16, 40, 8, 12, 30, 24, 32, 14, 55, 45]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── UI ──────────────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "📂 Sube los PDFs de Órdenes de Admisión GNP",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("⚡ Extraer datos"):
        all_rows = []
        errores = []
        progress = st.progress(0)

        for i, f in enumerate(uploaded_files):
            try:
                row = extract_data(f)
                all_rows.append(row)
            except Exception as e:
                errores.append(f"❌ `{f.name}` — {e}")
            progress.progress((i + 1) / len(uploaded_files))

        for msg in errores:
            st.warning(msg)

        if all_rows:
            excel_bytes = build_excel(all_rows)

            st.markdown('<div class="result-box">', unsafe_allow_html=True)
            st.markdown(f"✅ **{len(all_rows)} orden(es)** procesada(s) de **{len(uploaded_files)} PDF(s)**")
            st.markdown('</div>', unsafe_allow_html=True)

            df_show = pd.DataFrame(all_rows)[COLUMNS]
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            out_name = (
                f"orden_{all_rows[0]['No. Siniestro']}.xlsx"
                if len(all_rows) == 1
                else "ordenes_admision_gnp.xlsx"
            )

            st.download_button(
                label="📥  Descargar Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("No se pudo extraer información de ningún PDF.")
else:
    st.info("👆 Sube uno o más PDFs de Órdenes de Admisión GNP para comenzar.")

st.markdown("---")
st.caption("New Roads · Extractor Órdenes de Admisión GNP")
