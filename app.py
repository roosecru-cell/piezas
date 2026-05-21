import streamlit as st
import pdfplumber
import re
import io
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Configuración de la página ───────────────────────────────
st.set_page_config(
    page_title="Extractor de Piezas GNP",
    page_icon="🔧",
    layout="centered"
)

st.title("🔧 Extractor de Piezas Sustituidas")
st.markdown("**Valuaciones GNP / Audatex** — Sube tu PDF y descarga el Excel al instante.")
st.divider()

# ── Funciones ────────────────────────────────────────────────

def extraer_texto(contenido_bytes):
    texto_completo = []
    with pdfplumber.open(io.BytesIO(contenido_bytes)) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)
    return "\n".join(texto_completo)


def parsear_piezas(texto):
    lineas = texto.splitlines()

    inicio = None
    for i, linea in enumerate(lineas):
        if "PIEZAS SUSTITUIDAS" in linea.upper():
            inicio = i + 1
            break

    if inicio is None:
        return []

    fin = len(lineas)
    for i in range(inicio, len(lineas)):
        if re.search(r"^ahorro\b|^sub\s*total\b", lineas[i], re.IGNORECASE):
            fin = i
            break

    bloque = lineas[inicio:fin]

    patron = re.compile(
        r"^(\$[\d,]+\.\d{2})[A-Za-z]?\s+"
        r"\S+\s+"
        r"(.+?)\s+"
        r"[A-Z0-9]{8,}\s+"
        r"\d+\s*$"
    )

    piezas = []
    for linea in bloque:
        linea = linea.strip()
        if not linea:
            continue
        m = patron.match(linea)
        if m:
            precio = m.group(1)
            desc   = m.group(2).strip()
            if desc and not re.match(r"^(precio|referencia|descripci)", desc, re.IGNORECASE):
                piezas.append({"precio": precio, "descripcion": desc})

    return piezas


def generar_excel(piezas, nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "Piezas Sustituidas"

    azul  = "1F4E79"
    claro = "D9E1F2"

    fill_hdr  = PatternFill("solid", start_color=azul,      end_color=azul)
    fill_alt  = PatternFill("solid", start_color=claro,     end_color=claro)
    fill_blco = PatternFill("solid", start_color="FFFFFF",  end_color="FFFFFF")
    fill_tot  = PatternFill("solid", start_color=azul,      end_color=azul)
    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    ws.merge_cells("A1:B1")
    ws["A1"] = f"PIEZAS SUSTITUIDAS  |  {nombre}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color=azul)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for col, txt in [("A", "Precio"), ("B", "Descripción")]:
        c = ws[f"{col}2"]
        c.value     = txt
        c.fill      = fill_hdr
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = borde
    ws.row_dimensions[2].height = 20

    for i, p in enumerate(piezas, start=3):
        fill = fill_alt if i % 2 == 0 else fill_blco
        cp = ws.cell(row=i, column=1, value=p["precio"])
        cd = ws.cell(row=i, column=2, value=p["descripcion"])
        for c in [cp, cd]:
            c.fill   = fill
            c.font   = Font(name="Arial", size=10)
            c.border = borde
        cp.alignment = Alignment(horizontal="right", vertical="center")
        cd.alignment = Alignment(horizontal="left",  vertical="center")

    ft = len(piezas) + 3
    ws.merge_cells(f"A{ft}:B{ft}")
    ws[f"A{ft}"] = f"Total piezas: {len(piezas)}"
    ws[f"A{ft}"].font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    ws[f"A{ft}"].fill      = fill_tot
    ws[f"A{ft}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"A{ft}"].border    = borde

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Interfaz ─────────────────────────────────────────────────

pdf_file = st.file_uploader(
    "📄 Sube tu valuación en PDF",
    type=["pdf"],
    help="Archivos PDF de valuaciones GNP / Audatex"
)

if pdf_file is not None:
    with st.spinner("Procesando PDF..."):
        contenido = pdf_file.read()
        texto     = extraer_texto(contenido)
        piezas    = parsear_piezas(texto)

    if not piezas:
        st.error("❌ No se encontró la sección 'PIEZAS SUSTITUIDAS' en este PDF. Verifica que sea una valuación GNP/Audatex.")
    else:
        st.success(f"✅ {len(piezas)} piezas encontradas")

        # Mostrar tabla
        st.subheader("Piezas sustituidas")
        st.dataframe(
            {"Precio": [p["precio"] for p in piezas],
             "Descripción": [p["descripcion"] for p in piezas]},
            use_container_width=True,
            hide_index=True
        )

        # Botón de descarga
        nombre_base  = Path(pdf_file.name).stem
        excel_buf    = generar_excel(piezas, nombre_base)
        nombre_excel = nombre_base + "_piezas.xlsx"

        st.download_button(
            label="📥 Descargar Excel",
            data=excel_buf,
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.divider()
st.caption("Vanguardia Body & Paint — Extractor de valuaciones GNP")

