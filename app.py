import streamlit as st
import pdfplumber
import re
import io
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Extractor de Piezas GNP", page_icon="🔧", layout="centered")
st.title("🔧 Extractor de Piezas Sustituidas")
st.markdown("**Valuaciones GNP / Audatex** — Sube uno o varios PDFs y descarga todo en un solo Excel.")
st.divider()

def extraer_numero_orden(nombre_archivo):
    m = re.match(r'^(\d+)', Path(nombre_archivo).stem)
    return m.group(1) if m else Path(nombre_archivo).stem

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
        if re.search(r"ahorro|sub\s*total|total\s*piezas", lineas[i], re.IGNORECASE):
            fin = i
            break

    piezas = []
    for linea in lineas[inicio:fin]:
        linea = linea.strip()
        if not linea:
            continue
        m_precio = re.match(r'^([\$][\d,]+\.\d{2})[\*A-Za-z]?\s+', linea)
        if not m_precio:
            continue
        try:
            precio_num = float(m_precio.group(1).replace('$', '').replace(',', ''))
        except:
            precio_num = 0.0
        tokens = linea[m_precio.end():].split()
        if len(tokens) < 4:
            continue
        descripcion = ' '.join(tokens[1:-2]).strip()
        if not descripcion:
            continue
        if re.match(r'^(precio|referencia|descripci)', descripcion, re.IGNORECASE):
            continue
        piezas.append({"precio": precio_num, "descripcion": descripcion})

    return piezas

def generar_excel(resultados):
    """
    resultados = lista de dicts: {numero_orden, piezas}
    Genera un Excel con todas las órdenes, separadas por un bloque de total por orden
    y un gran total al final.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Piezas Sustituidas"

    azul       = "1F4E79"
    azul_medio = "2E75B6"
    claro      = "D9E1F2"
    fill_hdr   = PatternFill("solid", start_color=azul,       end_color=azul)
    fill_orden = PatternFill("solid", start_color=azul_medio, end_color=azul_medio)
    fill_alt   = PatternFill("solid", start_color=claro,      end_color=claro)
    fill_blco  = PatternFill("solid", start_color="FFFFFF",   end_color="FFFFFF")
    fill_sub   = PatternFill("solid", start_color="BDD7EE",   end_color="BDD7EE")
    fill_tot   = PatternFill("solid", start_color=azul,       end_color=azul)

    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )
    borde_top = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="medium"), bottom=Side(style="thin")
    )

    # ── Encabezados globales ──────────────────────────────────
    headers = [("A","No. Orden"), ("B","Precio"), ("C","Descripción")]
    for col, txt in headers:
        c = ws[f"{col}1"]
        c.value     = txt
        c.fill      = fill_hdr
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = borde
    ws.row_dimensions[1].height = 20

    fila = 2
    subtotal_ranges = []  # para el gran total al final

    for resultado in resultados:
        numero_orden = resultado["numero_orden"]
        piezas       = resultado["piezas"]

        if not piezas:
            continue

        fila_inicio_orden = fila

        for i, p in enumerate(piezas):
            fill = fill_alt if i % 2 == 0 else fill_blco

            co = ws.cell(row=fila, column=1, value=numero_orden)
            co.fill = fill; co.font = Font(name="Arial", size=10)
            co.border = borde
            co.alignment = Alignment(horizontal="center", vertical="center")

            cp = ws.cell(row=fila, column=2, value=p["precio"])
            cp.number_format = '"$"#,##0.00'
            cp.fill = fill; cp.font = Font(name="Arial", size=10)
            cp.border = borde
            cp.alignment = Alignment(horizontal="right", vertical="center")

            cd = ws.cell(row=fila, column=3, value=p["descripcion"])
            cd.fill = fill; cd.font = Font(name="Arial", size=10)
            cd.border = borde
            cd.alignment = Alignment(horizontal="left", vertical="center")

            fila += 1

        # Subtotal por orden
        fila_fin_orden = fila - 1
        subtotal_ranges.append(f"B{fila_inicio_orden}:B{fila_fin_orden}")

        cs_label = ws.cell(row=fila, column=1, value=f"Subtotal Orden {numero_orden}")
        cs_label.font      = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        cs_label.fill      = fill_sub
        cs_label.border    = borde_top
        cs_label.alignment = Alignment(horizontal="center", vertical="center")
        cs_label.font      = Font(name="Arial", bold=True, size=10, color=azul)
        cs_label.fill      = fill_sub

        cs_sum = ws.cell(row=fila, column=2,
                         value=f"=SUM(B{fila_inicio_orden}:B{fila_fin_orden})")
        cs_sum.number_format = '"$"#,##0.00'
        cs_sum.font      = Font(name="Arial", bold=True, size=10, color=azul)
        cs_sum.fill      = fill_sub
        cs_sum.border    = borde_top
        cs_sum.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=fila, column=3).fill   = fill_sub
        ws.cell(row=fila, column=3).border = borde_top

        ws.row_dimensions[fila].height = 18
        fila += 1  # espacio entre órdenes
        fila += 1

    # ── Gran Total ───────────────────────────────────────────
    suma_formula = "+".join([f"SUM({r})" for r in subtotal_ranges])
    gt_label = ws.cell(row=fila, column=1, value="GRAN TOTAL")
    gt_label.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    gt_label.fill      = fill_tot
    gt_label.border    = borde
    gt_label.alignment = Alignment(horizontal="center", vertical="center")

    gt_sum = ws.cell(row=fila, column=2, value=f"={suma_formula}")
    gt_sum.number_format = '"$"#,##0.00'
    gt_sum.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    gt_sum.fill      = fill_tot
    gt_sum.border    = borde
    gt_sum.alignment = Alignment(horizontal="right", vertical="center")

    ws.cell(row=fila, column=3).fill   = fill_tot
    ws.cell(row=fila, column=3).border = borde
    ws.row_dimensions[fila].height = 22

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 40

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Interfaz ─────────────────────────────────────────────────
pdf_files = st.file_uploader(
    "📄 Sube tus valuaciones en PDF",
    type=["pdf"],
    accept_multiple_files=True,
    help="Puedes seleccionar varios PDFs a la vez"
)

if pdf_files:
    resultados = []
    errores    = []

    with st.spinner(f"Procesando {len(pdf_files)} archivo(s)..."):
        for pdf_file in sorted(pdf_files, key=lambda f: f.name):
            numero_orden = extraer_numero_orden(pdf_file.name)
            contenido    = pdf_file.read()
            texto        = extraer_texto(contenido)
            piezas       = parsear_piezas(texto)

            if piezas:
                resultados.append({"numero_orden": numero_orden, "piezas": piezas})
            else:
                errores.append(f"⚠️ {pdf_file.name} — no se encontraron piezas")

    if errores:
        for e in errores:
            st.warning(e)

    if resultados:
        total_piezas = sum(len(r["piezas"]) for r in resultados)
        total_monto  = sum(p["precio"] for r in resultados for p in r["piezas"])

        st.success(f"✅ {len(resultados)} orden(es) procesadas  |  {total_piezas} piezas  |  Total: ${total_monto:,.2f}")

        # Mostrar tabla por orden
        for r in resultados:
            subtotal = sum(p["precio"] for p in r["piezas"])
            with st.expander(f"📋 Orden {r['numero_orden']} — {len(r['piezas'])} piezas  |  ${subtotal:,.2f}"):
                st.dataframe(
                    {
                        "No. Orden":   [r["numero_orden"]] * len(r["piezas"]),
                        "Precio":      [f"${p['precio']:,.2f}" for p in r["piezas"]],
                        "Descripción": [p["descripcion"] for p in r["piezas"]],
                    },
                    use_container_width=True,
                    hide_index=True
                )

        # Descargar Excel con todo
        excel_buf    = generar_excel(resultados)
        ordenes      = "_".join(r["numero_orden"] for r in resultados)
        nombre_excel = f"piezas_{ordenes}.xlsx" if len(resultados) <= 5 else f"piezas_{len(resultados)}_ordenes.xlsx"

        st.download_button(
            label="📥 Descargar Excel con todas las órdenes",
            data=excel_buf,
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.divider()
st.caption("Vanguardia Body & Paint — Extractor de valuaciones GNP")

