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
    """
    Retorna: (lista de piezas, porcentaje de descuento)
    Si no hay descuento, porcentaje = 0.0
    """
    lineas = texto.splitlines()

    # Encontrar inicio de sección PIEZAS SUSTITUIDAS
    inicio = None
    for i, linea in enumerate(lineas):
        if "PIEZAS SUSTITUIDAS" in linea.upper():
            inicio = i + 1
            break
    if inicio is None:
        return [], 0.0

    # Encontrar fin de sección (Ahorro o SubTotal)
    fin = len(lineas)
    for i in range(inicio, len(lineas)):
        if re.search(r"^ahorro\b|^sub\s*total\b", lineas[i], re.IGNORECASE):
            fin = i
            break

    # Buscar porcentaje de descuento en las líneas después del fin
    descuento_pct = 0.0
    for linea in lineas[fin:fin+15]:
        m_pct = re.search(r'\(?\s*%\s*\)?\s*(\d+(?:\.\d+)?)\s*%', linea)
        if m_pct:
            descuento_pct = float(m_pct.group(1))
            break
        # También buscar formato "40 %" o "(%) 40 %"
        m_pct2 = re.search(r'(\d+(?:\.\d+)?)\s*%', linea)
        if m_pct2:
            val = float(m_pct2.group(1))
            if 0 < val < 100:  # es un porcentaje razonable
                descuento_pct = val
                break

    # Parsear piezas
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

    return piezas, descuento_pct

def generar_excel(resultados):
    wb = Workbook()
    ws = wb.active
    ws.title = "Piezas Sustituidas"

    azul       = "1F4E79"
    azul_medio = "2E75B6"
    claro      = "D9E1F2"
    fill_hdr   = PatternFill("solid", start_color=azul,       end_color=azul)
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

    # Encabezados: No.Orden | Precio Lista | Descuento % | Descuento $ | Precio Final
    cols = [
        ("A", "No. Orden"),
        ("B", "Precio Lista"),
        ("C", "Descuento %"),
        ("D", "Descuento $"),
        ("E", "Precio Final"),
        ("F", "Descripción"),
    ]
    for col, txt in cols:
        c = ws[f"{col}1"]
        c.value     = txt
        c.fill      = fill_hdr
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = borde
    ws.row_dimensions[1].height = 20

    fila = 2
    subtotal_ranges_final = []

    for resultado in resultados:
        numero_orden  = resultado["numero_orden"]
        piezas        = resultado["piezas"]
        descuento_pct = resultado["descuento_pct"]

        if not piezas:
            continue

        fila_inicio = fila

        for i, p in enumerate(piezas):
            fill = fill_alt if i % 2 == 0 else fill_blco

            precio       = p["precio"]
            desc_monto   = round(precio * descuento_pct / 100, 2)
            precio_final = round(precio - desc_monto, 2)

            # A: No. Orden
            c = ws.cell(row=fila, column=1, value=numero_orden)
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="center", vertical="center")

            # B: Precio Lista
            c = ws.cell(row=fila, column=2, value=precio)
            c.number_format = '"$"#,##0.00'
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="right", vertical="center")

            # C: Descuento %
            c = ws.cell(row=fila, column=3, value=descuento_pct / 100 if descuento_pct else 0)
            c.number_format = '0%'
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="center", vertical="center")

            # D: Descuento $
            c = ws.cell(row=fila, column=4, value=desc_monto)
            c.number_format = '"$"#,##0.00'
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="right", vertical="center")

            # E: Precio Final
            c = ws.cell(row=fila, column=5, value=precio_final)
            c.number_format = '"$"#,##0.00'
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="right", vertical="center")

            # F: Descripción
            c = ws.cell(row=fila, column=6, value=p["descripcion"])
            c.fill = fill; c.font = Font(name="Arial", size=10)
            c.border = borde; c.alignment = Alignment(horizontal="left", vertical="center")

            fila += 1

        fila_fin = fila - 1
        subtotal_ranges_final.append(f"E{fila_inicio}:E{fila_fin}")

        # Fila subtotal por orden
        for col in range(1, 7):
            c = ws.cell(row=fila, column=col)
            c.fill = fill_sub; c.border = borde_top

        ws.cell(row=fila, column=1, value=f"Subtotal Orden {numero_orden}").font = Font(name="Arial", bold=True, size=10, color=azul)
        ws.cell(row=fila, column=1).alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=fila, column=2, value=f"=SUM(B{fila_inicio}:B{fila_fin})")
        ws.cell(row=fila, column=2).number_format = '"$"#,##0.00'
        ws.cell(row=fila, column=2).font = Font(name="Arial", bold=True, size=10, color=azul)
        ws.cell(row=fila, column=2).alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=fila, column=4, value=f"=SUM(D{fila_inicio}:D{fila_fin})")
        ws.cell(row=fila, column=4).number_format = '"$"#,##0.00'
        ws.cell(row=fila, column=4).font = Font(name="Arial", bold=True, size=10, color=azul)
        ws.cell(row=fila, column=4).alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=fila, column=5, value=f"=SUM(E{fila_inicio}:E{fila_fin})")
        ws.cell(row=fila, column=5).number_format = '"$"#,##0.00'
        ws.cell(row=fila, column=5).font = Font(name="Arial", bold=True, size=10, color=azul)
        ws.cell(row=fila, column=5).alignment = Alignment(horizontal="right", vertical="center")

        ws.row_dimensions[fila].height = 18
        fila += 2  # espacio entre órdenes

    # Gran Total
    suma = "+".join([f"SUM({r})" for r in subtotal_ranges_final])
    for col in range(1, 7):
        c = ws.cell(row=fila, column=col)
        c.fill = fill_tot; c.border = borde

    ws.cell(row=fila, column=1, value="GRAN TOTAL").font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    ws.cell(row=fila, column=1).fill = fill_tot
    ws.cell(row=fila, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=fila, column=1).border = borde

    ws.cell(row=fila, column=5, value=f"={suma}")
    ws.cell(row=fila, column=5).number_format = '"$"#,##0.00'
    ws.cell(row=fila, column=5).font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    ws.cell(row=fila, column=5).fill = fill_tot
    ws.cell(row=fila, column=5).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=fila, column=5).border = borde
    ws.row_dimensions[fila].height = 22

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 38

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
            piezas, descuento_pct = parsear_piezas(texto)

            if piezas:
                resultados.append({
                    "numero_orden":  numero_orden,
                    "piezas":        piezas,
                    "descuento_pct": descuento_pct,
                })
            else:
                errores.append(f"⚠️ {pdf_file.name} — no se encontraron piezas")

    if errores:
        for e in errores:
            st.warning(e)

    if resultados:
        total_final = sum(
            p["precio"] * (1 - r["descuento_pct"] / 100)
            for r in resultados for p in r["piezas"]
        )
        total_piezas = sum(len(r["piezas"]) for r in resultados)

        st.success(f"✅ {len(resultados)} orden(es)  |  {total_piezas} piezas  |  Total final: ${total_final:,.2f}")

        for r in resultados:
            subtotal_final = sum(p["precio"] * (1 - r["descuento_pct"] / 100) for p in r["piezas"])
            desc_label = f"  |  Descuento: {r['descuento_pct']:.0f}%" if r["descuento_pct"] > 0 else "  |  Sin descuento"
            with st.expander(f"📋 Orden {r['numero_orden']} — {len(r['piezas'])} piezas{desc_label}  |  ${subtotal_final:,.2f}"):
                st.dataframe(
                    {
                        "No. Orden":    [r["numero_orden"]] * len(r["piezas"]),
                        "Precio Lista": [f"${p['precio']:,.2f}" for p in r["piezas"]],
                        "Descuento %":  [f"{r['descuento_pct']:.0f}%" for _ in r["piezas"]],
                        "Descuento $":  [f"${p['precio'] * r['descuento_pct'] / 100:,.2f}" for p in r["piezas"]],
                        "Precio Final": [f"${p['precio'] * (1 - r['descuento_pct']/100):,.2f}" for p in r["piezas"]],
                        "Descripción":  [p["descripcion"] for p in r["piezas"]],
                    },
                    use_container_width=True,
                    hide_index=True
                )

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

