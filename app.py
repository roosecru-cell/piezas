import streamlit as st
import pdfplumber
import re
import io
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Extractor Refacciones-Audatex", page_icon="🔧", layout="centered")
st.title("🔧 Extractor de Refacciones Audatex")
st.markdown("**Valuaciones / Audatex** — Sube uno o varios PDFs y descarga todo en un solo Excel.")
st.divider()

def extraer_numero_orden(nombre_archivo):
    m = re.match(r'^(\w+)', Path(nombre_archivo).stem)
    return m.group(1) if m else Path(nombre_archivo).stem

def extraer_texto(contenido_bytes):
    texto_completo = []
    with pdfplumber.open(io.BytesIO(contenido_bytes)) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)
    return "\n".join(texto_completo)

def es_token_descripcion(token):
    """Determina si un token pertenece a la descripción o al número de pieza."""
    if len(token) == 1: return False           # letra sola (R, T) = numpieza
    if token == '..': return False             # especificación llanta
    if re.search(r'\d', token): return False   # contiene número = numpieza
    if token.endswith('.'): return True        # termina en punto = descripción
    if '.' in token and re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑ\.]+$', token): return True  # I.DIRECCION
    if re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑ\-]+$', token): return True  # palabra pura
    return False

def extraer_descripcion_linea(linea):
    """Extrae precio y descripción de una línea de pieza."""
    m = re.match(r'^([\$][\d,]+\.\d{2})[\*A-Za-z]?\s+', linea)
    if not m: return None, None
    precio = m.group(1)
    tokens = linea[m.end():].split()
    if len(tokens) < 3: return precio, None

    # tokens[0]=referencia, tokens[-1]=Pos.BD (solo dígitos)
    middle = tokens[1:-1]

    # Buscar el último token que pertenece a la descripción
    desc_end = 0
    for i, t in enumerate(middle):
        if es_token_descripcion(t):
            desc_end = i + 1

    desc = ' '.join(middle[:desc_end]).strip()
    return precio, desc if desc else None

def parsear_piezas(texto):
    lineas = texto.splitlines()

    inicio = None
    for i, linea in enumerate(lineas):
        if "PIEZAS SUSTITUIDAS" in linea.upper():
            inicio = i + 1
            break
    if inicio is None:
        return [], 0.0

    fin = len(lineas)
    for i in range(inicio, len(lineas)):
        if re.search(r"^ahorro\b|^sub\s*total\b", lineas[i], re.IGNORECASE):
            fin = i
            break

    # Buscar porcentaje de descuento después del bloque
    descuento_pct = 0.0
    for linea in lineas[fin:fin+15]:
        m_pct = re.search(r'(\d+(?:\.\d+)?)\s*%', linea)
        if m_pct:
            val = float(m_pct.group(1))
            if 0 < val < 100:
                descuento_pct = val
                break

    piezas = []
    for linea in lineas[inicio:fin]:
        linea = linea.strip()
        if not linea:
            continue
        precio, desc = extraer_descripcion_linea(linea)
        if not precio or not desc:
            continue
        if re.match(r'^(precio|referencia|descripci)', desc, re.IGNORECASE):
            continue
        try:
            precio_num = float(precio.replace('$', '').replace(',', ''))
        except:
            precio_num = 0.0
        piezas.append({"precio": precio_num, "descripcion": desc})

    return piezas, descuento_pct

def generar_excel(resultados):
    wb = Workbook()
    ws = wb.active
    ws.title = "Piezas Sustituidas"

    azul  = "1F4E79"
    claro = "D9E1F2"
    fill_hdr  = PatternFill("solid", start_color=azul,     end_color=azul)
    fill_alt  = PatternFill("solid", start_color=claro,    end_color=claro)
    fill_blco = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    fill_sub  = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
    fill_tot  = PatternFill("solid", start_color=azul,     end_color=azul)
    borde = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"),  bottom=Side(style="thin"))
    borde_top = Border(left=Side(style="thin"), right=Side(style="thin"),
                       top=Side(style="medium"), bottom=Side(style="thin"))

    cols = [("A","No. Orden"),("B","Descripción"),("C","Precio Lista"),
            ("D","Descuento %"),("E","Descuento $"),("F","Precio Final")]
    for col, txt in cols:
        c = ws[f"{col}1"]
        c.value = txt; c.fill = fill_hdr
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[1].height = 20

    fila = 2
    subtotal_ranges = []

    for resultado in resultados:
        numero_orden  = resultado["numero_orden"]
        piezas        = resultado["piezas"]
        descuento_pct = resultado["descuento_pct"]
        if not piezas: continue

        fila_inicio = fila
        for i, p in enumerate(piezas):
            fill         = fill_alt if i % 2 == 0 else fill_blco
            precio       = p["precio"]
            desc_monto   = round(precio * descuento_pct / 100, 2)
            precio_final = round(precio - desc_monto, 2)

            def cell(col, val, fmt=None, align="left"):
                c = ws.cell(row=fila, column=col, value=val)
                c.fill = fill; c.font = Font(name="Arial", size=10)
                c.border = borde
                c.alignment = Alignment(horizontal=align, vertical="center")
                if fmt: c.number_format = fmt
                return c

            cell(1, numero_orden, align="center")
            cell(2, p["descripcion"])
            cell(3, precio, '"$"#,##0.00', "right")
            cell(4, descuento_pct/100 if descuento_pct else 0, '0%', "center")
            cell(5, desc_monto, '"$"#,##0.00', "right")
            cell(6, precio_final, '"$"#,##0.00', "right")
            fila += 1

        fila_fin = fila - 1
        subtotal_ranges.append(f"F{fila_inicio}:F{fila_fin}")

        # Subtotal por orden
        for col in range(1, 7):
            c = ws.cell(row=fila, column=col)
            c.fill = fill_sub; c.border = borde_top

        def sub(col, val=None, fmt=None):
            c = ws.cell(row=fila, column=col, value=val)
            c.fill = fill_sub
            c.font = Font(name="Arial", bold=True, size=10, color=azul)
            c.border = borde_top
            c.alignment = Alignment(horizontal="right" if fmt else "center", vertical="center")
            if fmt: c.number_format = fmt
            return c

        sub(1, f"Subtotal Orden {numero_orden}")
        sub(3, f"=SUM(C{fila_inicio}:C{fila_fin})", '"$"#,##0.00')
        sub(5, f"=SUM(E{fila_inicio}:E{fila_fin})", '"$"#,##0.00')
        sub(6, f"=SUM(F{fila_inicio}:F{fila_fin})", '"$"#,##0.00')
        ws.row_dimensions[fila].height = 18
        fila += 2

    # Gran Total
    suma = "+".join([f"SUM({r})" for r in subtotal_ranges])
    for col in range(1, 7):
        c = ws.cell(row=fila, column=col)
        c.fill = fill_tot; c.border = borde

    ws.cell(row=fila, column=1, value="GRAN TOTAL")
    ws.cell(row=fila, column=1).font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    ws.cell(row=fila, column=1).fill = fill_tot
    ws.cell(row=fila, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=fila, column=1).border = borde

    gt = ws.cell(row=fila, column=6, value=f"={suma}")
    gt.number_format = '"$"#,##0.00'
    gt.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    gt.fill = fill_tot
    gt.alignment = Alignment(horizontal="right", vertical="center")
    gt.border = borde
    ws.row_dimensions[fila].height = 22

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

# ── Interfaz ─────────────────────────────────────────────────
pdf_files = st.file_uploader(
    "📄 Sube tus valuaciones en PDF", type=["pdf"],
    accept_multiple_files=True, help="Puedes seleccionar varios PDFs a la vez"
)

if pdf_files:
    resultados = []; errores = []
    with st.spinner(f"Procesando {len(pdf_files)} archivo(s)..."):
        for pdf_file in sorted(pdf_files, key=lambda f: f.name):
            numero_orden = extraer_numero_orden(pdf_file.name)
            contenido    = pdf_file.read()
            texto        = extraer_texto(contenido)
            piezas, descuento_pct = parsear_piezas(texto)
            if piezas:
                resultados.append({"numero_orden": numero_orden, "piezas": piezas, "descuento_pct": descuento_pct})
            else:
                errores.append(f"⚠️ {pdf_file.name} — no se encontraron piezas")

    for e in errores: st.warning(e)

    if resultados:
        total_final  = sum(p["precio"]*(1-r["descuento_pct"]/100) for r in resultados for p in r["piezas"])
        total_piezas = sum(len(r["piezas"]) for r in resultados)
        st.success(f"✅ {len(resultados)} orden(es)  |  {total_piezas} piezas  |  Total final: ${total_final:,.2f}")

        for r in resultados:
            subtotal   = sum(p["precio"]*(1-r["descuento_pct"]/100) for p in r["piezas"])
            desc_label = f"  |  Descuento: {r['descuento_pct']:.0f}%" if r["descuento_pct"] > 0 else "  |  Sin descuento"
            with st.expander(f"📋 Orden {r['numero_orden']} — {len(r['piezas'])} piezas{desc_label}  |  ${subtotal:,.2f}"):
                st.dataframe({
                    "No. Orden":    [r["numero_orden"]] * len(r["piezas"]),
                    "Descripción":  [p["descripcion"] for p in r["piezas"]],
                    "Precio Lista": [f"${p['precio']:,.2f}" for p in r["piezas"]],
                    "Descuento %":  [f"{r['descuento_pct']:.0f}%" for _ in r["piezas"]],
                    "Descuento $":  [f"${p['precio']*r['descuento_pct']/100:,.2f}" for p in r["piezas"]],
                    "Precio Final": [f"${p['precio']*(1-r['descuento_pct']/100):,.2f}" for p in r["piezas"]],
                }, use_container_width=True, hide_index=True)

        excel_buf    = generar_excel(resultados)
        ordenes      = "_".join(r["numero_orden"] for r in resultados)
        nombre_excel = f"piezas_{ordenes}.xlsx" if len(resultados) <= 5 else f"piezas_{len(resultados)}_ordenes.xlsx"
        st.download_button(label="📥 Descargar Excel con todas las órdenes",
            data=excel_buf, file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

st.divider()
st.caption("Extractor de Refacciones")
