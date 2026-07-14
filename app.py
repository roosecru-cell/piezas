import streamlit as st
import pdfplumber
import re
import io
import os
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="Extractor AUDATEX Completo",
    page_icon="🔧",
    layout="wide"
)

st.markdown("""
<style>
.main-header{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    padding:2rem;border-radius:12px;margin-bottom:2rem;text-align:center;}
.main-header h1{color:#e94560;margin:0;font-size:2rem;}
.main-header p{color:#a8b2d8;margin:.4rem 0 0;font-size:.95rem;}
.metric-box{background:white;border-radius:8px;padding:.8rem 1rem;
    text-align:center;box-shadow:0 2px 6px rgba(0,0,0,.08);}
.metric-value{font-size:1.6rem;font-weight:800;color:#0f3460;}
.metric-label{font-size:.78rem;color:#6c757d;margin-top:.2rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🔧 Extractor AUDATEX Completo</h1>
  <p>Extrae <b>Refacciones</b> · <b>Mano de Obra Hojal/Mecánica</b> · <b>Pintura de Carrocería</b> en un solo Excel</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CATÁLOGO DE TÉRMINOS MECÁNICA
# ══════════════════════════════════════════════════════════════════════════════
MECANICA_TERMS_RAW = [
    "RIN DL.D.:D+M", "RIN DL.I.:D+M", "RIN DEL.D. REPARAR",
    "BRAZO SUSPENS.DL.I.INF.:D+M", "LLANTA DL.D.:D+M", "LLANTA DL.I.:D+M",
    "LLANTA DL.I.:D+M Y BALANCEAR", "RADIADOR:D+M", "RIN DEL.I. REPARAR",
    "03 17 00 LIQUIDO REFRIG.:VAC-LLENAR", "05 19 00 RIN DL.D.:D+M",
    "38 17 70 LIQUIDO REFRIG.:VACIAR-RELLENAR", "50 00 ZAX FUNCION GFS/AJUSTES",
    "AIRE ACONDIC.:VAC-LLENAR", "ALINEACION", "ALINEACION Y BALANCE",
    "AMORTIG.DL.D.:D+M", "AMORTIG.DL.D.:DESPIEZ-ENSAMB.(DESMONT.)",
    "AMORTIG.DL.I.:DESPIEZ-ENSAMBLAR(DESMONT)",
    "AMORTIG.DL.I./D.:DESP-ENSAMB.(DESMONT.)", "AMORTIG.DL.I.CPL.:D+M",
    "AMORTIG.FACIA DL.:D+M", "ANTICONGELANTE", "ARNES",
    "ASIST.CAMBIO CARRIL:D+M(TRAB.ADIC.)", "BALANCEO",
    "BIELETA I.ESTABI.:D+M", "BRAZO SUSPENS.DL.INF.:D+M (LLANTA DESM.)",
    "CALIPER FRENO DL.I.:SOLT-FIJ.", "CAMARA 360 GRADOS: AJUSTAR",
    "CARGA DE GAS", "CARGA GAS", "DEPOS.EXPANSION:D+M",
    "DES-/MONTAR FLUIDO AIRE ACONDIC.", "DES+MON RIN TRA.D.",
    "ESCANEO AIRBAG", "FILTRO AIRE:D+M", "FLECHA CARDAN DIREC.I.:D+M",
    "FLECHA MOTRIZ DL.I.CPL.:D+M", "GALON ANTICONGELANTE",
    "KA) AMORTIG./MANGUETA DL.D.:SOLT-FIJ", "KA) AMORTIG.DL.D.:D+M",
    "KA) AMORTIG.DL.D./CARROCER.:SOLT.-FIJ",
    "KA) AMORTIGUADOR/S DL.:D+M TRAB.ADIC.", "KA) LLANTA/LLANTAS:D+M TRAB.ADIC.",
    "KA) RIN DL.I.:D+M", "LIQUIDO REFRIGERANTE DES+MON/SUSTITUIR",
    "LIQUIDO REFRIGERANTE REPARAR", "LLANTA DL.D.:D+M(LLANTA DESMONT.)",
    "LLANTA DL.D.:MONTAR/BALANCEAR", "LLANTA Y/O RIN(ES):D+M",
    "PROG SENSOR RADAR", "R RIN DEL DER", "R RIN DEL IZQ",
    "RADIADOR EGR:D+M", "REC.RODAM.DL.D.:D+M", "REFRIGERANTE A.ACOND REPARAR",
    "REP ARNES SENSO REVE", "RESONADOR INF. REPARAR", "RIN TRA.D. REPARAR",
    "ROTULA BIELETA DIREC.I.:D+M", "ROTULA SOP.DL.I.:SOLT-FIJ",
    "SENSOR BOLSA AIRE DL.:D+M", "SENSOR TR.CN.ASIST.ESTAC.:D+M",
    "SOP.D.RADIADOR:SUST.", "VA) BRAZO SUSPENS.DL.D.:D+M",
    "VA) BRAZO SUSPENS.DL.I./D.:D+M TRAB.ADIC.", "VA) LLANTA DL.D.:BALANCEAR",
    "VA) PUNTA BIELETA DIREC.:D+M TRAB.ADIC.", "VA) ROTULA BIELETA DIREC.D.:D+M",
    "VALV.PRESION DEL.I. REPARAR",
]

def _norm(s: str) -> str:
    s = s.strip().upper()
    s = re.sub(r'\d{8,}\s*$', '', s).strip()
    return re.sub(r'\s+', ' ', s)

MECANICA_SET = {_norm(t) for t in MECANICA_TERMS_RAW}

def clasificar_mo(desc: str) -> str:
    return "Mecánica" if _norm(desc) in MECANICA_SET else "Hojalatería"

# ══════════════════════════════════════════════════════════════════════════════
# EXTRACCIÓN DE REFACCIONES
# ══════════════════════════════════════════════════════════════════════════════

def es_token_descripcion(token):
    if len(token) == 1: return False
    if token == '..': return False
    if re.search(r'\d', token): return False
    if token.endswith('.'): return True
    if '.' in token and re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑ\.]+$', token): return True
    if re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑ\-]+$', token): return True
    return False

def extraer_descripcion_linea(linea):
    m = re.match(r'^([\$][\d,]+\.\d{2})[\*A-Za-z]?\s+', linea)
    if not m: return None, None
    precio = m.group(1)
    tokens = linea[m.end():].split()
    if len(tokens) < 3: return precio, None
    middle = tokens[1:-1]
    desc_end = 0
    for i, t in enumerate(middle):
        if es_token_descripcion(t):
            desc_end = i + 1
    desc = ' '.join(middle[:desc_end]).strip()
    return precio, desc if desc else None

def extraer_refacciones(texto: str):
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
        if not linea: continue
        precio, desc = extraer_descripcion_linea(linea)
        if not precio or not desc: continue
        if re.match(r'^(precio|referencia|descripci)', desc, re.IGNORECASE): continue
        try:
            precio_num = float(precio.replace('$','').replace(',',''))
        except:
            precio_num = 0.0
        piezas.append({"precio": precio_num, "descripcion": desc})

    return piezas, descuento_pct

# ══════════════════════════════════════════════════════════════════════════════
# EXTRACCIÓN MO + PINTURA
# ══════════════════════════════════════════════════════════════════════════════

def parse_price(s: str) -> float:
    return float(s.replace(",", ""))

RE_ITEM    = re.compile(r'^(\S+)\s+(.+?)\s+[\d.]+\s*(?:[R*]+)?\s*\$([\d,]+\.?\d*)\*?$')
RE_TIEMPO  = re.compile(r'^(TIEMPO\s+DE\s+PREP\..+?)\s+[\d.]+\s*R\s+\$([\d,]+\.?\d*)', re.I)
RE_MO_A    = re.compile(r'^Mano\s+de\s+Obra', re.I)
RE_MO_B    = re.compile(r'^Hojal/Mec', re.I)
RE_MO_TBL  = re.compile(r'NR\s+Operaci.+Trabajo\s+UT\s+Precio', re.I)
RE_MO_STOP = re.compile(r'^Total\s+(Unidades|M\.O\.)', re.I)
RE_PIN_HDR = re.compile(r'^PINTURA\s+DE\s+CARROCER', re.I)
RE_PIN_TBL = re.compile(r'NR\s+Operaci.+UT\s+Precio', re.I)
RE_PIN_STP = re.compile(r'^(RESUMEN\s+M\.O|Total\s+de\s+Horas)', re.I)
SKIP       = {'TOTAL','SUBTOTAL','SUMA','IVA','RESUMEN','PIEZAS'}

def extraer_mo_pintura(pdf_bytes: bytes) -> dict:
    mo_items, pin_items, meta = [], [], {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        lines = []
        for page in pdf.pages:
            lines.extend((page.extract_text() or "").splitlines())

    for line in lines:
        s = line.strip()
        def g(pat): m=re.search(pat,line); return m.group(1).strip() if m else None
        if "Número de Expediente:" in line and "expediente" not in meta:
            meta["expediente"] = g(r'Número de Expediente:\s*(\S+)') or "–"
        if "Taller de Reparación:" in line and "taller" not in meta:
            meta["taller"] = g(r'Taller de Reparación:\s*(.+)') or "–"
        if "Fabricante:" in line and "fabricante" not in meta:
            meta["fabricante"] = g(r'Fabricante:\s*(.+)') or ""
        if re.search(r'^Modelo:\s*\S', line) and "modelo" not in meta:
            meta["modelo"] = g(r'Modelo:\s*(.+)') or ""
        if "No. VIN Visual:" in line and "vin" not in meta:
            meta["vin"] = g(r'No\. VIN Visual:\s*(\S+)') or "–"
        if "Total M.O. Hojal" in line:
            m2=re.search(r'\$([\d,]+\.?\d*)',line)
            if m2: meta["total_mo"]=parse_price(m2.group(1))
        if re.match(r'^TIEMPO\s+M\.O\b',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["tiempo_mo_pintura"]=parse_price(m2.group(1))
        if re.match(r'^TIEMPO\s+PREPARACION',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["tiempo_prep_pintura"]=parse_price(m2.group(1))
        if re.match(r'^TOTAL\s+M\.O\.\s+PINTURA',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["total_mo_pintura"]=parse_price(m2.group(1))
        if re.match(r'^MATERIALES\s+POR\s+SUPERFICIE',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["mat_por_superficie"]=parse_price(m2.group(1))
        if re.match(r'^CONSTANTE\s+MATERIAL',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["constante_material"]=parse_price(m2.group(1))
        if re.match(r'^TOTAL\s+MATERIALES',s):
            m2=re.search(r'\$([\d,]+\.?\d*)',s)
            if m2: meta["total_materiales"]=parse_price(m2.group(1))

    NONE,MO_PREHDR,MO_WAIT,MO,PIN_WAIT,PIN = range(6)
    state = NONE

    for line in lines:
        s = line.strip()
        if not s: continue

        if RE_PIN_HDR.match(s): state=PIN_WAIT; continue
        if state==PIN_WAIT:
            if RE_PIN_TBL.search(s): state=PIN
            continue
        if state==PIN:
            if RE_PIN_STP.match(s): state=NONE; continue
            if s.startswith("-"): continue
            m=RE_ITEM.match(s)
            if m and m.group(1).upper() not in SKIP:
                pin_items.append({"nr":m.group(1),"descripcion":m.group(2).strip(),
                                  "precio":parse_price(m.group(3))})
            continue

        if RE_MO_A.match(s): state=MO_PREHDR; continue
        if state==MO_PREHDR:
            if RE_MO_B.match(s): state=MO_WAIT; continue
            if RE_MO_TBL.search(s): state=MO; continue
            continue
        if state==MO_WAIT:
            if RE_MO_TBL.search(s): state=MO
            continue
        if state==MO:
            if RE_MO_STOP.match(s): state=NONE; continue
            if RE_PIN_HDR.match(s): state=PIN_WAIT; continue
            mt=RE_TIEMPO.match(s)
            if mt:
                desc=mt.group(1).strip()
                mo_items.append({"nr":"","descripcion":desc,
                                 "precio":parse_price(mt.group(2)),
                                 "categoria":clasificar_mo(desc)}); continue
            m=RE_ITEM.match(s)
            if m and m.group(1).upper() not in SKIP:
                nr   = m.group(1)
                desc = m.group(2).strip()
                if nr == "1000":
                    cat = "Hojalatería/Refacción"
                else:
                    cat = clasificar_mo(desc)
                mo_items.append({"nr":nr,"descripcion":desc,
                                 "precio":parse_price(m.group(3)),
                                 "categoria":cat})

    return {"mo":mo_items, "pintura":pin_items, "meta":meta}

# ══════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO COMPLETO POR PDF
# ══════════════════════════════════════════════════════════════════════════════

def procesar_pdf(pdf_file) -> dict:
    n_orden = Path(pdf_file.name).stem
    contenido = pdf_file.read()
    texto = ""
    with pdfplumber.open(io.BytesIO(contenido)) as pdf:
        for page in pdf.pages:
            texto += (page.extract_text() or "") + "\n"

    refacciones, descuento_pct = extraer_refacciones(texto)
    mo_data = extraer_mo_pintura(contenido)

    return {
        "n_orden":      n_orden,
        "refacciones":  refacciones,
        "descuento_pct":descuento_pct,
        "mo":           mo_data["mo"],
        "pintura":      mo_data["pintura"],
        "meta":         mo_data["meta"],
    }

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL UNIFICADO
# ══════════════════════════════════════════════════════════════════════════════

def build_excel(all_orders: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Valuación Completa"

    # Column widths
    ws.column_dimensions["A"].width = 11   # N° Orden
    ws.column_dimensions["B"].width = 20   # Tipo
    ws.column_dimensions["C"].width = 12   # NR/Pos.
    ws.column_dimensions["D"].width = 46   # Descripción
    ws.column_dimensions["E"].width = 15   # Precio Lista
    ws.column_dimensions["F"].width = 11   # Desc %
    ws.column_dimensions["G"].width = 13   # Desc $
    ws.column_dimensions["H"].width = 15   # Precio Final

    # Styles
    THIN  = Side(style="thin",   color="CCCCCC")
    THICK = Side(style="medium", color="666666")
    BDR   = Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
    BDR_T = Border(left=THIN,right=THIN,top=THICK,bottom=THIN)
    CTR   = Alignment(horizontal="center", vertical="center")
    LFT   = Alignment(horizontal="left",   vertical="center")
    RGT   = Alignment(horizontal="right",  vertical="center")
    MONEY = '"$"#,##0.00'
    PCT   = '0%'

    # Fills
    F_HDR    = PatternFill("solid", fgColor="0F3460")   # header cols
    F_GRP    = PatternFill("solid", fgColor="1A1A2E")   # order group header
    F_REF1   = PatternFill("solid", fgColor="D9E1F2")   # refacciones even
    F_REF2   = PatternFill("solid", fgColor="EEF2FF")   # refacciones odd
    F_HOJ1   = PatternFill("solid", fgColor="D6E4F7")   # hojalatería even
    F_HOJ2   = PatternFill("solid", fgColor="EBF3FB")   # hojalatería odd
    F_MEC1   = PatternFill("solid", fgColor="E2EFDA")   # mecánica even
    F_MEC2   = PatternFill("solid", fgColor="F2F9ED")   # mecánica odd
    F_PIN1   = PatternFill("solid", fgColor="FADADD")   # pintura even
    F_PIN2   = PatternFill("solid", fgColor="FDEEF0")   # pintura odd
    F_SUB    = PatternFill("solid", fgColor="16213E")   # subtotals
    F_RES    = PatternFill("solid", fgColor="1B3A2D")   # resumen pintura
    F_TOT    = PatternFill("solid", fgColor="0B2A40")   # total MO+Mat
    F_GRAND  = PatternFill("solid", fgColor="0B0B1A")   # gran total

    W  = Font(color="FFFFFF", bold=True, size=10)
    WS = Font(color="FFFFFF", bold=True, size=9)
    WI = Font(color="DDDDDD", italic=True, size=9)
    S9 = Font(size=9)

    # ── Column headers ────────────────────────────────────────────────────────
    COLS = ["N° ORDEN","TIPO","NR / POS.","DESCRIPCIÓN",
            "PRECIO LISTA","DCTO %","DCTO $","PRECIO FINAL"]
    ws.row_dimensions[1].height = 22
    for col, h in enumerate(COLS, 1):
        c = ws.cell(1, col, h)
        c.font=W; c.fill=F_HDR; c.alignment=CTR; c.border=BDR
    ws.freeze_panes = "A2"

    row = 2
    grand_total_refs = []   # row ranges for refacciones precio final
    grand_total_mo   = 0.0
    grand_total_pin  = 0.0

    def put(r, col, val, fill, font=None, fmt=None, align=LFT, bdr=BDR):
        c = ws.cell(r, col, val)
        c.fill=fill; c.border=bdr; c.alignment=align
        c.font = font or S9
        if fmt: c.number_format=fmt
        return c

    def summary_row(label, val, fill, font, cols=8, is_first=False, money_col=8):
        nonlocal row
        bdr = BDR_T if is_first else BDR
        ws.row_dimensions[row].height = 16
        ws.merge_cells(f"A{row}:{chr(64+cols-1)}{row}")
        put(row, 1, f"  {label}", fill, font, align=LFT, bdr=bdr)
        for c in range(2, cols):
            ws.cell(row,c).fill=fill; ws.cell(row,c).border=bdr
        put(row, cols, val, fill, font, MONEY, RGT, bdr)
        row += 1

    for order in all_orders:
        n_ord       = order["n_orden"]
        refacciones = order["refacciones"]
        desc_pct    = order["descuento_pct"]
        mo          = order["mo"]
        pintura     = order["pintura"]
        meta        = order["meta"]
        veh         = f"{meta.get('fabricante','')} {meta.get('modelo','')}".strip()

        # ── Order group header ────────────────────────────────────────────────
        ws.merge_cells(f"A{row}:H{row}")
        ws.row_dimensions[row].height = 18
        c = ws.cell(row, 1,
            f"  N° ORDEN: {n_ord}   |   {veh}   |   "
            f"VIN: {meta.get('vin','–')}   |   Taller: {meta.get('taller','–')}")
        c.font=W; c.fill=F_GRP; c.alignment=LFT
        c.border=Border(left=THICK,right=THICK,top=THICK,bottom=THICK)
        row += 1

        # ── 1. REFACCIONES ───────────────────────────────────────────────────
        ref_start = row
        for i, p in enumerate(refacciones):
            ws.row_dimensions[row].height = 16
            fill = F_REF1 if i%2==0 else F_REF2
            precio       = p["precio"]
            desc_monto   = round(precio * desc_pct / 100, 2)
            precio_final = round(precio - desc_monto, 2)

            put(row,1,n_ord,      fill,align=CTR)
            put(row,2,"Refacción",fill,align=LFT)
            put(row,3,"",         fill,align=CTR)
            put(row,4,p["descripcion"],fill)
            put(row,5,precio,     fill,fmt=MONEY,align=RGT)
            put(row,6,desc_pct/100 if desc_pct else 0, fill,fmt=PCT,align=CTR)
            put(row,7,desc_monto, fill,fmt=MONEY,align=RGT)
            put(row,8,precio_final,fill,fmt=MONEY,align=RGT)
            row += 1

        ref_end = row - 1
        if refacciones:
            grand_total_refs.append((ref_start, ref_end))
            ws.merge_cells(f"A{row}:G{row}")
            put(row,1,f"  SUBTOTAL Refacciones  ·  N° Orden {n_ord}",F_SUB,WS,align=LFT,bdr=BDR_T)
            for c in range(2,8): ws.cell(row,c).fill=F_SUB; ws.cell(row,c).border=BDR_T
            put(row,8,f"=SUM(H{ref_start}:H{ref_end})",F_SUB,WS,fmt=MONEY,align=RGT,bdr=BDR_T)
            ws.row_dimensions[row].height=16; row+=1

        # ── 2. HOJALATERÍA ───────────────────────────────────────────────────
        hoj = [x for x in mo if x.get("categoria")=="Hojalatería"]
        for i, item in enumerate(hoj):
            ws.row_dimensions[row].height=16
            fill = F_HOJ1 if i%2==0 else F_HOJ2
            put(row,1,n_ord,       fill,align=CTR)
            put(row,2,"Hojalatería",fill)
            put(row,3,item["nr"],  fill,align=CTR)
            put(row,4,item["descripcion"],fill)
            put(row,5,item["precio"],fill,fmt=MONEY,align=RGT)
            for c in (6,7,8): put(row,c,"",fill)
            row+=1

        # ── 2b. HOJALATERÍA/REFACCIÓN (NR=1000) ─────────────────────────────
        F_REF_HOJ1 = PatternFill("solid", fgColor="FFF2CC")   # amarillo claro even
        F_REF_HOJ2 = PatternFill("solid", fgColor="FFFDE7")   # amarillo muy claro odd
        hoj_ref = [x for x in mo if x.get("categoria")=="Hojalatería/Refacción"]
        for i, item in enumerate(hoj_ref):
            ws.row_dimensions[row].height=16
            fill = F_REF_HOJ1 if i%2==0 else F_REF_HOJ2
            put(row,1,n_ord,                    fill,align=CTR)
            put(row,2,"Hojalatería/Refacción",  fill)
            put(row,3,item["nr"],               fill,align=CTR)
            put(row,4,item["descripcion"],      fill)
            put(row,5,item["precio"],           fill,fmt=MONEY,align=RGT)
            for c in (6,7,8): put(row,c,"",fill)
            row+=1

        # ── 3. MECÁNICA ──────────────────────────────────────────────────────
        mec = [x for x in mo if x.get("categoria")=="Mecánica"]
        for i, item in enumerate(mec):
            ws.row_dimensions[row].height=16
            fill = F_MEC1 if i%2==0 else F_MEC2
            put(row,1,n_ord,      fill,align=CTR)
            put(row,2,"Mecánica", fill)
            put(row,3,item["nr"], fill,align=CTR)
            put(row,4,item["descripcion"],fill)
            put(row,5,item["precio"],fill,fmt=MONEY,align=RGT)
            for c in (6,7,8): put(row,c,"",fill)
            row+=1

        # MO subtotal
        if mo:
            total_mo = meta.get("total_mo", sum(x["precio"] for x in mo))
            ws.merge_cells(f"A{row}:G{row}")
            put(row,1,f"  SUBTOTAL M.O. Hojal/Mecánica  ·  N° Orden {n_ord}",F_SUB,WS,align=LFT,bdr=BDR_T)
            for c in range(2,8): ws.cell(row,c).fill=F_SUB; ws.cell(row,c).border=BDR_T
            put(row,8,total_mo,F_SUB,WS,fmt=MONEY,align=RGT,bdr=BDR_T)
            ws.row_dimensions[row].height=16; row+=1
            grand_total_mo += total_mo

        # ── 4. PINTURA ───────────────────────────────────────────────────────
        for i, item in enumerate(pintura):
            ws.row_dimensions[row].height=16
            fill = F_PIN1 if i%2==0 else F_PIN2
            put(row,1,n_ord,          fill,align=CTR)
            put(row,2,"Pintura",      fill)
            put(row,3,item["nr"],     fill,align=CTR)
            put(row,4,item["descripcion"],fill)
            put(row,5,item["precio"], fill,fmt=MONEY,align=RGT)
            for c in (6,7,8): put(row,c,"",fill)
            row+=1

        # Pintura resumen block
        if pintura:
            t_mo_p  = meta.get("tiempo_mo_pintura",  sum(x["precio"] for x in pintura))
            t_prep  = meta.get("tiempo_prep_pintura", 0.0)
            t_mo_tot= meta.get("total_mo_pintura",    t_mo_p+t_prep)
            mat_sup = meta.get("mat_por_superficie",  0.0)
            constante=meta.get("constante_material",  0.0)
            total_mat=meta.get("total_materiales",    mat_sup+constante)
            total_mo_mat = t_mo_tot + total_mat

            summary_row("Tiempo M.O.",              t_mo_p,     F_SUB, WS, is_first=True)
            summary_row("Tiempo Preparación",       t_prep,     F_SUB, WI)
            summary_row("Total M.O. Pintura",       t_mo_tot,   F_SUB, WS)
            summary_row("Materiales por Superficie",mat_sup,    F_RES, WI, is_first=True)
            summary_row("Constante Material",       constante,  F_RES, WI)
            summary_row("Total Materiales",         total_mat,  F_RES, WS)
            summary_row(f"TOTAL M.O. Y MATERIALES  ·  N° Orden {n_ord}",
                        total_mo_mat, F_TOT, WS, is_first=True)
            grand_total_pin += total_mo_mat

        row += 1  # spacer between orders

    # ── Gran Total ────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 22
    ws.merge_cells(f"A{row}:G{row}")
    ct = ws.cell(row,1,"  GRAN TOTAL  (Refacciones + M.O. Hojal/Mec. + Pintura)")
    ct.font=Font(bold=True,size=11,color="FFFFFF"); ct.fill=F_GRAND
    ct.alignment=LFT
    ct.border=Border(left=THICK,right=THICK,top=THICK,bottom=THICK)
    for c in range(2,8):
        ws.cell(row,c).fill=F_GRAND
        ws.cell(row,c).border=Border(left=THICK,right=THICK,top=THICK,bottom=THICK)

    # Sum all refacciones precio final ranges + MO + Pintura
    if grand_total_refs:
        suma_refs = "+".join(f"SUM(H{s}:H{e})" for s,e in grand_total_refs)
        gran_total_val = f"={suma_refs}+{grand_total_mo+grand_total_pin}"
    else:
        gran_total_val = grand_total_mo + grand_total_pin

    cv = ws.cell(row, 8, gran_total_val)
    cv.font=Font(bold=True,size=11,color="FFFFFF"); cv.fill=F_GRAND
    cv.number_format=MONEY; cv.alignment=RGT
    cv.border=Border(left=THICK,right=THICK,top=THICK,bottom=THICK)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### ⚙️ Info")
    st.markdown("**N° Orden:** nombre del archivo PDF")
    st.markdown("**Extrae por orden:**")
    st.markdown("🔩 Refacciones (con descuento)")
    st.markdown("🛠️ Hojalatería")
    st.markdown("🔧 Mecánica")
    st.markdown("🎨 Pintura de Carrocería")
    st.markdown("---")
    st.markdown("**Formato:** AUDATEX / GNP / AXA")

uploaded = st.file_uploader(
    "📂 Sube uno o varios PDFs AUDATEX",
    type=["pdf"], accept_multiple_files=True,
)

if not uploaded:
    st.info("ℹ️ Sube al menos un PDF para comenzar.")
    st.stop()

all_orders = []
prog = st.progress(0, text="Procesando…")
for i, f in enumerate(sorted(uploaded, key=lambda x: x.name)):
    prog.progress((i+1)/len(uploaded), text=f"Procesando: {f.name}")
    all_orders.append(procesar_pdf(f))
prog.empty()

# ── Métricas ──────────────────────────────────────────────────────────────────
total_ref = sum(
    p["precio"]*(1-o["descuento_pct"]/100)
    for o in all_orders for p in o["refacciones"]
)
total_mo  = sum(o["meta"].get("total_mo", sum(x["precio"] for x in o["mo"])) for o in all_orders)
total_pin = sum(
    o["meta"].get("total_mo_pintura",0) + o["meta"].get("total_materiales",0)
    for o in all_orders
)
n_ref = sum(len(o["refacciones"]) for o in all_orders)
n_mo  = sum(len(o["mo"])          for o in all_orders)
n_pin = sum(len(o["pintura"])     for o in all_orders)

cols = st.columns(4)
for col,label,val in [
    (cols[0], "📄 PDFs",              len(uploaded)),
    (cols[1], "🔩 Refacciones",       n_ref),
    (cols[2], "🛠️ Partidas M.O.",    n_mo),
    (cols[3], "🎨 Partidas Pintura",  n_pin),
]:
    col.markdown(f'<div class="metric-box"><div class="metric-value">{val}</div>'
                 f'<div class="metric-label">{label}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Vista previa por orden ────────────────────────────────────────────────────
for order in all_orders:
    n_ord = order["n_orden"]
    meta  = order["meta"]
    veh   = f"{meta.get('fabricante','')} {meta.get('modelo','')}".strip()

    with st.expander(f"📄 N° Orden {n_ord}  —  {veh}", expanded=False):
        mc1, mc2, mc3 = st.columns(3)
        mc1.markdown(f"**Expediente:** {meta.get('expediente','–')}")
        mc2.markdown(f"**Taller:** {meta.get('taller','–')}")
        mc3.markdown(f"**VIN:** {meta.get('vin','–')}")
        st.divider()

        # Refacciones
        if order["refacciones"]:
            st.markdown("**🔩 Refacciones**")
            dp = order["descuento_pct"]
            rows_ref = [{
                "N° Orden": n_ord,
                "Descripción": p["descripcion"],
                "Precio Lista": f"${p['precio']:,.2f}",
                "Dcto %": f"{dp:.0f}%",
                "Dcto $": f"${p['precio']*dp/100:,.2f}",
                "Precio Final": f"${p['precio']*(1-dp/100):,.2f}",
            } for p in order["refacciones"]]
            import pandas as pd
            st.dataframe(pd.DataFrame(rows_ref), use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)

        # MO
        with c1:
            if order["mo"]:
                st.markdown("**🛠️ Mano de Obra**")
                import pandas as pd
                df_mo = pd.DataFrame([{
                    "Sección": x.get("categoria","Hojalatería"),
                    "NR/Pos.": x["nr"],
                    "Trabajo": x["descripcion"],
                    "Precio": f"${x['precio']:,.2f}",
                } for x in order["mo"]])
                st.dataframe(df_mo, use_container_width=True, hide_index=True)
                t = meta.get("total_mo", sum(x["precio"] for x in order["mo"]))
                n_mec     = sum(1 for x in order["mo"] if x.get("categoria")=="Mecánica")
                n_hoj_ref = sum(1 for x in order["mo"] if x.get("categoria")=="Hojalatería/Refacción")
                n_hoj     = len(order["mo"]) - n_mec - n_hoj_ref
                st.caption(f"🔧 Mecánica: {n_mec}  ·  🛠️ Hojalatería: {n_hoj}  ·  🔩 Hojal/Refacción: {n_hoj_ref}")
                st.success(f"**Total M.O.: ${t:,.2f}**")

        # Pintura
        with c2:
            if order["pintura"]:
                st.markdown("**🎨 Pintura**")
                import pandas as pd
                df_pin = pd.DataFrame([{
                    "NR/Pos.": x["nr"],
                    "Descripción": x["descripcion"],
                    "Precio": f"${x['precio']:,.2f}",
                } for x in order["pintura"]])
                st.dataframe(df_pin, use_container_width=True, hide_index=True)
                t_mo_p   = meta.get("tiempo_mo_pintura", sum(x["precio"] for x in order["pintura"]))
                t_prep   = meta.get("tiempo_prep_pintura", 0.0)
                t_mat    = meta.get("total_materiales", 0.0)
                t_mo_tot = meta.get("total_mo_pintura", t_mo_p+t_prep)
                st.info(f"M.O.: **${t_mo_p:,.2f}** | Prep.: **${t_prep:,.2f}** | Total M.O.: **${t_mo_tot:,.2f}**")
                st.info(f"Materiales: **${t_mat:,.2f}**")
                st.error(f"**Total M.O.+Mat.: ${t_mo_tot+t_mat:,.2f}**")

# ── Descarga ──────────────────────────────────────────────────────────────────
st.markdown("---")
xlsx = build_excel(all_orders)
ordenes = "_".join(o["n_orden"] for o in all_orders)
nombre  = f"valuacion_{ordenes}.xlsx" if len(all_orders)<=5 else f"valuacion_{len(all_orders)}_ordenes.xlsx"

st.download_button(
    "📥 Descargar Excel completo (Refacciones + M.O. + Pintura)",
    data=xlsx,
    file_name=nombre,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
