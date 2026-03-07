"""
KPI INTERLOG — Generador de PowerPoint
Usa las imágenes de brand reales de INTERLOG como fondos.
Mantiene la firma: generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes)
"""

import io
import os
import base64
from collections import Counter
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─────────────────────────────────────────────────────────────
# SLIDE DIMENSIONS
# ─────────────────────────────────────────────────────────────
W_IN = 13.33
H_IN = 7.5

# ─────────────────────────────────────────────────────────────
# PALETA — brand INTERLOG
# ─────────────────────────────────────────────────────────────
def rgb(r, g, b): return RGBColor(r, g, b)

TEAL    = rgb(0x3F, 0xA8, 0x9C)
DARK    = rgb(0x1E, 0x2D, 0x35)
GRAY    = rgb(0x5A, 0x64, 0x72)
LGRAY   = rgb(0x9A, 0xA5, 0xAF)
WHITE   = rgb(0xFF, 0xFF, 0xFF)
BG_CARD = rgb(0xF4, 0xF5, 0xF6)
BG_SIDE = rgb(0xF0, 0xF4, 0xF6)
LINE_C  = rgb(0xD8, 0xDD, 0xE2)
RED_K   = rgb(0xC0, 0x39, 0x2B)
ORANGE  = rgb(0xD4, 0x69, 0x0A)
VERDE_C = rgb(0x27, 0xAE, 0x60)
NARANJ_C= rgb(0xE6, 0x7E, 0x22)
ROJO_C  = rgb(0xC0, 0x39, 0x2B)

def kpi_color(pct):
    if pct >= 95: return TEAL
    if pct >= 80: return ORANGE
    return RED_K

# ─────────────────────────────────────────────────────────────
# IMÁGENES DE BRAND — cargadas desde el mismo directorio
# ─────────────────────────────────────────────────────────────
_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_img(filename):
    """Carga imagen; busca primero junto al script, luego en rutas conocidas."""
    candidates = [
        os.path.join(_DIR, filename),
        os.path.join('/app', filename),
        os.path.join(os.getcwd(), filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

BG_PORTADA_PATH   = _load_img('bg_portada.jpg')
BG_CONTENIDO_PATH = _load_img('bg_contenido.jpg')

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def inch(v): return Inches(v)

def add_bg_image(slide, path):
    """Pone una imagen como fondo completo del slide."""
    if path and os.path.exists(path):
        slide.shapes.add_picture(path, inch(0), inch(0), inch(W_IN), inch(H_IN))

def add_rect(slide, x, y, w, h, fill_rgb, line=False, line_rgb=None):
    shape = slide.shapes.add_shape(1, inch(x), inch(y), inch(w), inch(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if not line:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_rgb or fill_rgb
        shape.line.width = Pt(0.5)
    return shape

def add_text(slide, text, x, y, w, h, size=10, bold=False, color=None,
             align='left', valign='top', italic=False):
    txBox = slide.shapes.add_textbox(inch(x), inch(y), inch(w), inch(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    if align == 'center':
        p.alignment = PP_ALIGN.CENTER
    elif align == 'right':
        p.alignment = PP_ALIGN.RIGHT
    else:
        p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color or DARK
    run.font.name = 'Calibri'
    return txBox

def add_hbar(slide, x, y, w, h, pct, color):
    """Barra de progreso horizontal."""
    add_rect(slide, x, y, w, h, rgb(0xE8, 0xEA, 0xEC))
    fw = max(0.02, w * pct / 100)
    add_rect(slide, x, y, fw, h, color)

def add_card(slide, x, y, w, h, accent=None):
    """Card blanca con sombra simulada y acento izquierdo opcional."""
    # sombra
    add_rect(slide, x+0.015, y+0.015, w, h, rgb(0xD0, 0xD4, 0xD8))
    # card
    shape = slide.shapes.add_shape(1, inch(x), inch(y), inch(w), inch(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.color.rgb = rgb(0xE0, 0xE4, 0xE8)
    shape.line.width = Pt(0.5)
    if accent:
        add_rect(slide, x, y, 0.07, h, accent)

def add_slide_title(slide, text):
    """Título principal del slide con línea teal."""
    add_text(slide, text, 0.45, 0.22, W_IN - 0.9, 0.52,
             size=22, bold=True, color=DARK, align='left')
    add_rect(slide, 0.45, 0.76, 3.5, 0.04, TEAL)

def add_sec_header(slide, text, x, y, w):
    add_text(slide, text, x, y, w, 0.28,
             size=9, bold=True, color=TEAL, align='left')

def kpi_card_big(slide, x, y, w, h, val, label, sub, pct):
    col = kpi_color(pct)
    add_card(slide, x, y, w, h, col)
    add_text(slide, val, x+0.1, y+0.06, w-0.18, h*0.52,
             size=26, bold=True, color=col, align='center')
    add_text(slide, label, x+0.1, y+h*0.56, w-0.18, 0.28,
             size=9, bold=True, color=DARK, align='center')
    if sub:
        add_text(slide, sub, x+0.1, y+h*0.56+0.28, w-0.18, 0.22,
                 size=7.5, color=LGRAY, align='center')

def metric_card(slide, x, y, w, h, val, label, col):
    add_card(slide, x, y, w, h)
    add_text(slide, str(val), x+0.06, y+0.06, w-0.12, h*0.62,
             size=24, bold=True, color=col, align='center')
    add_text(slide, label, x+0.06, y+h*0.65, w-0.12, 0.24,
             size=8, bold=True, color=GRAY, align='center')

def via_row(slide, x, y, w, label, total, in_n, out_n, pct,
            verde, naranja, rojo):
    """Fila de vía con KPI + 3 columnas de canales."""
    H_ROW = 0.80
    col = kpi_color(pct)
    add_card(slide, x, y, w, H_ROW, col)

    # Label
    add_text(slide, label, x+0.15, y+0.08, 1.5, 0.28,
             size=11, bold=True, color=DARK)
    add_text(slide, f"{total} ops · {in_n} IN · {out_n} OUT",
             x+0.15, y+0.37, 1.5, 0.22, size=8, color=LGRAY)

    # KPI %
    add_text(slide, f"{pct:.0f}%", x+1.75, y+0.08, 0.95, 0.56,
             size=22, bold=True, color=col, align='center')

    # Canales
    canales = [
        ('VERDE',   verde,   VERDE_C),
        ('NARANJA', naranja, NARANJ_C),
        ('ROJO',    rojo,    ROJO_C),
    ]
    bar_x0 = x + 2.82
    bar_col_w = (w - 3.05) / 3

    for idx, (clabel, cdata, ccol) in enumerate(canales):
        bx = bar_x0 + idx * bar_col_w + 0.04
        bw = bar_col_w - 0.10
        ct, ci, co = cdata

        if ct == 0:
            add_text(slide, f"{clabel}\n—", bx, y+0.08, bw, H_ROW-0.16,
                     size=7.5, color=LGRAY, align='center')
            continue

        bar_pct = (ci / ct) * 100
        add_text(slide, clabel, bx, y+0.08, bw, 0.18,
                 size=7, bold=True, color=ccol, align='left')
        add_hbar(slide, bx, y+0.28, bw, 0.13, bar_pct, ccol)
        add_text(slide, f"{ci}/{ct}", bx, y+0.44, bw*0.5, 0.22,
                 size=9, bold=True, color=DARK)
        add_text(slide, f"{bar_pct:.0f}%", bx+bw*0.5, y+0.44, bw*0.5, 0.22,
                 size=9, bold=True, color=ccol, align='right')


def two_panel_kpi(slide, x, nombre, d_pct, d_total, d_in, d_out,
                  label_tipo, limite_str):
    """Panel izquierdo+derecho para Oficialización / Canal Verde."""
    cw = W_IN/2 - 0.65
    ch = H_IN - 1.85
    add_card(slide, x, 1.0, cw, ch, TEAL)
    col = kpi_color(d_pct)

    # Panel izquierdo gris
    lw = cw * 0.42
    shape = slide.shapes.add_shape(1, inch(x+0.07), inch(1.0), inch(lw), inch(ch))
    shape.fill.solid(); shape.fill.fore_color.rgb = BG_SIDE
    shape.line.fill.background()

    add_text(slide, nombre, x+0.18, 1.18, lw-0.22, 0.32,
             size=15, bold=True, color=DARK, align='left')
    add_text(slide, label_tipo, x+0.18, 1.55, lw-0.22, 0.22,
             size=8, bold=True, color=LGRAY, align='left')
    add_rect(slide, x+0.18, 1.84, lw-0.36, 0.03, LINE_C)

    for ii, (lbl, val) in enumerate([("TARGET", "95%"), ("LÍMITE", limite_str)]):
        iy = 1.98 + ii * 0.75
        add_text(slide, lbl, x+0.18, iy, lw-0.22, 0.22,
                 size=7.5, bold=True, color=LGRAY, align='left')
        add_text(slide, val, x+0.18, iy+0.22, lw-0.22, 0.34,
                 size=12, bold=True, color=DARK, align='left')

    add_rect(slide, x+0.18, 3.60, lw-0.36, 0.03, LINE_C)

    for idx, (st_label, st_val, st_col) in enumerate([
        ("TOTAL", d_total, DARK),
        ("IN",    d_in,    TEAL),
        ("OUT",   d_out,   RED_K if d_out > 0 else TEAL),
    ]):
        sy = 3.75 + idx * 0.70
        add_text(slide, str(st_val), x+0.18, sy, 0.55, 0.36,
                 size=18, bold=True, color=st_col, align='left')
        add_text(slide, st_label, x+0.76, sy+0.10, lw-0.80, 0.26,
                 size=8, bold=True, color=LGRAY, align='left')

    # Panel derecho: KPI
    rx = x + lw + 0.07
    rw = cw - lw - 0.07
    add_text(slide, f"{d_pct:.0f}%", rx, 1.6, rw, 2.0,
             size=52, bold=True, color=col, align='center')
    add_text(slide, "KPI IN", rx, 3.70, rw, 0.24,
             size=9, bold=True, color=LGRAY, align='center')
    add_hbar(slide, rx+0.2, 4.00, rw-0.4, 0.12, d_pct, col)
    add_text(slide, "TARGET 95%", rx, 4.20, rw, 0.22,
             size=8, color=LGRAY, align='center')


# ─────────────────────────────────────────────────────────────
# SLIDES
# ─────────────────────────────────────────────────────────────

def slide_portada(prs, mes):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_PORTADA_PATH)
    add_text(s, "KPI FASA / FSM",  7.8, 2.2, 5.0, 0.75,
             size=28, bold=True, color=DARK, align='center')
    add_text(s, mes.upper(),        7.8, 3.05, 5.0, 0.55,
             size=22, bold=True, color=DARK, align='center')
    add_text(s, "Reporte de Indicadores de Desempeño",
             7.8, 3.72, 5.0, 0.32, size=11, color=DARK, align='center')


def slide_resumen(prs, lib_items, ofi_items, cm_pre_items):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, "RESUMEN EJECUTIVO")

    def kpi_pct(items, nombre=None, desvio_field='desvio'):
        sub = [i for i in items if nombre is None or i.get('nombre') == nombre]
        if not sub: return 0, 0
        sin = [i for i in sub if not i[desvio_field]]
        return len(sin)/len(sub)*100, len(sub)

    fasa_lib_pct, fasa_lib_n = kpi_pct(lib_items, 'FASA')
    fsm_lib_pct,  fsm_lib_n  = kpi_pct(lib_items, 'FSM')
    fasa_ofi_pct, fasa_ofi_n = kpi_pct(ofi_items, 'FASA')
    fsm_ofi_pct,  fsm_ofi_n  = kpi_pct(ofi_items, 'FSM')
    cm_pre_pct,   cm_pre_n   = kpi_pct(cm_pre_items)

    kpis = [
        (f"{fasa_lib_pct:.0f}%", "LIBERACIÓN\nFASA",      f"{fasa_lib_n} operaciones", fasa_lib_pct),
        (f"{fsm_lib_pct:.0f}%",  "LIBERACIÓN\nFSM",       f"{fsm_lib_n} operaciones",  fsm_lib_pct),
        (f"{fasa_ofi_pct:.0f}%", "OFICIALIZACIÓN\nFASA",  f"{fasa_ofi_n} operaciones", fasa_ofi_pct),
        (f"{fsm_ofi_pct:.0f}%",  "OFICIALIZACIÓN\nFSM",   f"{fsm_ofi_n} operaciones",  fsm_ofi_pct),
        (f"{cm_pre_pct:.0f}%",   "CM\nPRESENTADOS",       f"{cm_pre_n} expedientes",   cm_pre_pct),
    ]
    cw, ch, gap = 2.2, 2.4, 0.27
    tw = len(kpis)*cw + (len(kpis)-1)*gap
    sx = (W_IN - tw) / 2
    for i, (val, lbl, sub, pct) in enumerate(kpis):
        kpi_card_big(s, sx + i*(cw+gap), 1.0, cw, ch, val, lbl, sub, pct)

    add_text(s, "Target: 95%  ·  Parámetro ≠ INTERLOG → operación IN",
             sx, 3.60, tw, 0.26, size=8.5, color=LGRAY, align='center')

    # Desvíos
    add_sec_header(s, "DESVÍOS IMPUTABLES A INTERLOG", 0.45, 4.05, W_IN-0.9)

    devs = [
        ("FASA · Liberación",      len([i for i in lib_items  if i['nombre']=='FASA' and i['desvio']]),  fasa_lib_n),
        ("FSM · Liberación",       len([i for i in lib_items  if i['nombre']=='FSM'  and i['desvio']]),  fsm_lib_n),
        ("FASA · Oficialización",  len([i for i in ofi_items  if i['nombre']=='FASA' and i['desvio']]),  fasa_ofi_n),
        ("FSM · Oficialización",   len([i for i in ofi_items  if i['nombre']=='FSM'  and i['desvio']]),  fsm_ofi_n),
    ]
    for i, (lbl, out, total) in enumerate(devs):
        dx = 0.45 + i * 3.1
        add_card(s, dx, 4.42, 2.90, 0.82)
        add_text(s, lbl,           dx+0.12, 4.50, 2.66, 0.22, size=8.5, color=GRAY)
        oc = RED_K if out > 0 else TEAL
        add_text(s, f"{out} OUT",  dx+0.12, 4.70, 1.30, 0.35, size=14, bold=True, color=oc)
        add_text(s, f"de {total} ops", dx+1.45, 4.76, 1.30, 0.26, size=9, color=LGRAY)


def slide_liberacion(prs, lib_items, nombre):
    """Slide liberación por razón social con filas por vía."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, f"{nombre} · LIBERACIÓN")

    items = [i for i in lib_items if i['nombre'] == nombre]
    total = len(items)
    in_n  = len([i for i in items if not i['desvio']])
    out_n = total - in_n
    pct   = in_n/total*100 if total else 0

    add_sec_header(s, "RESUMEN", 0.45, 1.0, 5)
    mets = [
        (total, "TOTAL OPS", DARK),
        (in_n,  "IN",        TEAL),
        (out_n, "OUT",       RED_K if out_n > 0 else TEAL),
        (f"{pct:.0f}%", "KPI", kpi_color(pct)),
    ]
    for i, (val, lbl, col) in enumerate(mets):
        metric_card(s, 0.45 + i*1.48, 1.38, 1.28, 1.0, val, lbl, col)

    add_sec_header(s, "DETALLE POR VÍA", 0.45, 2.60, W_IN-0.9)

    vias_map = {'AVION': '✈ AÉREO', 'CAMION': '🚛 CAMIÓN', 'MARITIMO': '🚢 MARÍTIMO'}
    via_order = [v for v in ['AVION', 'CAMION', 'MARITIMO']
                 if any(i['via'] == v for i in items)]

    for row_idx, via in enumerate(via_order):
        vi = [i for i in items if i['via'] == via]
        vt = len(vi)
        if vt == 0: continue
        vi_in  = len([i for i in vi if not i['desvio']])
        vi_out = vt - vi_in
        vi_pct = vi_in/vt*100

        def canal(c): 
            cv = [i for i in vi if i.get('canal') == c]
            ci = len([i for i in cv if not i['desvio']])
            return (len(cv), ci, len(cv)-ci)

        via_row(s, 0.45, 2.98 + row_idx * 0.93, W_IN-0.9,
                vias_map.get(via, via), vt, vi_in, vi_out, vi_pct,
                canal('VERDE'), canal('NARANJA'), canal('ROJO'))

    add_text(s, "Target: 95%", W_IN-2.5, H_IN-0.90, 2.2, 0.24,
             size=8.5, color=LGRAY, align='right')


def slide_oficializacion(prs, ofi_items):
    """Oficialización FASA + FSM — panel izquierdo/derecho."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, "OFICIALIZACIÓN · FASA + FSM")

    for nombre, x, limite_str in [
        ("FASA", 0.45, "24 hs hábiles"),
        ("FSM",  W_IN/2 + 0.2, "24 hs (48 hs Marítimo)"),
    ]:
        items = [i for i in ofi_items if i['nombre'] == nombre]
        total = len(items)
        in_n  = len([i for i in items if not i['desvio']])
        out_n = total - in_n
        pct   = in_n/total*100 if total else 0
        two_panel_kpi(s, x, nombre, pct, total, in_n, out_n, "OFICIALIZACIÓN", limite_str)


def slide_canal_verde(prs, lib_items):
    """Canal Verde · Vía Aérea — FASA y FSM."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, "CANAL VERDE · VÍA AÉREA")

    for nombre, x in [("FASA", 0.45), ("FSM", W_IN/2 + 0.2)]:
        vi = [i for i in lib_items if i['nombre'] == nombre
              and i.get('via') == 'AVION' and i.get('canal') == 'VERDE']
        total = len(vi)
        in_n  = len([i for i in vi if not i['desvio']])
        out_n = total - in_n
        pct   = in_n/total*100 if total else 0
        two_panel_kpi(s, x, nombre, pct, total, in_n, out_n, "CANAL VERDE AÉREO", "1 día hábil")


def slide_distribucion_canales(prs, lib_items):
    """Distribución de canales — barras horizontales apiladas."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, "DISTRIBUCIÓN DE CANALES · LIBERACIONES")

    filas = []
    for nombre in ['FASA', 'FSM']:
        for via, via_label in [('AVION','AÉREO'), ('CAMION','CAMIÓN'), ('MARITIMO','MARÍTIMO')]:
            vi = [i for i in lib_items if i['nombre'] == nombre and i['via'] == via]
            if not vi: continue
            verde   = len([i for i in vi if i.get('canal') == 'VERDE'])
            naranja = len([i for i in vi if i.get('canal') == 'NARANJA'])
            rojo    = len([i for i in vi if i.get('canal') == 'ROJO'])
            filas.append((f"{nombre} · {via_label}", verde, naranja, rojo, len(vi)))

    BAR_W = W_IN - 4.3
    gap = (H_IN - 1.1 - 0.9 - len(filas) * 0.48) / max(len(filas) - 1, 1)
    row_h = 0.48 + gap

    for i, (lbl, verde, naranja, rojo, total) in enumerate(filas):
        y = 1.0 + i * row_h
        add_text(s, lbl,   0.45, y+0.05, 2.1, 0.28, size=10, bold=True, color=DARK)
        add_text(s, f"{total} ops", 0.45, y+0.34, 2.1, 0.22, size=8, color=LGRAY)

        bx = 2.65; bh = 0.44
        for val, col in [(verde, VERDE_C), (naranja, NARANJ_C), (rojo, ROJO_C)]:
            if val == 0: continue
            fw = BAR_W * (val / total)
            add_rect(s, bx, y, fw, bh, col)
            if fw > 0.45:
                pct_s = f"{val} ({val/total*100:.0f}%)"
                add_text(s, pct_s, bx+0.04, y, fw-0.08, bh,
                         size=8.5, bold=True, color=WHITE, align='center')
            bx += fw

    # Leyenda
    for idx, (col, lbl) in enumerate([(VERDE_C,"Canal Verde"), (NARANJ_C,"Canal Naranja"), (ROJO_C,"Canal Rojo")]):
        lx = 2.65 + idx * 2.8
        add_rect(s, lx, H_IN-0.88, 0.20, 0.14, col)
        add_text(s, lbl, lx+0.28, H_IN-0.92, 2.4, 0.24, size=8, color=GRAY)


def slide_cm(prs, cm_pre_items, cm_apr_items):
    """Certificados Mineros."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_image(s, BG_CONTENIDO_PATH)
    add_slide_title(s, "CERTIFICADOS MINEROS")

    # Presentados
    total_pre = len(cm_pre_items)
    in_pre    = len([i for i in cm_pre_items if not i['desvio']])
    out_pre   = total_pre - in_pre
    pct_pre   = in_pre/total_pre*100 if total_pre else 0

    add_sec_header(s, "PRESENTADOS · KPI DE GESTIÓN", 0.45, 0.96, 7)
    for i, (val, lbl, col) in enumerate([
        (total_pre, "TOTAL",  DARK),
        (in_pre,    "IN",     TEAL),
        (out_pre,   "OUT",    RED_K if out_pre > 0 else TEAL),
        (f"{pct_pre:.0f}%", "KPI", kpi_color(pct_pre)),
    ]):
        metric_card(s, 0.45 + i*1.48, 1.33, 1.28, 0.95, val, lbl, col)

    add_text(s, "Límite: 48 hs hábiles desde TAD Subido",
             0.45, 2.38, 6.0, 0.22, size=8, color=LGRAY)

    # Aprobados
    add_sec_header(s, "APROBADOS · TIEMPO DE APROBACIÓN (informativo)", 0.45, 2.78, W_IN-0.9)
    rangos = Counter(i.get('rango') for i in cm_apr_items)
    total_apr = len(cm_apr_items)
    BAR_W = W_IN - 4.5

    for i, (rango_lbl, col) in enumerate([("0 a 7", VERDE_C), ("8 a 15", NARANJ_C), ("+15", ROJO_C)]):
        val = rangos.get(rango_lbl, 0)
        pct = val/total_apr*100 if total_apr else 0
        y = 3.15 + i * 0.82

        add_text(s, f"{rango_lbl} días", 0.45, y+0.08, 1.7, 0.32,
                 size=9.5, bold=True, color=col, align='right')

        # fondo barra
        add_rect(s, 2.3, y, BAR_W, 0.50, rgb(0xEA, 0xEC, 0xEE))

        if val > 0:
            fw = max(0.05, BAR_W * pct / 100)
            add_rect(s, 2.3, y, fw, 0.50, col)
            add_text(s, f"{val} exp · {pct:.0f}%", 2.35, y+0.08, fw-0.1, 0.34,
                     size=10, bold=True, color=WHITE, align='left')
        else:
            add_text(s, "Sin expedientes", 2.35, y+0.08, BAR_W-0.1, 0.34,
                     size=9, color=LGRAY)

    add_text(s, f"Total aprobados: {total_apr} expedientes",
             0.45, H_IN-0.88, W_IN-0.9, 0.24, size=8.5, color=LGRAY)


# ─────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────────────────────

def generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes='MES'):
    """
    Genera el PowerPoint con brand INTERLOG.
    Recibe los items ya procesados del Streamlit y devuelve un io.BytesIO.
    """
    prs = Presentation()
    prs.slide_width  = Inches(W_IN)
    prs.slide_height = Inches(H_IN)

    slide_portada(prs, mes)
    slide_resumen(prs, lib_items, ofi_items, cm_pre_items)
    slide_liberacion(prs, lib_items, 'FASA')
    slide_liberacion(prs, lib_items, 'FSM')
    slide_oficializacion(prs, ofi_items)
    slide_canal_verde(prs, lib_items)
    slide_distribucion_canales(prs, lib_items)
    slide_cm(prs, cm_pre_items, cm_apr_items)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
