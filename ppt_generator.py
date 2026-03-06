"""
KPI INTERLOG - Generador de PowerPoint Profesional
Diciembre 2025
"""

import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np
from collections import defaultdict, Counter
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import copy

# ============================================================
# PALETA DE COLORES - Dark Executive / Premium
# ============================================================
C = {
    'bg_dark':    RGBColor(0x0D, 0x1B, 0x2A),   # Azul muy oscuro
    'bg_mid':     RGBColor(0x1B, 0x2B, 0x3E),   # Azul medio
    'accent':     RGBColor(0x00, 0xC9, 0xA7),   # Teal brillante
    'accent2':    RGBColor(0x00, 0x8B, 0x74),   # Teal oscuro
    'white':      RGBColor(0xFF, 0xFF, 0xFF),
    'light_gray': RGBColor(0xB0, 0xBE, 0xC5),
    'mid_gray':   RGBColor(0x5E, 0x6E, 0x7F),
    'card_bg':    RGBColor(0x16, 0x28, 0x3A),   # Card background
    'verde':      RGBColor(0x00, 0xC9, 0xA7),
    'naranja':    RGBColor(0xFF, 0x8C, 0x42),
    'rojo':       RGBColor(0xFF, 0x3D, 0x5E),
    'kpi_gold':   RGBColor(0xFF, 0xD7, 0x00),
}

# Hex strings para matplotlib
HEX = {
    'bg_dark':    '#0D1B2A',
    'bg_mid':     '#1B2B3E',
    'accent':     '#00C9A7',
    'accent2':    '#008B74',
    'white':      '#FFFFFF',
    'light_gray': '#B0BEC5',
    'mid_gray':   '#5E6E7F',
    'card_bg':    '#16283A',
    'verde':      '#00C9A7',
    'naranja':    '#FF8C42',
    'rojo':       '#FF3D5E',
    'kpi_gold':   '#FFD700',
    'bar_main':   '#00C9A7',
    'bar_alt':    '#0E4D8C',
}

W = 13.33  # slide width inches (LAYOUT_WIDE)
H = 7.5    # slide height inches

# ============================================================
# HELPERS
# ============================================================

def rgb_hex(r, g, b):
    return RGBColor(r, g, b)

def set_bg(slide, rgb: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb

def add_rect(slide, x, y, w, h, fill_rgb, alpha=None, line=False, line_rgb=None, line_w=1):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if not line:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_rgb or fill_rgb
        shape.line.width = Pt(line_w)
    return shape

def add_text(slide, text, x, y, w, h, size=12, bold=False, color=None, align='left', italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT if align == 'left' else (PP_ALIGN.CENTER if align == 'center' else PP_ALIGN.RIGHT)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color or C['white']
    run.font.name = 'Calibri'
    return txBox

def fig_to_img(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', 
                facecolor=HEX['bg_dark'], edgecolor='none')
    buf.seek(0)
    return buf

def add_divider(slide, x, y, w):
    """Línea divisoria teal"""
    add_rect(slide, x, y, w, 0.03, C['accent'])

# ============================================================
# CHART GENERATORS
# ============================================================

def make_bar_chart(labels, op_vals, kpi_vals, title, w_in=5.5, h_in=2.8, show_kpi_line=False):
    n = len(labels)
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])

    # Ancho adaptivo equilibrado: elegante con pocas o muchas barras
    bar_width = min(0.35, max(0.15, 0.9 / max(n, 1)))
    padding = max(0.5, 1.2 / max(n, 1))
    

    x = np.arange(n)
    bar_colors = [HEX['accent'], HEX['naranja'], HEX['rojo'], '#7C3AED', '#0E4D8C']
    colors = [bar_colors[i % len(bar_colors)] for i in range(n)]

    bars = ax.bar(x, op_vals, color=colors, width=bar_width, zorder=3,
                  edgecolor='none', linewidth=0, align='center')

    max_val = max(op_vals) if op_vals else 1

    for bar, val, color in zip(bars, op_vals, colors):
        # Valor encima de la barra
        ax.text(bar.get_x() + bar.get_width()/2,
                bar.get_height() + max_val * 0.04,
                str(int(val)),
                ha='center', va='bottom', color=HEX['white'],
                fontsize=9, fontweight='bold', fontfamily='DejaVu Sans')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, color=HEX['light_gray'], fontsize=8,
                       fontfamily='DejaVu Sans', rotation=0)
    ax.set_yticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color(HEX['mid_gray'])
    ax.spines['bottom'].set_linewidth(0.5)
    ax.tick_params(axis='x', colors=HEX['light_gray'], length=0, pad=4)
    ax.set_ylim(0, max_val * 1.35)
    ax.yaxis.set_visible(False)

    # Líneas guía sutiles horizontales
    ax.yaxis.grid(True, color=HEX['mid_gray'], alpha=0.15, linewidth=0.5, zorder=0)
    ax.set_axisbelow(True)

    # KPI % debajo del label
    for i, (lbl, kv) in enumerate(zip(labels, kpi_vals)):
        if kv is not None:
            pct = kv if isinstance(kv, str) else f"{kv*100:.1f}%"
            ax.text(i, -max_val * 0.15, pct, ha='center', va='top',
                    color=HEX['kpi_gold'], fontsize=7.5, fontweight='bold',
                    fontfamily='DejaVu Sans')

    ax.set_xlim(-padding, n - 1 + padding)

    plt.tight_layout(pad=0.5)
    return fig


def make_gauge(pct, label, w_in=3.0, h_in=2.2, color=None):
    """Gauge semicircular para KPI único"""
    color = color or (HEX['accent'] if pct >= 95 else (HEX['naranja'] if pct >= 80 else HEX['rojo']))
    fig, ax = plt.subplots(figsize=(w_in, h_in), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    theta_bg = np.linspace(np.pi, 0, 200)
    ax.plot(theta_bg, [0.85]*200, color='#1E3A50', linewidth=20, solid_capstyle='round', zorder=1)
    theta_val = np.linspace(np.pi, np.pi - (pct/100)*np.pi, 200)
    ax.plot(theta_val, [0.85]*200, color=color, linewidth=20, solid_capstyle='round', zorder=2)
    ax.text(0, 0.05, f"{pct:.1f}%", ha='center', va='center', fontsize=20,
            fontweight='bold', color=color, transform=ax.transData)
    ax.text(0, -0.35, label, ha='center', va='center', fontsize=8,
            color=HEX['light_gray'], transform=ax.transData)
    ax.text(0, -0.62, "Target: 95%", ha='center', va='center', fontsize=7,
            color=HEX['mid_gray'], transform=ax.transData)
    ax.set_ylim(0, 1)
    ax.axis('off')
    plt.tight_layout(pad=0.1)
    return fig


def make_donut_kpi(labels, values, colors, center_text, center_sub, w_in=3.2, h_in=3.0):
    """Donut con texto central y leyenda"""
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    wedges, _ = ax.pie(values, labels=None, colors=colors, startangle=90,
                       wedgeprops=dict(width=0.52, edgecolor=HEX['bg_mid'], linewidth=2.5),
                       counterclock=False)
    ax.text(0, 0.12, center_text, ha='center', va='center', fontsize=20,
            fontweight='bold', color=HEX['white'])
    ax.text(0, -0.25, center_sub, ha='center', va='center', fontsize=8,
            color=HEX['light_gray'])
    total = sum(values)
    for i, (lbl, val, col) in enumerate(zip(labels, values, colors)):
        pct = val/total*100 if total else 0
        y = 0.55 - i*0.38
        ax.plot([0.72, 0.82], [y, y], color=col, linewidth=2.5, solid_capstyle='round')
        ax.text(0.87, y+0.08, f"{lbl}", ha='left', va='center', fontsize=7.5, color=HEX['light_gray'])
        ax.text(0.87, y-0.12, f"{int(val)}  ({pct:.0f}%)", ha='left', va='center', fontsize=7, color=col, fontweight='bold')
    ax.set_xlim(-1.1, 1.6)
    ax.axis('equal')
    plt.tight_layout(pad=0.1)
    return fig


def make_lollipop(labels, values, kpi_pcts, w_in=5.5, h_in=2.8, accent_idx=0):
    """Lollipop horizontal elegante"""
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    y = np.arange(len(labels))
    colors = [HEX['accent'] if i == accent_idx else HEX['naranja'] if 'PEND' in str(labels[i]).upper() else HEX['accent2'] for i in range(len(labels))]
    max_v = max(values) if values else 1
    for i, (val, col) in enumerate(zip(values, colors)):
        ax.plot([0, val], [i, i], color=col, linewidth=2.5, alpha=0.5, zorder=2)
        ax.scatter([val], [i], color=col, s=150, zorder=3, edgecolors='white', linewidth=1.5)
        ax.text(val + max_v*0.03, i, str(int(val)), va='center', ha='left',
                color=HEX['white'], fontsize=10, fontweight='bold')
        if kpi_pcts:
            ax.text(max_v*1.22, i, f"{kpi_pcts[i]*100:.1f}%", va='center', ha='right',
                    color=HEX['kpi_gold'], fontsize=8.5, fontweight='bold')
    ax.set_yticks(y)
    ax.set_yticklabels(labels, color=HEX['light_gray'], fontsize=9.5)
    ax.set_xlim(-max_v*0.05, max_v*1.35)
    ax.set_ylim(-0.6, len(labels)-0.4)
    ax.invert_yaxis()
    ax.set_xticks([])
    for sp in ['top','right','bottom','left']:
        ax.spines[sp].set_visible(False)
    ax.tick_params(length=0)
    ax.grid(axis='x', color=HEX['card_bg'], alpha=0.5, zorder=0)
    plt.tight_layout(pad=0.3)
    return fig


def make_hbar(labels, values, kpi_pcts, w_in=5.5, h_in=2.6, bar_color=None):
    """Barras horizontales con fondo oscuro y KPI"""
    bar_color = bar_color or HEX['accent']
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    y = np.arange(len(labels))
    max_v = max(values) if values else 1
    for i, (val, lbl) in enumerate(zip(values, labels)):
        ax.barh(i, max_v*1.05, color='#1E3A50', height=0.55, zorder=1, edgecolor='none')
        c = HEX['naranja'] if 'PEND' in str(lbl).upper() else bar_color
        ax.barh(i, val, color=c, height=0.55, zorder=2, edgecolor='none')
        ax.text(val + max_v*0.02, i, str(int(val)), va='center', ha='left',
                color=HEX['white'], fontsize=10, fontweight='bold', zorder=3)
        if kpi_pcts:
            ax.text(max_v*1.18, i, f"{kpi_pcts[i]*100:.1f}%", va='center', ha='right',
                    color=HEX['kpi_gold'], fontsize=8.5, fontweight='bold', zorder=3)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, color=HEX['light_gray'], fontsize=9)
    ax.set_xlim(0, max_v*1.3)
    ax.set_ylim(-0.6, len(labels)-0.4)
    ax.invert_yaxis()
    ax.set_xticks([])
    for sp in ['top','right','bottom','left']:
        ax.spines[sp].set_visible(False)
    ax.tick_params(length=0)
    plt.tight_layout(pad=0.3)
    return fig


def make_canal_donut(by_canal, via_label, w_in=3.0, h_in=2.8):
    """Donut compacto para distribución de canales por vía"""
    keys = ['VERDE','NARANJA','ROJO']
    vals = [by_canal.get(k,0) for k in keys]
    colors = [HEX['verde'],HEX['naranja'],HEX['rojo']]
    labels = ['Verde','Naranja','Rojo']
    total = sum(vals)
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    filt = [(v,c,l) for v,c,l in zip(vals,colors,labels) if v>0]
    if not filt:
        ax.text(0.5,0.5,'Sin datos',ha='center',va='center',color=HEX['mid_gray'],fontsize=10,transform=ax.transAxes)
        ax.axis('off')
        plt.tight_layout(pad=0.1)
        return fig
    fv,fc,fl = zip(*filt)
    ax.pie(fv, colors=fc, startangle=90,
           wedgeprops=dict(width=0.5, edgecolor=HEX['bg_mid'], linewidth=2), counterclock=False)
    ax.text(0,0.1,str(total),ha='center',va='center',fontsize=18,fontweight='bold',color=HEX['white'])
    ax.text(0,-0.28,via_label,ha='center',va='center',fontsize=8,color=HEX['light_gray'])
    for i,(v,c,l) in enumerate(zip(fv,fc,fl)):
        x_pos = -0.55 + i*0.55
        ax.text(x_pos,-0.85,f"{l}",ha='center',va='center',fontsize=6.5,color=c)
        ax.text(x_pos,-1.1,str(v),ha='center',va='center',fontsize=9,fontweight='bold',color=HEX['white'])
    ax.set_ylim(-1.3,1.1)
    ax.axis('equal')
    plt.tight_layout(pad=0.1)
    return fig


def make_scatter_tiempo(items, limite, w_in=5.5, h_in=2.5):
    """Scatter de tiempos de cada operación vs límite"""
    hs_vals = [i['hs'] for i in items if i['hs'] is not None]
    if not hs_vals:
        return None
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    x = np.arange(len(hs_vals))
    colors = [HEX['rojo'] if h > limite else HEX['accent'] for h in hs_vals]
    ax.scatter(x, hs_vals, c=colors, s=60, zorder=3, edgecolors='white', linewidth=0.6)
    ax.axhline(y=limite, color=HEX['naranja'], linewidth=1.5, linestyle='--', alpha=0.85, zorder=2)
    ax.fill_between([-0.5, len(hs_vals)-0.5], 0, limite, color=HEX['accent'], alpha=0.05, zorder=1)
    ax.fill_between([-0.5, len(hs_vals)-0.5], limite, max(hs_vals)*1.15, color=HEX['rojo'], alpha=0.05, zorder=1)
    ax.text(len(hs_vals)-0.3, limite + max(hs_vals)*0.05, f'Límite: {limite}hs',
            ha='right', va='bottom', color=HEX['naranja'], fontsize=7.5, fontstyle='italic')
    ax.set_xlim(-0.5, len(hs_vals)-0.5)
    ax.set_ylim(0, max(hs_vals)*1.2)
    ax.set_xticks([])
    ax.set_yticks([0, limite, max(hs_vals)])
    ax.set_yticklabels([f'0',f'{limite}hs',f'{max(hs_vals):.0f}hs'], color=HEX['light_gray'], fontsize=7.5)
    for sp in ['top','right','bottom']:
        ax.spines[sp].set_visible(False)
    ax.spines['left'].set_color(HEX['mid_gray'])
    ax.tick_params(length=0)
    in_c = sum(1 for h in hs_vals if h <= limite)
    out_c = len(hs_vals)-in_c
    ax.text(0.02,0.95,f'● IN: {in_c}',transform=ax.transAxes,color=HEX['accent'],fontsize=8,va='top',fontweight='bold')
    ax.text(0.18,0.95,f'● OUT: {out_c}',transform=ax.transAxes,color=HEX['rojo'],fontsize=8,va='top',fontweight='bold')
    plt.tight_layout(pad=0.3)
    return fig


def make_stacked_100(labels, verde_vals, naranja_vals, rojo_vals, w_in=5.5, h_in=2.5):
    """Barras apiladas 100% para composición de canales"""
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    x = np.arange(len(labels))
    totals = [v+n+r for v,n,r in zip(verde_vals,naranja_vals,rojo_vals)]
    pv = [v/t*100 if t else 0 for v,t in zip(verde_vals,totals)]
    pn = [n/t*100 if t else 0 for n,t in zip(naranja_vals,totals)]
    pr = [r/t*100 if t else 0 for r,t in zip(rojo_vals,totals)]
    width = 0.5
    ax.bar(x, pv, width, color=HEX['verde'], edgecolor='none', zorder=3)
    ax.bar(x, pn, width, bottom=pv, color=HEX['naranja'], edgecolor='none', zorder=3)
    b3_bot = [a+b for a,b in zip(pv,pn)]
    ax.bar(x, pr, width, bottom=b3_bot, color=HEX['rojo'], edgecolor='none', zorder=3)
    for i in range(len(labels)):
        if pv[i] > 8: ax.text(i, pv[i]/2, f'{verde_vals[i]}', ha='center', va='center', color='white', fontsize=8, fontweight='bold')
        if pn[i] > 8: ax.text(i, pv[i]+pn[i]/2, f'{naranja_vals[i]}', ha='center', va='center', color='white', fontsize=8, fontweight='bold')
        if pr[i] > 8: ax.text(i, pv[i]+pn[i]+pr[i]/2, f'{rojo_vals[i]}', ha='center', va='center', color='white', fontsize=8, fontweight='bold')
        ax.text(i, 103, f'n={totals[i]}', ha='center', va='bottom', color=HEX['light_gray'], fontsize=7.5)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, color=HEX['light_gray'], fontsize=9)
    ax.set_ylim(0, 115)
    ax.set_yticks([0,50,100])
    ax.set_yticklabels(['0%','50%','100%'], color=HEX['mid_gray'], fontsize=7.5)
    for sp in ['top','right']: ax.spines[sp].set_visible(False)
    ax.spines['left'].set_color(HEX['mid_gray'])
    ax.spines['bottom'].set_color(HEX['mid_gray'])
    ax.tick_params(length=0)
    ax.grid(axis='y', color=HEX['mid_gray'], alpha=0.15, zorder=0)
    ax.legend(handles=[mpatches.Patch(color=c,label=l) for c,l in [(HEX['verde'],'Verde'),(HEX['naranja'],'Naranja'),(HEX['rojo'],'Rojo')]],
              frameon=False, fontsize=7.5, labelcolor=HEX['light_gray'], loc='upper right', ncol=3)
    plt.tight_layout(pad=0.3)
    return fig


def make_donut(labels, values, colors, w_in=2.8, h_in=2.8):
    """Donut clásico (fallback)"""
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    fig.patch.set_facecolor(HEX['bg_mid'])
    ax.set_facecolor(HEX['bg_mid'])
    wedges, texts, autotexts = ax.pie(values, labels=None, colors=colors, autopct='%1.0f%%',
                                       startangle=90, wedgeprops=dict(width=0.55, edgecolor=HEX['bg_mid'], linewidth=2),
                                       pctdistance=0.75)
    for at in autotexts:
        at.set_color(HEX['white']); at.set_fontsize(8); at.set_fontweight('bold')
    ax.axis('equal')
    plt.tight_layout(pad=0.1)
    return fig

# ============================================================
# TABLE HELPER
# ============================================================

def add_table_styled(slide, data, x, y, w, h, header_rgb=None):
    """Tabla estilizada con fondo oscuro"""
    rows = len(data)
    cols = len(data[0])
    col_w = w / cols

    header_rgb = header_rgb or C['accent2']

    for ri, row in enumerate(data):
        for ci, cell_val in enumerate(row):
            cx = x + ci * col_w
            cy = y + ri * (h / rows)
            cw = col_w
            ch = h / rows

            if ri == 0:
                bg = header_rgb
                fc = C['white']
                fb = True
                fs = 7.5
            else:
                bg = C['card_bg'] if ri % 2 == 0 else C['bg_mid']
                fc = C['white'] if str(cell_val) not in ['IN', '100,00%'] else C['accent']
                fb = False
                fs = 7.5

            add_rect(slide, cx, cy, cw - 0.01, ch - 0.005, bg)
            add_text(slide, str(cell_val), cx + 0.05, cy + 0.01,
                     cw - 0.1, ch - 0.02, size=fs, bold=fb, color=fc, align='center')


# ============================================================
# KPI CARD
# ============================================================

def add_kpi_card(slide, x, y, w, h, value, label, sub=None, color=None):
    color = color or C['accent']
    # Card background
    add_rect(slide, x, y, w, h, C['card_bg'])
    # Accent line top
    add_rect(slide, x, y, w, 0.04, color)
    # Value
    add_text(slide, value, x, y + 0.12, w, 0.45, size=28, bold=True, color=color, align='center')
    # Label
    add_text(slide, label, x, y + 0.52, w, 0.22, size=8, bold=False, color=C['light_gray'], align='center')
    if sub:
        add_text(slide, sub, x, y + 0.72, w, 0.18, size=7, bold=False, color=C['mid_gray'], align='center')


# ============================================================
# SLIDE BUILDERS
# ============================================================

def build_slide_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    # Geometric accent shapes
    add_rect(slide, 0, 0, 0.5, H, C['bg_mid'])
    add_rect(slide, 0, 0, 0.5, 3.5, C['accent'])
    add_rect(slide, W - 4.5, 0, 4.5, H, C['bg_mid'])

    # Diagonal accent
    from pptx.util import Emu
    shape = slide.shapes.add_shape(1, Inches(W-4.8), Inches(0), Inches(0.4), Inches(H))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C['accent2']
    shape.line.fill.background()

    # Logo area - teal circle
    circle = slide.shapes.add_shape(9, Inches(W-3.5), Inches(H/2-1.2), Inches(2.4), Inches(2.4))
    circle.fill.solid()
    circle.fill.fore_color.rgb = C['accent']
    circle.line.fill.background()

    # INTERLOG text in circle
    add_text(slide, "INTERLOG", W-3.5, H/2-0.45, 2.4, 0.5,
             size=13, bold=True, color=C['bg_dark'], align='center')
    add_text(slide, "Comercio Exterior", W-3.5, H/2+0.1, 2.4, 0.3,
             size=7.5, bold=False, color=C['bg_dark'], align='center')

    # Main title
    add_text(slide, "KPI", 0.9, 1.6, 7, 1.1, size=72, bold=True, color=C['white'], align='left')
    add_text(slide, "FASA / FSM", 0.9, 2.6, 8, 0.8, size=36, bold=False, color=C['accent'], align='left')

    # Divider
    add_rect(slide, 0.9, 3.55, 5.5, 0.04, C['accent'])

    # Subtitle
    add_text(slide, "DICIEMBRE 2025", 0.9, 3.75, 7, 0.4,
             size=16, bold=False, color=C['light_gray'], align='left')
    add_text(slide, "Reporte Mensual de Indicadores de Desempeño", 0.9, 4.2, 8, 0.35,
             size=10, bold=False, color=C['mid_gray'], align='left')


def build_slide_ofi_lib(prs, razon_items_ofi, razon_items_lib, nombre, kpi_data):
    """Slide de Oficialización + Liberación para una razón social"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    # Header bar
    add_rect(slide, 0, 0, W, 0.75, C['bg_mid'])
    add_rect(slide, 0, 0, 0.08, 0.75, C['accent'])
    add_text(slide, f"MODAL MULTIMODAL · {nombre}", 0.2, 0.1, 10, 0.55,
             size=18, bold=True, color=C['white'], align='left')
    add_text(slide, "DICIEMBRE 2025", W-2.5, 0.22, 2.3, 0.3,
             size=9, bold=False, color=C['accent'], align='right')

    # ---- OFICIALIZACIÓN (izquierda) ----
    ofi_items = razon_items_ofi
    total_ofi = len(ofi_items)
    desvios_ofi = [i for i in ofi_items if i['desvio']]
    sin_desvio_ofi = [i for i in ofi_items if not i['desvio']]
    pct_ofi = len(sin_desvio_ofi)/total_ofi*100 if total_ofi else 100

    # Panel OFI
    add_rect(slide, 0.3, 0.85, 6.1, 5.6, C['bg_mid'])
    add_rect(slide, 0.3, 0.85, 6.1, 0.04, C['accent2'])
    add_text(slide, "OFICIALIZACIÓN", 0.5, 0.92, 4, 0.35, size=11, bold=True, color=C['accent'], align='left')

    # Gauge KPI
    fig_ofi = make_gauge(pct_ofi, "KPI IN", w_in=3.2, h_in=2.2)
    img_ofi = fig_to_img(fig_ofi)
    plt.close(fig_ofi)
    slide.shapes.add_picture(img_ofi, Inches(0.5), Inches(1.25), Inches(3.0), Inches(2.1))

    # Lollipop OFI
    ofi_lbl = ['S/DESVIO']
    ofi_ops = [len(sin_desvio_ofi)]
    ofi_kpis_l = [len(sin_desvio_ofi)/total_ofi if total_ofi else 0]
    if desvios_ofi:
        ofi_lbl.append('PENDIENTE')
        ofi_ops.append(len(desvios_ofi))
        ofi_kpis_l.append(len(desvios_ofi)/total_ofi)
    fig_lol_ofi = make_lollipop(ofi_lbl, ofi_ops, ofi_kpis_l, w_in=3.2, h_in=1.8)
    img_lol_ofi = fig_to_img(fig_lol_ofi)
    plt.close(fig_lol_ofi)
    slide.shapes.add_picture(img_lol_ofi, Inches(3.6), Inches(1.35), Inches(2.6), Inches(1.9))

    # Stats OFI
    add_kpi_card(slide, 0.45, 3.5, 1.4, 0.95, str(total_ofi), "Total ops", None, C['accent2'])
    add_kpi_card(slide, 2.05, 3.5, 1.4, 0.95, str(len(sin_desvio_ofi)), "S/Desvío", None, C['verde'])
    add_kpi_card(slide, 3.65, 3.5, 1.4, 0.95, str(len(desvios_ofi)), "Desvíos", None, C['rojo'] if desvios_ofi else C['mid_gray'])

    # Tabla OFI
    ofi_table = [["Rango", "OP", "KPI"], ["IN", str(total_ofi), "100,00%"],
                 ["S/DESVIO", str(len(sin_desvio_ofi)), f"{pct_ofi:.2f}%"]]
    if desvios_ofi:
        ofi_table.append(["PENDIENTE", str(len(desvios_ofi)), f"{len(desvios_ofi)/total_ofi*100:.2f}%"])
    ofi_table.append(["Total", str(total_ofi), "100,00%"])
    add_table_styled(slide, ofi_table, 0.35, 4.6, 6.0, 1.1)

    # ---- LIBERACIÓN (derecha) ----
    lib_items = razon_items_lib
    total_lib = len(lib_items)
    desvios_lib = [i for i in lib_items if i['desvio']]
    sin_desvio = [i for i in lib_items if not i['desvio']]
    params_count = Counter()
    params_count['S/DESVIO'] = len(sin_desvio)
    if desvios_lib:
        params_count['PENDIENTE'] = len(desvios_lib)
    lib_labels = list(params_count.keys())
    lib_ops = [params_count[k] for k in lib_labels]
    lib_kpis = [v/total_lib for v in lib_ops]
    pct_lib_in = (params_count.get('S/DESVIO', 0) / total_lib * 100) if total_lib else 0

    # Panel LIB
    add_rect(slide, 6.9, 0.85, 6.1, 5.6, C['bg_mid'])
    add_rect(slide, 6.9, 0.85, 6.1, 0.04, C['naranja'])
    add_text(slide, "LIBERACIÓN", 7.1, 0.92, 4, 0.35, size=11, bold=True, color=C['naranja'], align='left')

    # Gauge LIB
    kpi_color = C['accent'] if pct_lib_in >= 95 else (C['naranja'] if pct_lib_in >= 80 else C['rojo'])
    fig_lib_g = make_gauge(pct_lib_in, "KPI IN", w_in=3.2, h_in=2.2,
                           color=HEX['accent'] if pct_lib_in >= 95 else (HEX['naranja'] if pct_lib_in >= 80 else HEX['rojo']))
    img_lib_g = fig_to_img(fig_lib_g)
    plt.close(fig_lib_g)
    slide.shapes.add_picture(img_lib_g, Inches(7.1), Inches(1.25), Inches(3.0), Inches(2.1))

    # Lollipop LIB
    fig_lol_lib = make_lollipop(lib_labels, lib_ops, lib_kpis, w_in=3.2, h_in=1.8)
    img_lol_lib = fig_to_img(fig_lol_lib)
    plt.close(fig_lol_lib)
    slide.shapes.add_picture(img_lol_lib, Inches(10.3), Inches(1.35), Inches(2.6), Inches(1.9))

    # Stats LIB
    add_kpi_card(slide, 7.05, 3.5, 1.4, 0.95, str(total_lib), "Total ops", None, C['accent2'])
    add_kpi_card(slide, 8.65, 3.5, 1.4, 0.95, str(len(sin_desvio)), "S/Desvío", None, C['verde'])
    add_kpi_card(slide, 10.25, 3.5, 1.4, 0.95, str(len(desvios_lib)), "Desvíos", None, C['rojo'] if desvios_lib else C['mid_gray'])

    # Tabla LIB
    lib_table = [["Rango", "OP", "KPI"], ["IN", str(total_lib), "100,00%"]]
    for lbl, op, kpi in zip(lib_labels, lib_ops, lib_kpis):
        lib_table.append([lbl, str(op), f"{kpi*100:.2f}%"])
    lib_table.append(["Total", str(total_lib), "100,00%"])
    add_table_styled(slide, lib_table, 6.95, 4.6, 6.0, 1.1, header_rgb=C['accent2'])

    # Footer
    add_rect(slide, 0, H-0.28, W, 0.28, C['bg_mid'])
    add_text(slide, "INTERLOG · Comercio Exterior · Reporte Mensual", 0.3, H-0.24, 8, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='left')
    add_text(slide, "Diciembre 2025", W-2.2, H-0.24, 2, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='right')


def build_slide_verde_avion(prs, lib_fasa, lib_fsm, nombre_razon, items_razon):
    """Slide Canal Verde Avión para una razón social"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    # Header
    add_rect(slide, 0, 0, W, 0.75, C['bg_mid'])
    add_rect(slide, 0, 0, 0.08, 0.75, C['verde'])
    add_text(slide, f"CANAL VERDE · VÍA AÉREA · {nombre_razon}", 0.2, 0.1, 11, 0.55,
             size=18, bold=True, color=C['white'], align='left')
    add_text(slide, "DICIEMBRE 2025", W-2.5, 0.22, 2.3, 0.3,
             size=9, bold=False, color=C['accent'], align='right')

    # Filtrar canal verde avion de esta razón
    va_items = [i for i in items_razon if i['via'] == 'AVION' and i['canal'] == 'VERDE']
    total = len(va_items)
    sin_desvio = [i for i in va_items if not i['desvio']]
    con_desvio = [i for i in va_items if i['desvio']]

    pct_in = len(sin_desvio)/total*100 if total else 0

    # KPI cards top
    kpi_color = C['accent'] if pct_in >= 95 else (C['naranja'] if pct_in >= 80 else C['rojo'])
    add_kpi_card(slide, 0.3, 0.9, 2.5, 1.2, f"{pct_in:.1f}%", "KPI IN", f"Target: 95%", kpi_color)
    add_kpi_card(slide, 3.1, 0.9, 2.5, 1.2, str(total), "Total Operaciones", "Canal Verde Avión", C['accent2'])
    add_kpi_card(slide, 5.9, 0.9, 2.5, 1.2, str(len(sin_desvio)), "S/Desvío", "Dentro del rango", C['verde'])
    add_kpi_card(slide, 8.7, 0.9, 2.5, 1.2, str(len(con_desvio)), "Con Desvío", "Fuera del rango", C['rojo'] if con_desvio else C['mid_gray'])

    # Panel izquierdo — Scatter de tiempos
    add_rect(slide, 0.3, 2.25, 7.8, 3.2, C['bg_mid'])
    add_rect(slide, 0.3, 2.25, 7.8, 0.04, C['verde'])
    add_text(slide, "TIEMPOS POR OPERACIÓN vs. LÍMITE (24 hs hábiles)", 0.5, 2.32, 6, 0.3,
             size=9, bold=True, color=C['verde'], align='left')

    fig_scatter = make_scatter_tiempo(va_items, 24, w_in=6.5, h_in=2.4)
    if fig_scatter:
        img_scatter = fig_to_img(fig_scatter)
        plt.close(fig_scatter)
        slide.shapes.add_picture(img_scatter, Inches(0.5), Inches(2.65), Inches(7.4), Inches(2.55))

    # Panel derecho — Donut distribución
    add_rect(slide, 8.4, 2.25, 4.6, 3.2, C['bg_mid'])
    add_rect(slide, 8.4, 2.25, 4.6, 0.04, C['accent2'])
    add_text(slide, "DISTRIBUCIÓN", 8.6, 2.32, 3.5, 0.3, size=9, bold=True, color=C['accent'], align='left')

    d_labels = ['S/DESVIO', 'PENDIENTE'] if con_desvio else ['S/DESVIO']
    d_values = [len(sin_desvio), len(con_desvio)] if con_desvio else [len(sin_desvio)]
    d_colors = [HEX['accent'], HEX['naranja']] if con_desvio else [HEX['accent']]

    fig_donut = make_donut_kpi(d_labels, d_values, d_colors,
                                str(total), "ops total", w_in=3.5, h_in=2.6)
    img_donut = fig_to_img(fig_donut)
    plt.close(fig_donut)
    slide.shapes.add_picture(img_donut, Inches(8.5), Inches(2.6), Inches(4.2), Inches(2.65))

    # Footer
    add_rect(slide, 0, H-0.28, W, 0.28, C['bg_mid'])
    add_text(slide, "INTERLOG · Comercio Exterior · Reporte Mensual", 0.3, H-0.24, 8, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='left')
    add_text(slide, "Diciembre 2025", W-2.2, H-0.24, 2, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='right')


def build_slide_maritimos(prs, lib_data, ofi_data):
    """Slide Modal Marítimos - todas las razones"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    add_rect(slide, 0, 0, W, 0.75, C['bg_mid'])
    add_rect(slide, 0, 0, 0.08, 0.75, RGBColor(0x0E, 0x4D, 0x8C))
    add_text(slide, "MODAL MARÍTIMOS · TODAS LAS RAZONES SOCIALES", 0.2, 0.1, 11, 0.55,
             size=18, bold=True, color=C['white'], align='left')
    add_text(slide, "DICIEMBRE 2025", W-2.5, 0.22, 2.3, 0.3,
             size=9, bold=False, color=C['accent'], align='right')

    all_lib = list(lib_data['FASA']) + list(lib_data['FSM'])
    all_ofi = list(ofi_data['FASA']) + list(ofi_data['FSM'])

    mar_lib = [i for i in all_lib if i['via'] == 'MARITIMO']
    mar_ofi = [i for i in all_ofi if i.get('via') == 'MARITIMO']

    total_ml = len(mar_lib)
    total_mo = len(mar_ofi)

    # KPI cards
    add_kpi_card(slide, 0.3, 0.9, 3.0, 1.1, str(total_ml), "Liberaciones Marítimas", None, RGBColor(0x0E, 0x4D, 0x8C))
    add_kpi_card(slide, 3.6, 0.9, 3.0, 1.1, str(total_mo), "Oficializaciones Marítimas", None, C['accent2'])

    # Canales marítimos
    by_canal_lib = Counter(i['canal'] for i in mar_lib)
    add_kpi_card(slide, 7.0, 0.9, 1.8, 1.1, str(by_canal_lib.get('VERDE', 0)), "Verde", None, C['verde'])
    add_kpi_card(slide, 9.0, 0.9, 1.8, 1.1, str(by_canal_lib.get('NARANJA', 0)), "Naranja", None, C['naranja'])
    add_kpi_card(slide, 11.0, 0.9, 1.8, 1.1, str(by_canal_lib.get('ROJO', 0)), "Rojo", None, C['rojo'])

    # Gráfico LIB marítimo
    sin_dev_ml = [i for i in mar_lib if not i['desvio']]
    con_dev_ml = [i for i in mar_lib if i['desvio']]
    labels_ml = ['S/DESVIO']
    ops_ml = [len(sin_dev_ml)]
    kpis_ml = [len(sin_dev_ml)/total_ml if total_ml else 0]
    if con_dev_ml:
        labels_ml.append('PENDIENTE')
        ops_ml.append(len(con_dev_ml))
        kpis_ml.append(len(con_dev_ml)/total_ml)

    fig_ml = make_bar_chart(labels_ml, ops_ml, kpis_ml, '', w_in=4.2, h_in=2.4)
    img_ml = fig_to_img(fig_ml)
    plt.close(fig_ml)

    add_rect(slide, 0.3, 2.15, 5.9, 2.8, C['bg_mid'])
    add_rect(slide, 0.3, 2.15, 5.9, 0.04, C['naranja'])
    add_text(slide, "LIBERACIÓN MARÍTIMA", 0.5, 2.22, 4, 0.3, size=10, bold=True, color=C['naranja'], align='left')
    slide.shapes.add_picture(img_ml, Inches(0.5), Inches(2.55), Inches(4.0), Inches(2.2))

    pct_ml = len(sin_dev_ml)/total_ml*100 if total_ml else 100
    add_kpi_card(slide, 4.7, 2.3, 1.3, 0.95, f"{pct_ml:.0f}%", "KPI IN", None,
                 C['accent'] if pct_ml >= 95 else C['naranja'])

    # Tabla LIB
    table_ml = [["Rango", "OP", "KPI"], ["IN", str(total_ml), "100,00%"]]
    for lbl, op, kpi in zip(labels_ml, ops_ml, kpis_ml):
        table_ml.append([lbl, str(op), f"{kpi*100:.2f}%"])
    table_ml.append(["Total", str(total_ml), "100,00%"])
    add_table_styled(slide, table_ml, 0.35, 5.05, 5.8, 1.0)

    # Gráfico OFI marítimo
    sin_dev_mo = [i for i in mar_ofi if not i['desvio']]
    con_dev_mo = [i for i in mar_ofi if i['desvio']]
    labels_mo = ['S/DESVIO']
    ops_mo = [len(sin_dev_mo)]
    kpis_mo = [len(sin_dev_mo)/total_mo if total_mo else 0]
    if con_dev_mo:
        labels_mo.append('PENDIENTE')
        ops_mo.append(len(con_dev_mo))
        kpis_mo.append(len(con_dev_mo)/total_mo)

    fig_mo = make_bar_chart(labels_mo, ops_mo, kpis_mo, '', w_in=4.2, h_in=2.4)
    img_mo = fig_to_img(fig_mo)
    plt.close(fig_mo)

    add_rect(slide, 6.9, 2.15, 6.0, 2.8, C['bg_mid'])
    add_rect(slide, 6.9, 2.15, 6.0, 0.04, C['accent'])
    add_text(slide, "OFICIALIZACIÓN MARÍTIMA", 7.1, 2.22, 4, 0.3, size=10, bold=True, color=C['accent'], align='left')
    slide.shapes.add_picture(img_mo, Inches(7.1), Inches(2.55), Inches(4.0), Inches(2.2))

    pct_mo = len(sin_dev_mo)/total_mo*100 if total_mo else 100
    add_kpi_card(slide, 11.4, 2.3, 1.3, 0.95, f"{pct_mo:.0f}%", "KPI IN", None,
                 C['accent'] if pct_mo >= 95 else C['naranja'])

    table_mo = [["Rango", "OP", "KPI"], ["IN", str(total_mo), "100,00%"]]
    for lbl, op, kpi in zip(labels_mo, ops_mo, kpis_mo):
        table_mo.append([lbl, str(op), f"{kpi*100:.2f}%"])
    table_mo.append(["Total", str(total_mo), "100,00%"])
    add_table_styled(slide, table_mo, 6.95, 5.05, 5.9, 1.0, header_rgb=C['accent2'])

    add_rect(slide, 0, H-0.28, W, 0.28, C['bg_mid'])
    add_text(slide, "INTERLOG · Comercio Exterior · Reporte Mensual", 0.3, H-0.24, 8, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='left')
    add_text(slide, "Diciembre 2025", W-2.2, H-0.24, 2, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='right')


def build_slide_canales(prs, lib_fasa, lib_fsm):
    """Slide distribución de canales por vía y razón social"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    add_rect(slide, 0, 0, W, 0.75, C['bg_mid'])
    add_rect(slide, 0, 0, 0.08, 0.75, C['kpi_gold'])
    add_text(slide, "DISTRIBUCIÓN DE CANALES · POR VÍA Y RAZÓN SOCIAL", 0.2, 0.1, 11, 0.55,
             size=18, bold=True, color=C['white'], align='left')
    add_text(slide, "DICIEMBRE 2025", W-2.5, 0.22, 2.3, 0.3,
             size=9, bold=False, color=C['accent'], align='right')

    vias = ['AVION', 'MARITIMO', 'CAMION']
    canal_colors = [HEX['verde'], HEX['naranja'], HEX['rojo']]
    canal_labels = ['Verde', 'Naranja', 'Rojo']
    canal_keys = ['VERDE', 'NARANJA', 'ROJO']

    for col_idx, (razon_items, nombre) in enumerate([(lib_fasa, 'FASA'), (lib_fsm, 'FSM')]):
        ox = 0.3 + col_idx * 6.5

        add_rect(slide, ox, 0.85, 6.2, 5.6, C['bg_mid'])
        add_rect(slide, ox, 0.85, 6.2, 0.04, C['accent'] if col_idx == 0 else C['naranja'])
        add_text(slide, nombre, ox + 0.15, 0.92, 5, 0.35,
                 size=13, bold=True,
                 color=C['accent'] if col_idx == 0 else C['naranja'], align='left')

        for row_idx, via in enumerate(vias):
            via_items = [i for i in razon_items if i['via'] == via]
            vy = 1.05 + row_idx * 1.7
            add_text(slide, via, ox + 0.15, vy + 0.05, 5, 0.3,
                     size=9, bold=True, color=C['light_gray'], align='left')

            if not via_items:
                add_text(slide, "Sin operaciones", ox + 0.15, vy + 0.45, 4, 0.3,
                         size=8, bold=False, color=C['mid_gray'], align='left')
                continue

            by_canal = Counter(i['canal'] for i in via_items)
            # Donut compacto por vía
            fig_donut = make_canal_donut(by_canal, f"Total: {len(via_items)}", w_in=2.8, h_in=1.45)
            buf = fig_to_img(fig_donut)
            plt.close(fig_donut)
            slide.shapes.add_picture(buf, Inches(ox + 0.15), Inches(vy + 0.32), Inches(3.0), Inches(1.35))

            # Mini stats horizontales
            for ci, (ckey, clbl, ccolor) in enumerate([('VERDE','V',C['verde']),('NARANJA','N',C['naranja']),('ROJO','R',C['rojo'])]):
                cv = by_canal.get(ckey, 0)
                if cv > 0:
                    add_kpi_card(slide, ox + 3.35 + ci * 0.92, vy + 0.38, 0.85, 0.85,
                                 str(cv), clbl, None, ccolor)

    add_rect(slide, 0, H-0.28, W, 0.28, C['bg_mid'])
    add_text(slide, "INTERLOG · Comercio Exterior · Reporte Mensual", 0.3, H-0.24, 8, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='left')
    add_text(slide, "Diciembre 2025", W-2.2, H-0.24, 2, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='right')


def build_slide_cm(prs, cm_pre_results, cm_apr_results):
    """Slide KPI Certificados Mineros"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    add_rect(slide, 0, 0, W, 0.75, C['bg_mid'])
    add_rect(slide, 0, 0, 0.08, 0.75, RGBColor(0x7C, 0x3A, 0xED))
    add_text(slide, "KPI CERTIFICADOS MINEROS", 0.2, 0.1, 11, 0.55,
             size=18, bold=True, color=C['white'], align='left')
    add_text(slide, "DICIEMBRE 2025", W-2.5, 0.22, 2.3, 0.3,
             size=9, bold=False, color=C['accent'], align='right')

    # ---- PRESENTADOS (izq) ----
    total_pre = len(cm_pre_results)
    desvios_pre = [r for r in cm_pre_results if r['desvio']]
    sin_desvio_pre = [r for r in cm_pre_results if not r['desvio']]
    pct_pre = len(sin_desvio_pre)/total_pre*100 if total_pre else 100

    add_rect(slide, 0.3, 0.85, 6.0, 5.6, C['bg_mid'])
    add_rect(slide, 0.3, 0.85, 6.0, 0.04, RGBColor(0x7C, 0x3A, 0xED))
    add_text(slide, "PRESENTACIÓN", 0.5, 0.92, 4, 0.35, size=12, bold=True,
             color=RGBColor(0x7C, 0x3A, 0xED), align='left')

    # KPI card presentados
    kpi_color_pre = C['accent'] if pct_pre >= 95 else (C['naranja'] if pct_pre >= 80 else C['rojo'])
    add_kpi_card(slide, 0.4, 1.35, 1.6, 1.1, f"{pct_pre:.1f}%", "KPI IN", f"Target: 95%", kpi_color_pre)
    add_kpi_card(slide, 2.2, 1.35, 1.5, 1.1, str(total_pre), "Total Extes.", None, RGBColor(0x7C, 0x3A, 0xED))
    add_kpi_card(slide, 3.9, 1.35, 1.2, 1.1, str(len(sin_desvio_pre)), "IN", None, C['verde'])
    add_kpi_card(slide, 5.3, 1.35, 0.9, 1.1, str(len(desvios_pre)), "OUT", None, C['rojo'] if desvios_pre else C['mid_gray'])

    # Gauge para presentados
    fig_gauge_pre = make_gauge(pct_pre, "KPI IN", w_in=3.5, h_in=2.4,
                                color=HEX['accent'] if pct_pre >= 95 else HEX['naranja'])
    img_gauge_pre = fig_to_img(fig_gauge_pre)
    plt.close(fig_gauge_pre)
    slide.shapes.add_picture(img_gauge_pre, Inches(0.5), Inches(2.6), Inches(3.2), Inches(2.2))

    # Hbar para presentados
    labels_pre = ['IN']
    ops_pre = [len(sin_desvio_pre)]
    kpis_pre = [pct_pre/100]
    if desvios_pre:
        labels_pre.append('PENDIENTE')
        ops_pre.append(len(desvios_pre))
        kpis_pre.append(len(desvios_pre)/total_pre)

    fig_hbar_pre = make_hbar(labels_pre, ops_pre, kpis_pre, w_in=3.8, h_in=2.0,
                              bar_color=HEX['accent'])
    img_hbar_pre = fig_to_img(fig_hbar_pre)
    plt.close(fig_hbar_pre)
    slide.shapes.add_picture(img_hbar_pre, Inches(3.5), Inches(2.7), Inches(2.7), Inches(1.9))

    table_pre = [["Rango", "Extes.", "KPI"]]
    table_pre.append(["in", str(len(sin_desvio_pre)), f"{pct_pre:.2f}%"])
    if desvios_pre:
        table_pre.append(["out", str(len(desvios_pre)), f"{len(desvios_pre)/total_pre*100:.2f}%"])
    table_pre.append(["Total general", str(total_pre), "100,00%"])
    add_table_styled(slide, table_pre, 0.35, 5.2, 5.9, 1.0, header_rgb=RGBColor(0x5B, 0x21, 0xB6))

    # ---- APROBADOS (der) ----
    total_apr = len(cm_apr_results)
    rangos = Counter(r['rango'] for r in cm_apr_results if r['rango'])

    add_rect(slide, 6.9, 0.85, 6.1, 5.6, C['bg_mid'])
    add_rect(slide, 6.9, 0.85, 6.1, 0.04, C['kpi_gold'])
    add_text(slide, "APROBADOS", 7.1, 0.92, 4, 0.35, size=12, bold=True,
             color=C['kpi_gold'], align='left')
    add_text(slide, "(Solo informativo · Sin target)", 9.8, 0.96, 3, 0.25, size=7.5,
             bold=False, color=C['mid_gray'], align='left')

    add_kpi_card(slide, 7.0, 1.35, 2.0, 1.1, str(total_apr), "Total Extes.", None, C['kpi_gold'])
    add_kpi_card(slide, 9.2, 1.35, 1.8, 1.1, str(rangos.get('0 a 7', 0)), "0 a 7 días", None, C['verde'])
    add_kpi_card(slide, 11.1, 1.35, 1.8, 1.1, str(rangos.get('8 a 15', 0)), "8 a 15 días", None, C['naranja'])

    apr_labels = ['0 a 7', '8 a 15', '+15']
    apr_ops = [rangos.get('0 a 7', 0), rangos.get('8 a 15', 0), rangos.get('+15', 0)]
    apr_kpis = [v/total_apr for v in apr_ops]

    fig_apr = make_bar_chart(apr_labels, apr_ops, apr_kpis, '', w_in=4.5, h_in=2.5)
    img_apr = fig_to_img(fig_apr)
    plt.close(fig_apr)
    slide.shapes.add_picture(img_apr, Inches(7.1), Inches(2.6), Inches(4.5), Inches(2.5))

    table_apr = [["Rango", "Extes.", "KPI"]]
    for lbl in ['0 a 7', '8 a 15', '+15']:
        v = rangos.get(lbl, 0)
        table_apr.append([lbl, str(v), f"{v/total_apr*100:.2f}%" if total_apr else "0%"])
    table_apr.append(["Total general", str(total_apr), "100,00%"])
    add_table_styled(slide, table_apr, 6.95, 5.2, 6.0, 1.1, header_rgb=RGBColor(0xB4, 0x8A, 0x00))

    add_rect(slide, 0, H-0.28, W, 0.28, C['bg_mid'])
    add_text(slide, "INTERLOG · Comercio Exterior · Reporte Mensual", 0.3, H-0.24, 8, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='left')
    add_text(slide, "Diciembre 2025", W-2.2, H-0.24, 2, 0.2,
             size=6.5, bold=False, color=C['mid_gray'], align='right')


def build_slide_cierre(prs, mes='MES'):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])

    add_rect(slide, 0, 0, W, H, C['bg_dark'])
    add_rect(slide, W-4.5, 0, 4.5, H, C['bg_mid'])
    add_rect(slide, W-4.8, 0, 0.4, H, C['accent2'])

    circle = slide.shapes.add_shape(9, Inches(W-3.5), Inches(H/2-1.2), Inches(2.4), Inches(2.4))
    circle.fill.solid()
    circle.fill.fore_color.rgb = C['accent']
    circle.line.fill.background()

    add_text(slide, "INTERLOG", W-3.5, H/2-0.45, 2.4, 0.5,
             size=13, bold=True, color=C['bg_dark'], align='center')
    add_text(slide, "Comercio Exterior", W-3.5, H/2+0.1, 2.4, 0.3,
             size=7.5, bold=False, color=C['bg_dark'], align='center')

    add_text(slide, "Gracias", 0.9, 1.8, 7, 1.0, size=60, bold=True, color=C['white'], align='left')
    add_rect(slide, 0.9, 3.0, 5.0, 0.04, C['accent'])
    add_text(slide, f"Reporte Mensual · {mes}", 0.9, 3.2, 8, 0.4,
             size=13, bold=False, color=C['light_gray'], align='left')
    add_text(slide, "INTERLOG · Comercio Exterior", 0.9, 3.75, 8, 0.3,
             size=9.5, bold=False, color=C['mid_gray'], align='left')



# ============================================================
# FUNCIÓN PRINCIPAL — recibe datos procesados, devuelve buffer
# ============================================================

def generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes='MES'):
    """
    Recibe listas de items ya procesados (con parametro/desvio_desc cargados)
    y devuelve un io.BytesIO con el .pptx listo para descargar.
    """
    FASA_KEY = 'FASA'
    FSM_KEY  = 'FSM'

    lib_fasa = [i for i in lib_items if i['nombre'] == FASA_KEY]
    lib_fsm  = [i for i in lib_items if i['nombre'] == FSM_KEY]
    ofi_fasa = [i for i in ofi_items if i['nombre'] == FASA_KEY]
    ofi_fsm  = [i for i in ofi_items if i['nombre'] == FSM_KEY]

    lib_data = {FASA_KEY: lib_fasa, FSM_KEY: lib_fsm}
    ofi_data = {FASA_KEY: ofi_fasa, FSM_KEY: ofi_fsm}

    # Adaptar formato: los items de la app tienen mismos campos que build_ppt
    # solo necesitamos que tengan 'desvio', 'via', 'canal', 'hs', 'limite'
    # que ya vienen del procesamiento

    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    # Portada con mes dinámico
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg_dark'])
    add_rect(slide, 0, 0, 0.5, H, C['bg_mid'])
    add_rect(slide, 0, 0, 0.5, 3.5, C['accent'])
    add_rect(slide, W - 4.5, 0, 4.5, H, C['bg_mid'])
    shape = slide.shapes.add_shape(1, Inches(W-4.8), Inches(0), Inches(0.4), Inches(H))
    shape.fill.solid(); shape.fill.fore_color.rgb = C['accent2']; shape.line.fill.background()
    circle = slide.shapes.add_shape(9, Inches(W-3.5), Inches(H/2-1.2), Inches(2.4), Inches(2.4))
    circle.fill.solid(); circle.fill.fore_color.rgb = C['accent']; circle.line.fill.background()
    add_text(slide, "INTERLOG", W-3.5, H/2-0.45, 2.4, 0.5, size=13, bold=True, color=C['bg_dark'], align='center')
    add_text(slide, "Comercio Exterior", W-3.5, H/2+0.1, 2.4, 0.3, size=7.5, bold=False, color=C['bg_dark'], align='center')
    add_text(slide, "KPI", 0.9, 1.6, 7, 1.1, size=72, bold=True, color=C['white'], align='left')
    add_text(slide, "FASA / FSM", 0.9, 2.6, 8, 0.8, size=36, bold=False, color=C['accent'], align='left')
    add_rect(slide, 0.9, 3.55, 5.5, 0.04, C['accent'])
    add_text(slide, mes.upper(), 0.9, 3.75, 7, 0.4, size=16, bold=False, color=C['light_gray'], align='left')
    add_text(slide, "Reporte Mensual de Indicadores de Desempeño", 0.9, 4.2, 8, 0.35, size=10, bold=False, color=C['mid_gray'], align='left')

    # Slides de contenido
    build_slide_ofi_lib(prs, ofi_fasa, lib_fasa, 'FASA', {})
    build_slide_verde_avion(prs, lib_fasa, lib_fsm, 'FASA', lib_fasa)
    build_slide_ofi_lib(prs, ofi_fsm, lib_fsm, 'FSM', {})
    build_slide_verde_avion(prs, lib_fasa, lib_fsm, 'FSM', lib_fsm)
    build_slide_maritimos(prs, lib_data, ofi_data)
    build_slide_canales(prs, lib_fasa, lib_fsm)
    build_slide_cm(prs, cm_pre_items, cm_apr_items)
    build_slide_cierre(prs, mes)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
