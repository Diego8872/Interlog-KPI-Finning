import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
from collections import Counter
import io
import pickle

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="INTERLOG · KPI Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
    background-color: #0A1628;
    color: #F0F4F8;
}
.stApp { background: #0A1628; }

h1, h2, h3 {
    font-family: 'Barlow Condensed', sans-serif;
    letter-spacing: 1px;
}

/* Cards métricas */
.metric-card {
    background: #132236;
    border: 1px solid rgba(0,201,167,0.15);
    border-top: 3px solid #00C9A7;
    border-radius: 8px;
    padding: 1.2rem;
    text-align: center;
}
.metric-value {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2.4rem; font-weight: 800;
    color: #00C9A7; line-height: 1;
}
.metric-value.orange { color: #FF8C42; }
.metric-value.red    { color: #FF3D5E; }
.metric-value.gold   { color: #FFD060; }
.metric-label {
    font-size: 0.7rem; font-weight: 600;
    color: #6B8099; letter-spacing: 1px;
    text-transform: uppercase; margin-top: 0.3rem;
}
.metric-sub {
    font-size: 0.75rem; color: #9AB0C4; margin-top: 0.2rem;
}

/* Section headers */
.section-header {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem; font-weight: 700;
    color: #00C9A7; letter-spacing: 2px;
    text-transform: uppercase;
    border-left: 3px solid #00C9A7;
    padding-left: 0.8rem; margin: 1.5rem 0 0.8rem;
}

/* Steps */
.step-badge {
    display: inline-block;
    background: #00C9A7; color: #0A1628;
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 800; font-size: 0.9rem;
    padding: 2px 10px; border-radius: 4px;
    margin-right: 0.5rem;
}

/* Upload zone */
.upload-card {
    background: #132236;
    border: 1px dashed rgba(0,201,167,0.3);
    border-radius: 8px; padding: 1.5rem;
    text-align: center;
}

/* Alert */
.alert-info {
    background: rgba(0,201,167,0.1);
    border: 1px solid rgba(0,201,167,0.3);
    border-radius: 6px; padding: 0.8rem 1rem;
    font-size: 0.85rem; color: #9AB0C4;
}
.alert-warn {
    background: rgba(255,140,66,0.1);
    border: 1px solid rgba(255,140,66,0.4);
    border-radius: 6px; padding: 0.8rem 1rem;
    font-size: 0.85rem; color: #FF8C42;
}
.alert-success {
    background: rgba(0,201,167,0.15);
    border: 1px solid #00C9A7;
    border-radius: 6px; padding: 0.8rem 1rem;
    font-size: 0.85rem; color: #00C9A7;
    font-weight: 600;
}

/* Tablas */
.dataframe { background: #132236 !important; }
thead tr th { background: #007A65 !important; color: white !important; }

/* Buttons */
.stButton > button {
    background: #00C9A7; color: #0A1628;
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 700; font-size: 1rem;
    letter-spacing: 1px; border: none;
    border-radius: 6px; padding: 0.6rem 2rem;
    transition: all 0.2s;
}
.stButton > button:hover { background: #00E0BA; transform: translateY(-1px); }

/* Sidebar */
.css-1d391kg { background: #0F2040; }

/* Hide streamlit branding */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
FASA = 'FINNING ARGENTINA SOCIEDAD ANO'
FSM  = 'FINNING SOLUCIONES MINERAS SA'

LIMITES_LIB = {
    'AVION':    {'VERDE': 24, 'NARANJA': 72, 'ROJO': 96},
    'MARITIMO': {'VERDE': 72, 'NARANJA': 96, 'ROJO': 120},
    'CAMION':   {'VERDE': 24, 'NARANJA': 48, 'ROJO': 72},
}

COLORS = {
    'accent':  '#00C9A7',
    'orange':  '#FF8C42',
    'red':     '#FF3D5E',
    'gold':    '#FFD060',
    'bg':      '#132236',
    'bg2':     '#1A2E48',
    'gray':    '#6B8099',
    'verde':   '#00C9A7',
    'naranja': '#FF8C42',
    'rojo':    '#FF3D5E',
}

CHART_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(19,34,54,0.8)',
    font=dict(family='Barlow, sans-serif', color='#9AB0C4'),
    margin=dict(l=10, r=10, t=30, b=10),
    xaxis=dict(gridcolor='rgba(107,128,153,0.15)', linecolor='rgba(107,128,153,0.3)'),
    yaxis=dict(gridcolor='rgba(107,128,153,0.15)', linecolor='rgba(107,128,153,0.3)'),
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def parse_date(val):
    if isinstance(val, datetime): return val
    if isinstance(val, str):
        for fmt in ['%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d']:
            try: return datetime.strptime(val, fmt)
            except: pass
    return None

def business_hours(d1, d2):
    if not d1 or not d2: return None
    if d2 < d1: d1, d2 = d2, d1
    total = 0
    current = d1
    while current < d2:
        if current.weekday() < 5:
            start = current.replace(hour=9, minute=0, second=0, microsecond=0)
            end   = current.replace(hour=18, minute=0, second=0, microsecond=0)
            if current < start: current = start
            day_end = min(end, d2)
            if current < day_end:
                total += (day_end - current).total_seconds() / 3600
        current = (current + timedelta(days=1)).replace(hour=9, minute=0, second=0, microsecond=0)
    return total

def limite_ofi(razon, via):
    return 48 if razon == FSM and via == 'MARITIMO' else 24

def color_kpi(pct):
    if pct >= 95: return COLORS['accent']
    if pct >= 80: return COLORS['orange']
    return COLORS['red']

def metric_html(value, label, sub=None, color='accent'):
    color_class = '' if color == 'accent' else color
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ''
    return f"""
    <div class="metric-card">
        <div class="metric-value {color_class}">{value}</div>
        <div class="metric-label">{label}</div>
        {sub_html}
    </div>
    """

# ─────────────────────────────────────────────
# PROCESAMIENTO
# ─────────────────────────────────────────────
def procesar_liberadas(df):
    results = []
    for _, r in df.iterrows():
        razon = r.get('Razon Social', '')
        via   = str(r.get('Via', '')).upper().strip()
        canal = str(r.get('Canal', '')).upper().strip()
        f_ofi = parse_date(r.get('Fecha Oficialización'))
        f_can = parse_date(r.get('Fecha Cancelada'))
        hs    = business_hours(f_ofi, f_can)
        limite = LIMITES_LIB.get(via, {}).get(canal, 9999)
        desvio = hs is not None and hs > limite
        results.append({
            'razon': razon, 'nombre': 'FASA' if razon == FASA else 'FSM',
            'ref': r.get('Referencia',''), 'carpeta': r.get('Carpeta',''),
            'via': via, 'canal': canal,
            'f_ofi': f_ofi, 'f_cancel': f_can,
            'hs': round(hs, 1) if hs else None,
            'limite': limite, 'desvio': desvio,
            'desvio_desc': '', 'parametro': ''
        })
    return results

def procesar_oficializados(df):
    results = []
    for _, r in df.iterrows():
        razon = r.get('Razon Social', '')
        via   = str(r.get('Via', '')).upper().strip()
        f_ofi = parse_date(r.get('Fecha Oficialización'))
        f_ult = parse_date(r.get('Ultimo Evento'))
        hs    = business_hours(f_ofi, f_ult)
        limite = limite_ofi(razon, via)
        desvio = hs is not None and hs > limite
        results.append({
            'razon': razon, 'nombre': 'FASA' if razon == FASA else 'FSM',
            'ref': r.get('Referencia',''), 'carpeta': r.get('Carpeta',''),
            'via': via,
            'f_ofi': f_ofi, 'f_ult': f_ult,
            'hs': round(hs, 1) if hs else None,
            'limite': limite, 'desvio': desvio,
            'desvio_desc': '', 'parametro': ''
        })
    return results

def procesar_cm_presentados(df):
    results = []
    for _, r in df.iterrows():
        f_tad = parse_date(r.get('TAD SUBIDO'))
        f_ult = parse_date(r.get('Ult evento'))
        hs    = business_hours(f_tad, f_ult)
        desvio = hs is not None and hs > 48
        results.append({
            'carpeta': r.get('CARPETA',''), 'exp': r.get('Expediente',''),
            'f_tad': f_tad, 'f_ult': f_ult,
            'hs': round(hs, 1) if hs else None,
            'desvio': desvio, 'desvio_desc': '', 'parametro': ''
        })
    return results

def procesar_cm_aprobados(df):
    results = []
    for _, r in df.iterrows():
        f1 = parse_date(r.get('Fecha'))
        f2 = parse_date(r.get('Fechadeaprobacion'))
        dias = (f2 - f1).days if f1 and f2 else None
        rango = None
        if dias is not None:
            rango = '0 a 7' if dias <= 7 else ('8 a 15' if dias <= 15 else '+15')
        results.append({
            'carpeta': r.get('CARPETA',''), 'exp': r.get('Expediente',''),
            'f_inicio': f1, 'f_apro': f2,
            'dias': dias, 'rango': rango
        })
    return results

def calcular_kpi(items, con_parametros=False):
    total = len(items)
    if total == 0: return 0, 0, 0
    if con_parametros:
        out = sum(1 for i in items if i['desvio'] and str(i.get('parametro','')).upper() == 'INTERLOG')
    else:
        out = sum(1 for i in items if i['desvio'])
    in_count = total - out
    return round(in_count / total * 100, 2), in_count, out

# ─────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────
def chart_gauge(pct, title):
    color = color_kpi(pct)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=pct,
        number={'suffix': '%', 'font': {'size': 36, 'color': color, 'family': 'Barlow Condensed'}},
        title={'text': title, 'font': {'size': 13, 'color': '#9AB0C4', 'family': 'Barlow'}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': '#6B8099',
                     'tickfont': {'size': 10, 'color': '#6B8099'}},
            'bar': {'color': color, 'thickness': 0.25},
            'bgcolor': '#1A2E48',
            'borderwidth': 0,
            'steps': [
                {'range': [0, 80],  'color': 'rgba(255,61,94,0.15)'},
                {'range': [80, 95], 'color': 'rgba(255,140,66,0.15)'},
                {'range': [95, 100],'color': 'rgba(0,201,167,0.15)'},
            ],
            'threshold': {'line': {'color': '#FFD060', 'width': 3}, 'thickness': 0.8, 'value': 95}
        }
    ))
    fig.update_layout(**CHART_LAYOUT, height=220)
    return fig

def chart_hbar(labels, values, kpi_pcts, colors=None):
    colors = colors or [COLORS['accent']] * len(labels)
    fig = go.Figure()
    for i, (lbl, val, pct, col) in enumerate(zip(labels, values, kpi_pcts, colors)):
        fig.add_trace(go.Bar(
            y=[lbl], x=[val], orientation='h',
            marker_color=col, marker_line_width=0,
            text=f"  {val}  ({pct:.1f}%)", textposition='outside',
            textfont=dict(color='#F0F4F8', size=11, family='Barlow'),
            name=lbl, showlegend=False
        ))
    fig.update_layout(
        **CHART_LAYOUT, height=max(120, len(labels) * 55 + 40),
        barmode='overlay',
        xaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
        yaxis=dict(tickfont=dict(size=11, family='Barlow', color='#9AB0C4')),
    )
    return fig

def chart_donut(labels, values, colors, title=''):
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.6, marker_colors=colors,
        textinfo='none',
        hovertemplate='%{label}: %{value} (%{percent})<extra></extra>'
    ))
    total = sum(values)
    fig.add_annotation(
        text=f"<b>{total}</b>", x=0.5, y=0.55,
        font=dict(size=28, color='#F0F4F8', family='Barlow Condensed'),
        showarrow=False
    )
    fig.add_annotation(
        text="ops", x=0.5, y=0.35,
        font=dict(size=12, color='#6B8099', family='Barlow'),
        showarrow=False
    )
    fig.update_layout(**CHART_LAYOUT, height=260, showlegend=True,
        legend=dict(orientation='h', yanchor='bottom', y=-0.15,
                    font=dict(size=10, color='#9AB0C4')))
    return fig

def chart_scatter_tiempo(items, limite, title=''):
    hs_vals = [i['hs'] for i in items if i['hs'] is not None]
    if not hs_vals: return None
    colors = [COLORS['red'] if h > limite else COLORS['accent'] for h in hs_vals]
    fig = go.Figure()
    fig.add_hline(y=limite, line_dash='dash', line_color=COLORS['orange'],
                  annotation_text=f'Límite {limite}hs', annotation_position='top right',
                  annotation_font=dict(color=COLORS['orange'], size=11))
    fig.add_trace(go.Scatter(
        x=list(range(len(hs_vals))), y=hs_vals,
        mode='markers',
        marker=dict(color=colors, size=10, line=dict(color='white', width=1.5)),
        hovertemplate='Op %{x}: %{y:.1f} hs<extra></extra>'
    ))
    fig.update_layout(**CHART_LAYOUT, height=220,
        xaxis=dict(showticklabels=False, title='Operaciones'),
        yaxis=dict(title='Horas hábiles'))
    return fig

def chart_stacked_canales(nombre, items):
    vias = ['AVION', 'MARITIMO', 'CAMION']
    verde_v, naranja_v, rojo_v, labels = [], [], [], []
    for via in vias:
        via_items = [i for i in items if i['via'] == via]
        if not via_items: continue
        by_canal = Counter(i['canal'] for i in via_items)
        verde_v.append(by_canal.get('VERDE', 0))
        naranja_v.append(by_canal.get('NARANJA', 0))
        rojo_v.append(by_canal.get('ROJO', 0))
        labels.append(f"{via} (n={len(via_items)})")

    if not labels: return None
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Verde',   x=labels, y=verde_v,   marker_color=COLORS['verde'],   text=verde_v,   textposition='inside', textfont=dict(color='white', size=11)))
    fig.add_trace(go.Bar(name='Naranja', x=labels, y=naranja_v, marker_color=COLORS['naranja'], text=naranja_v, textposition='inside', textfont=dict(color='white', size=11)))
    fig.add_trace(go.Bar(name='Rojo',    x=labels, y=rojo_v,    marker_color=COLORS['rojo'],    text=rojo_v,    textposition='inside', textfont=dict(color='white', size=11)))
    fig.update_layout(**CHART_LAYOUT, barmode='stack', height=280,
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    font=dict(color='#9AB0C4', size=11)),
        xaxis=dict(tickfont=dict(size=11, color='#9AB0C4')))
    return fig

# ─────────────────────────────────────────────
# GENERAR EXCEL DE DESVÍOS ESTILIZADO
# ─────────────────────────────────────────────
def generar_excel_desvios(lib_items, ofi_items, cm_pre_items, mes='MES'):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    DARK_BG = "0D1B2A"; MID_BG = "1B2B3E"; CARD = "132236"
    ACCENT  = "008B74"; GOLD = "FFD060";  ROJO = "FF3D5E"
    WHITE   = "FFFFFF"; LGRAY = "9AB0C4"; PURPLE = "5B21B6"

    def hfill(c): return PatternFill("solid", fgColor=c)
    def bw(s=10):  return Font(bold=True, color=WHITE, size=s, name="Calibri")
    def nw(s=10):  return Font(color=WHITE, size=s, name="Calibri")
    def nr(s=10):  return Font(color=ROJO, size=s, name="Calibri", bold=True)
    def ng(s=10):  return Font(color=GOLD, size=s, name="Calibri", bold=True)
    def cen():     return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def lft():     return Alignment(horizontal="left",   vertical="center", wrap_text=True)
    def brd():
        s = Side(border_style="thin", color="1B2B3E")
        return Border(left=s, right=s, top=s, bottom=s)

    wb = Workbook()

    def make_sheet(ws, title_text, headers, rows_data, header_color, col_widths,
                   edit_cols, ref_col_idx=None):
        # Título
        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        ws["A1"] = title_text
        ws["A1"].font = bw(13); ws["A1"].fill = hfill(DARK_BG)
        ws["A1"].alignment = cen(); ws.row_dimensions[1].height = 30

        # Subtítulo
        ws.merge_cells(f"A2:{get_column_letter(len(headers))}2")
        ws["A2"] = "⚠️  Completar columnas DESVÍO y PARÁMETRO — Si Parámetro ≠ INTERLOG la operación se considera IN"
        ws["A2"].font = Font(color=GOLD, size=9, name="Calibri")
        ws["A2"].fill = hfill(MID_BG); ws["A2"].alignment = cen()
        ws.row_dimensions[2].height = 20

        # Headers
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(3, ci, h)
            cell.font = bw(10); cell.fill = hfill(header_color)
            cell.alignment = cen(); cell.border = brd()
        ws.row_dimensions[3].height = 25

        # Datos
        if not rows_data:
            ws.merge_cells(f"A4:{get_column_letter(len(headers))}4")
            ws["A4"] = "✅  Sin desvíos detectados — KPI 100%"
            ws["A4"].font = Font(color="00C9A7", size=11, bold=True, name="Calibri")
            ws["A4"].fill = hfill(MID_BG); ws["A4"].alignment = cen()
            ws.row_dimensions[4].height = 30
        else:
            for ri, row_vals in enumerate(rows_data, 4):
                fill_c = CARD if ri % 2 == 0 else MID_BG
                for ci, val in enumerate(row_vals, 1):
                    cell = ws.cell(ri, ci, val)
                    cell.border = brd()
                    cell.alignment = cen() if ci > 1 else lft()
                    if ci in edit_cols:
                        cell.fill = hfill("1A3A2A"); cell.font = ng(10)
                    elif ci == ref_col_idx:
                        cell.fill = hfill(fill_c); cell.font = bw(10)
                    else:
                        cell.fill = hfill(fill_c)
                        # Hs en rojo
                        if headers[ci-1] in ['Hs Transcurridas']:
                            cell.font = nr(10)
                        else:
                            cell.font = nw(10)
                ws.row_dimensions[ri].height = 20

        # Anchos
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A4"

    # ── HOJA 1: LIBERADAS ──
    ws1 = wb.active; ws1.title = "LIBERADAS - DESVÍOS"
    lib_desvios = [i for i in lib_items if i['desvio']]
    rows_lib = [[
        'FASA' if i['razon'] == FASA else 'FSM',
        i['ref'], str(i['carpeta']), i['via'], i['canal'],
        i['f_ofi'].strftime('%d/%m/%Y') if i['f_ofi'] else '',
        i['f_cancel'].strftime('%d/%m/%Y %H:%M') if i['f_cancel'] else '',
        i['hs'], i['limite'], '', ''
    ] for i in lib_desvios]
    make_sheet(ws1,
        f"LIBERADAS {mes} — OPERACIONES CON DESVÍO",
        ["Razón Social","Referencia","Carpeta","Vía","Canal",
         "F. Oficialización","F. Cancelada","Hs Transcurridas","Límite (hs)","DESVÍO ✏️","PARÁMETRO ✏️"],
        rows_lib, ACCENT, [28,16,12,12,10,18,22,18,12,35,25],
        edit_cols=[10,11], ref_col_idx=2
    )

    # ── HOJA 2: OFICIALIZADOS ──
    ws2 = wb.create_sheet("OFICIALIZADOS - DESVÍOS")
    ofi_desvios = [i for i in ofi_items if i['desvio']]
    rows_ofi = [[
        'FASA' if i['razon'] == FASA else 'FSM',
        i['ref'], str(i['carpeta']), i['via'],
        i['f_ofi'].strftime('%d/%m/%Y') if i['f_ofi'] else '',
        i['f_ult'].strftime('%d/%m/%Y %H:%M') if i['f_ult'] else '',
        i['hs'], i['limite'], '', ''
    ] for i in ofi_desvios]
    make_sheet(ws2,
        f"OFICIALIZADOS {mes} — OPERACIONES CON DESVÍO",
        ["Razón Social","Referencia","Carpeta","Vía",
         "F. Oficialización","Último Evento","Hs Transcurridas","Límite (hs)","DESVÍO ✏️","PARÁMETRO ✏️"],
        rows_ofi, "005F52", [28,16,12,12,18,22,18,12,35,25],
        edit_cols=[9,10], ref_col_idx=2
    )

    # ── HOJA 3: CM PRESENTADOS ──
    ws3 = wb.create_sheet("CM PRESENTADOS - DESVÍOS")
    cm_desvios = [i for i in cm_pre_items if i['desvio']]
    rows_cm = [[
        str(i['carpeta']), str(i['exp']),
        i['f_tad'].strftime('%d/%m/%Y') if i['f_tad'] else '',
        i['f_ult'].strftime('%d/%m/%Y') if i['f_ult'] else '',
        i['hs'], 48, '', ''
    ] for i in cm_desvios]
    make_sheet(ws3,
        f"CM PRESENTADOS {mes} — OPERACIONES CON DESVÍO",
        ["Carpeta","Expediente","TAD Subido","Último Evento",
         "Hs Transcurridas","Límite (hs)","DESVÍO ✏️","PARÁMETRO ✏️"],
        rows_cm, PURPLE, [12,28,18,18,18,12,35,25],
        edit_cols=[7,8], ref_col_idx=2
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# EXPORT EXCEL DESVÍOS (simple, para historial)
# ─────────────────────────────────────────────
def export_excel_desvios(lib_items, ofi_items, cm_pre_items):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Liberadas
        lib_desvios = [i for i in lib_items if i['desvio']]
        if lib_desvios:
            df = pd.DataFrame([{
                'Razón Social': i['nombre'], 'Referencia': i['ref'],
                'Vía': i['via'], 'Canal': i['canal'],
                'Hs Transcurridas': i['hs'], 'Límite (hs)': i['limite'],
                'DESVÍO': i.get('desvio_desc',''), 'PARÁMETRO': i.get('parametro','')
            } for i in lib_desvios])
            df.to_excel(writer, sheet_name='LIBERADAS - DESVÍOS', index=False)

        # Oficializados
        ofi_desvios = [i for i in ofi_items if i['desvio']]
        if ofi_desvios:
            df = pd.DataFrame([{
                'Razón Social': i['nombre'], 'Referencia': i['ref'],
                'Vía': i['via'], 'Hs Transcurridas': i['hs'], 'Límite (hs)': i['limite'],
                'DESVÍO': i.get('desvio_desc',''), 'PARÁMETRO': i.get('parametro','')
            } for i in ofi_desvios])
            df.to_excel(writer, sheet_name='OFICIALIZADOS - DESVÍOS', index=False)

        # CM Presentados
        cm_desvios = [i for i in cm_pre_items if i['desvio']]
        if cm_desvios:
            df = pd.DataFrame([{
                'Carpeta': i['carpeta'], 'Expediente': i['exp'],
                'Hs Transcurridas': i['hs'], 'Límite (hs)': 48,
                'DESVÍO': i.get('desvio_desc',''), 'PARÁMETRO': i.get('parametro','')
            } for i in cm_desvios])
            df.to_excel(writer, sheet_name='CM PRESENTADOS - DESVÍOS', index=False)

    output.seek(0)
    return output

# ─────────────────────────────────────────────
# SECCIÓN KPI
# ─────────────────────────────────────────────
def render_kpi_section(nombre, lib_items, ofi_items, con_param=False):
    lib_r = [i for i in lib_items if i['nombre'] == nombre]
    ofi_r = [i for i in ofi_items if i['nombre'] == nombre]

    pct_lib, in_lib, out_lib = calcular_kpi(lib_r, con_param)
    pct_ofi, in_ofi, out_ofi = calcular_kpi(ofi_r, con_param)

    st.markdown(f'<div class="section-header">MODAL MULTIMODAL · {nombre}</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**OFICIALIZACIÓN**")
        c1, c2, c3 = st.columns(3)
        c1.plotly_chart(chart_gauge(pct_ofi, "KPI IN"), use_container_width=True)

        params_ofi = Counter(i['parametro'] or 'S/DESVIO' for i in ofi_r if not i['desvio'])
        params_ofi.update({i['parametro'] or 'PENDIENTE': 1 for i in ofi_r if i['desvio']})
        lbls = list(params_ofi.keys())
        vals = list(params_ofi.values())
        cols_bar = [COLORS['accent'] if l == 'S/DESVIO' else (COLORS['red'] if l == 'INTERLOG' else COLORS['orange']) for l in lbls]
        kpi_p = [v/len(ofi_r)*100 for v in vals]

        with c2:
            st.plotly_chart(chart_hbar(lbls, vals, kpi_p, cols_bar), use_container_width=True)
        with c3:
            st.markdown(metric_html(str(len(ofi_r)), "Total ops", None, 'accent'), unsafe_allow_html=True)
            st.markdown(metric_html(str(in_ofi), "IN", None, 'accent'), unsafe_allow_html=True)
            st.markdown(metric_html(str(out_ofi), "OUT", None, 'red' if out_ofi else 'accent'), unsafe_allow_html=True)

    with col2:
        st.markdown("**LIBERACIÓN**")
        c1, c2, c3 = st.columns(3)
        kpi_col = 'accent' if pct_lib >= 95 else ('orange' if pct_lib >= 80 else 'red')
        c1.plotly_chart(chart_gauge(pct_lib, "KPI IN"), use_container_width=True)

        params_lib = Counter(i['parametro'] or 'S/DESVIO' for i in lib_r if not i['desvio'])
        params_lib.update({i['parametro'] or 'PENDIENTE': 1 for i in lib_r if i['desvio']})
        lbls = list(params_lib.keys())
        vals = list(params_lib.values())
        cols_bar = [COLORS['accent'] if l == 'S/DESVIO' else (COLORS['red'] if l == 'INTERLOG' else COLORS['orange']) for l in lbls]
        kpi_p = [v/len(lib_r)*100 for v in vals]

        with c2:
            st.plotly_chart(chart_hbar(lbls, vals, kpi_p, cols_bar), use_container_width=True)
        with c3:
            st.markdown(metric_html(str(len(lib_r)), "Total ops", None, 'accent'), unsafe_allow_html=True)
            st.markdown(metric_html(str(in_lib), "IN", None, 'accent'), unsafe_allow_html=True)
            st.markdown(metric_html(str(out_lib), "OUT", None, 'red' if out_lib else 'accent'), unsafe_allow_html=True)

    # Canal Verde Avión
    va = [i for i in lib_r if i['via'] == 'AVION' and i['canal'] == 'VERDE']
    if va:
        st.markdown(f'<div class="section-header" style="border-color:#00C9A7">CANAL VERDE · VÍA AÉREA · {nombre}</div>', unsafe_allow_html=True)
        pct_va, in_va, out_va = calcular_kpi(va, con_param)
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(metric_html(f"{pct_va:.1f}%", "KPI IN", "Target: 95%", 'accent' if pct_va >= 95 else 'orange'), unsafe_allow_html=True)
        c2.markdown(metric_html(str(len(va)), "Total ops", "Canal Verde Avión", 'accent'), unsafe_allow_html=True)
        c3.markdown(metric_html(str(in_va), "IN", "Dentro del rango", 'accent'), unsafe_allow_html=True)
        c4.markdown(metric_html(str(out_va), "OUT", "Fuera del rango", 'red' if out_va else 'accent'), unsafe_allow_html=True)

        fig_scatter = chart_scatter_tiempo(va, 24)
        if fig_scatter:
            st.plotly_chart(fig_scatter, use_container_width=True)

    # Canales
    st.markdown(f'<div class="section-header" style="border-color:#FFD060">DISTRIBUCIÓN DE CANALES · {nombre}</div>', unsafe_allow_html=True)
    fig_canal = chart_stacked_canales(nombre, lib_r)
    if fig_canal:
        st.plotly_chart(fig_canal, use_container_width=True)

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────

# Header
st.markdown("""
<div style="background: linear-gradient(135deg, #0F2040 0%, #132236 100%);
     border-bottom: 1px solid rgba(0,201,167,0.2);
     padding: 1.5rem 2rem; margin: -1rem -1rem 1.5rem; border-radius: 0 0 12px 12px;">
  <div style="display:flex; align-items:center; gap:1rem;">
    <div>
      <div style="font-family:'Barlow Condensed',sans-serif; font-size:2rem;
           font-weight:800; color:#00C9A7; letter-spacing:3px;">INTERLOG</div>
      <div style="font-family:'Barlow Condensed',sans-serif; font-size:1rem;
           color:#6B8099; letter-spacing:2px; text-transform:uppercase;">KPI Dashboard · Comercio Exterior</div>
    </div>
    <div style="margin-left:auto; text-align:right;">
      <div style="font-size:0.75rem; color:#6B8099; text-transform:uppercase; letter-spacing:1px;">Reporte Mensual</div>
      <div style="font-family:'Barlow Condensed',sans-serif; font-size:1.2rem;
           color:#F0F4F8; font-weight:600;">FASA / FSM</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── ESTADO DE SESIÓN ──
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'lib_items' not in st.session_state:
    st.session_state.lib_items = []
if 'ofi_items' not in st.session_state:
    st.session_state.ofi_items = []
if 'cm_pre_items' not in st.session_state:
    st.session_state.cm_pre_items = []
if 'cm_apr_items' not in st.session_state:
    st.session_state.cm_apr_items = []
if 'mes' not in st.session_state:
    st.session_state.mes = ''

# ── STEP INDICATOR ──
steps = ["📁 Cargar archivos", "⚠️ Revisar desvíos", "📊 Dashboard", "📥 Exportar"]
cols_steps = st.columns(4)
for i, (col, label) in enumerate(zip(cols_steps, steps)):
    active = st.session_state.step == i+1
    done   = st.session_state.step > i+1
    bg = '#00C9A7' if active else ('#007A65' if done else '#132236')
    tc = '#0A1628' if active else ('#F0F4F8' if done else '#6B8099')
    col.markdown(f"""
    <div style="background:{bg}; color:{tc}; border-radius:6px; padding:0.5rem;
         text-align:center; font-family:'Barlow Condensed',sans-serif;
         font-size:0.85rem; font-weight:700; letter-spacing:0.5px;">
        {label}
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# STEP 1 — CARGAR ARCHIVOS
# ══════════════════════════════════════════════
if st.session_state.step == 1:
    st.markdown('<div class="section-header">PASO 1 · CARGAR ARCHIVOS EXCEL</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="alert-info">
    Subí los 4 archivos Excel del mes. El sistema procesará automáticamente todas las métricas.
    </div><br>
    """, unsafe_allow_html=True)

    mes = st.text_input("📅 Mes del reporte (ej: DICIEMBRE 2025)", value=st.session_state.mes)

    col1, col2 = st.columns(2)
    with col1:
        f_lib = st.file_uploader("📦 LIBERADAS", type=['xlsx'], key='lib')
        f_ofi = st.file_uploader("📋 OFICIALIZADOS", type=['xlsx'], key='ofi')
    with col2:
        f_cm_pre = st.file_uploader("📜 CM PRESENTADOS", type=['xlsx'], key='cmpre')
        f_cm_apr = st.file_uploader("✅ CM APROBADOS", type=['xlsx'], key='cmapr')

    archivos_ok = all([f_lib, f_ofi, f_cm_pre, f_cm_apr])

    if archivos_ok:
        st.markdown('<div class="alert-success">✅ Los 4 archivos cargados correctamente</div><br>', unsafe_allow_html=True)

    if st.button("▶  PROCESAR Y CONTINUAR", disabled=not archivos_ok):
        with st.spinner("Procesando datos..."):
            df_lib     = pd.read_excel(f_lib)
            df_ofi     = pd.read_excel(f_ofi)
            df_cm_pre  = pd.read_excel(f_cm_pre)
            df_cm_apr  = pd.read_excel(f_cm_apr)

            st.session_state.lib_items     = procesar_liberadas(df_lib)
            st.session_state.ofi_items     = procesar_oficializados(df_ofi)
            st.session_state.cm_pre_items  = procesar_cm_presentados(df_cm_pre)
            st.session_state.cm_apr_items  = procesar_cm_aprobados(df_cm_apr)
            st.session_state.mes           = mes
            st.session_state.step          = 2
            st.rerun()

# ══════════════════════════════════════════════
# STEP 2 — REVISAR DESVÍOS
# ══════════════════════════════════════════════
elif st.session_state.step == 2:
    lib_items    = st.session_state.lib_items
    ofi_items    = st.session_state.ofi_items
    cm_pre_items = st.session_state.cm_pre_items

    desvios_lib = [i for i in lib_items if i['desvio']]
    desvios_ofi = [i for i in ofi_items if i['desvio']]
    desvios_cm  = [i for i in cm_pre_items if i['desvio']]
    total_desvios = len(desvios_lib) + len(desvios_ofi) + len(desvios_cm)

    st.markdown('<div class="section-header">PASO 2 · INSTANCIA DE REVISIÓN DE DESVÍOS</div>', unsafe_allow_html=True)

    if total_desvios == 0:
        st.markdown('<div class="alert-success">✅ Sin desvíos detectados en ningún proceso. KPI 100% en todo.</div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("▶  IR AL DASHBOARD"):
            st.session_state.step = 3
            st.rerun()
    else:
        # ── SUB-ESTADO: ¿ya descargó o está subiendo? ──
        if 'desvio_sub_step' not in st.session_state:
            st.session_state.desvio_sub_step = 'descargar'

        # ──────────────────────────────────────────
        # SUB-STEP A: DESCARGAR EXCEL
        # ──────────────────────────────────────────
        if st.session_state.desvio_sub_step == 'descargar':
            st.markdown(f"""
            <div class="alert-warn">
            ⚠️  Se detectaron <b>{total_desvios} operaciones fuera del rango</b>.
            Descargá el Excel, completá las columnas <b>DESVÍO</b> y <b>PARÁMETRO</b>, y volvé a subirlo.<br><br>
            <small>💡 Si el Parámetro <b>no es INTERLOG</b> → la operación se considera <b>IN</b> automáticamente.</small>
            </div><br>
            """, unsafe_allow_html=True)

            # Generar Excel estilizado
            excel_buf = generar_excel_desvios(lib_items, ofi_items, cm_pre_items,
                                               st.session_state.mes or 'MES')

            c1, c2, c3 = st.columns([2, 2, 3])
            with c1:
                st.download_button(
                    label="⬇  DESCARGAR EXCEL DE DESVÍOS",
                    data=excel_buf,
                    file_name=f"DESVIOS_{(st.session_state.mes or 'MES').replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with c2:
                if st.button("✅  YA LO COMPLETÉ — SUBIR AHORA", use_container_width=True):
                    st.session_state.desvio_sub_step = 'subir'
                    st.rerun()
            with c3:
                if st.button("◀  VOLVER A CARGAR ARCHIVOS", use_container_width=True):
                    st.session_state.step = 1
                    st.session_state.desvio_sub_step = 'descargar'
                    st.rerun()

        # ──────────────────────────────────────────
        # SUB-STEP B: SUBIR EXCEL COMPLETADO
        # ──────────────────────────────────────────
        elif st.session_state.desvio_sub_step == 'subir':
            st.markdown("""
            <div class="alert-info">
            📂 Subí el Excel de desvíos que ya completaste. El sistema leerá los campos
            <b>DESVÍO</b> y <b>PARÁMETRO</b> de cada hoja y recalculará los KPIs.
            </div><br>
            """, unsafe_allow_html=True)

            f_desvios = st.file_uploader(
                "📋 Excel de Desvíos completado",
                type=['xlsx'], key='desvios_upload'
            )

            if f_desvios:
                try:
                    wb_dev = pd.ExcelFile(f_desvios)
                    errores = []

                    # ── Leer LIBERADAS ──
                    if 'LIBERADAS - DESVÍOS' in wb_dev.sheet_names:
                        df_dev_lib = wb_dev.parse('LIBERADAS - DESVÍOS', skiprows=2)
                        # Normalizar columnas
                        df_dev_lib.columns = [str(c).strip() for c in df_dev_lib.columns]
                        col_desv  = next((c for c in df_dev_lib.columns if 'DESVÍO' in c or 'DESVIO' in c.upper()), None)
                        col_param = next((c for c in df_dev_lib.columns if 'PARÁMETRO' in c or 'PARAMETRO' in c.upper()), None)
                        col_ref   = next((c for c in df_dev_lib.columns if 'REFERENCIA' in c.upper() or 'REF' in c.upper()), None)

                        if col_ref and col_param:
                            ref_map = {}
                            for _, row in df_dev_lib.iterrows():
                                ref = str(row.get(col_ref, '')).strip()
                                if ref:
                                    ref_map[ref] = {
                                        'desvio_desc': str(row.get(col_desv, '') or '').strip(),
                                        'parametro':   str(row.get(col_param,'') or '').strip()
                                    }
                            for item in lib_items:
                                if item['ref'] in ref_map:
                                    item['desvio_desc'] = ref_map[item['ref']]['desvio_desc']
                                    item['parametro']   = ref_map[item['ref']]['parametro']

                    # ── Leer OFICIALIZADOS ──
                    if 'OFICIALIZADOS - DESVÍOS' in wb_dev.sheet_names:
                        df_dev_ofi = wb_dev.parse('OFICIALIZADOS - DESVÍOS', skiprows=2)
                        df_dev_ofi.columns = [str(c).strip() for c in df_dev_ofi.columns]
                        col_desv  = next((c for c in df_dev_ofi.columns if 'DESVÍO' in c or 'DESVIO' in c.upper()), None)
                        col_param = next((c for c in df_dev_ofi.columns if 'PARÁMETRO' in c or 'PARAMETRO' in c.upper()), None)
                        col_ref   = next((c for c in df_dev_ofi.columns if 'REFERENCIA' in c.upper() or 'REF' in c.upper()), None)

                        if col_ref and col_param:
                            ref_map = {}
                            for _, row in df_dev_ofi.iterrows():
                                ref = str(row.get(col_ref, '')).strip()
                                if ref:
                                    ref_map[ref] = {
                                        'desvio_desc': str(row.get(col_desv, '') or '').strip(),
                                        'parametro':   str(row.get(col_param,'') or '').strip()
                                    }
                            for item in ofi_items:
                                if item['ref'] in ref_map:
                                    item['desvio_desc'] = ref_map[item['ref']]['desvio_desc']
                                    item['parametro']   = ref_map[item['ref']]['parametro']

                    # ── Leer CM PRESENTADOS ──
                    if 'CM PRESENTADOS - DESVÍOS' in wb_dev.sheet_names:
                        df_dev_cm = wb_dev.parse('CM PRESENTADOS - DESVÍOS', skiprows=2)
                        df_dev_cm.columns = [str(c).strip() for c in df_dev_cm.columns]
                        col_desv  = next((c for c in df_dev_cm.columns if 'DESVÍO' in c or 'DESVIO' in c.upper()), None)
                        col_param = next((c for c in df_dev_cm.columns if 'PARÁMETRO' in c or 'PARAMETRO' in c.upper()), None)
                        col_exp   = next((c for c in df_dev_cm.columns if 'EXPEDIENTE' in c.upper() or 'EXP' in c.upper()), None)

                        if col_exp and col_param:
                            exp_map = {}
                            for _, row in df_dev_cm.iterrows():
                                exp = str(row.get(col_exp, '')).strip()
                                if exp:
                                    exp_map[exp] = {
                                        'desvio_desc': str(row.get(col_desv, '') or '').strip(),
                                        'parametro':   str(row.get(col_param,'') or '').strip()
                                    }
                            for item in cm_pre_items:
                                if str(item['exp']).strip() in exp_map:
                                    item['desvio_desc'] = exp_map[str(item['exp']).strip()]['desvio_desc']
                                    item['parametro']   = exp_map[str(item['exp']).strip()]['parametro']

                    # ── Verificar completitud ──
                    pendientes_lib = [i for i in lib_items  if i['desvio'] and not i.get('parametro','').strip()]
                    pendientes_ofi = [i for i in ofi_items  if i['desvio'] and not i.get('parametro','').strip()]
                    pendientes_cm  = [i for i in cm_pre_items if i['desvio'] and not i.get('parametro','').strip()]
                    total_pend = len(pendientes_lib) + len(pendientes_ofi) + len(pendientes_cm)

                    if total_pend == 0:
                        st.markdown('<div class="alert-success">✅ Todos los desvíos fueron completados correctamente.</div>', unsafe_allow_html=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                        # Preview tabla
                        todos = (
                            [{'Proceso':'LIBERADAS', 'Ref': i['ref'], 'Razón':i['nombre'],
                              'Vía':i['via'], 'Canal':i.get('canal',''), 'Hs':i['hs'],
                              'Límite':i['limite'], 'Desvío':i['desvio_desc'], 'Parámetro':i['parametro']}
                             for i in lib_items if i['desvio']] +
                            [{'Proceso':'OFICIALIZADOS', 'Ref': i['ref'], 'Razón':i['nombre'],
                              'Vía':i['via'], 'Canal':'', 'Hs':i['hs'],
                              'Límite':i['limite'], 'Desvío':i['desvio_desc'], 'Parámetro':i['parametro']}
                             for i in ofi_items if i['desvio']] +
                            [{'Proceso':'CM PRES.', 'Ref': i['exp'], 'Razón':'',
                              'Vía':'', 'Canal':'', 'Hs':i['hs'],
                              'Límite':48, 'Desvío':i['desvio_desc'], 'Parámetro':i['parametro']}
                             for i in cm_pre_items if i['desvio']]
                        )
                        df_prev = pd.DataFrame(todos)
                        # Color por parámetro
                        def color_param(val):
                            if str(val).upper() == 'INTERLOG':
                                return 'background-color: rgba(255,61,94,0.2); color: #FF3D5E; font-weight:bold'
                            elif val:
                                return 'background-color: rgba(0,201,167,0.15); color: #00C9A7; font-weight:bold'
                            return ''
                        st.dataframe(
                            df_prev.style.applymap(color_param, subset=['Parámetro']),
                            use_container_width=True, hide_index=True
                        )

                        c1, c2 = st.columns([1, 3])
                        with c2:
                            if st.button("▶  GENERAR DASHBOARD", use_container_width=True):
                                st.session_state.lib_items    = lib_items
                                st.session_state.ofi_items    = ofi_items
                                st.session_state.cm_pre_items = cm_pre_items
                                st.session_state.desvio_sub_step = 'descargar'
                                st.session_state.step = 3
                                st.rerun()
                    else:
                        st.markdown(f'<div class="alert-warn">⚠️ Faltan completar {total_pend} desvíos. Revisá el Excel y volvé a subirlo.</div>', unsafe_allow_html=True)
                        if pendientes_lib:
                            st.markdown("**Sin completar en LIBERADAS:**")
                            for i in pendientes_lib:
                                st.markdown(f"- `{i['ref']}` · {i['nombre']} · {i['via']} {i.get('canal','')}")
                        if pendientes_ofi:
                            st.markdown("**Sin completar en OFICIALIZADOS:**")
                            for i in pendientes_ofi:
                                st.markdown(f"- `{i['ref']}` · {i['nombre']}")
                        if pendientes_cm:
                            st.markdown("**Sin completar en CM PRESENTADOS:**")
                            for i in pendientes_cm:
                                st.markdown(f"- `{i['exp']}`")

                except Exception as e:
                    st.markdown(f'<div class="alert-warn">❌ Error al leer el Excel: {str(e)}</div>', unsafe_allow_html=True)

            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button("◀  VOLVER"):
                    st.session_state.desvio_sub_step = 'descargar'
                    st.rerun()

# ══════════════════════════════════════════════
# STEP 3 — DASHBOARD
# ══════════════════════════════════════════════
elif st.session_state.step == 3:
    lib_items    = st.session_state.lib_items
    ofi_items    = st.session_state.ofi_items
    cm_pre_items = st.session_state.cm_pre_items
    cm_apr_items = st.session_state.cm_apr_items
    mes          = st.session_state.mes or 'MENSUAL'

    st.markdown(f'<div class="section-header">📊 DASHBOARD KPI · {mes}</div>', unsafe_allow_html=True)

    # ── RESUMEN EJECUTIVO ──
    pct_lib_fasa, _, _ = calcular_kpi([i for i in lib_items if i['nombre'] == 'FASA'], True)
    pct_lib_fsm,  _, _ = calcular_kpi([i for i in lib_items if i['nombre'] == 'FSM'],  True)
    pct_ofi_fasa, _, _ = calcular_kpi([i for i in ofi_items if i['nombre'] == 'FASA'], True)
    pct_ofi_fsm,  _, _ = calcular_kpi([i for i in ofi_items if i['nombre'] == 'FSM'],  True)
    pct_cm, _, _       = calcular_kpi(cm_pre_items, True)

    c1,c2,c3,c4,c5 = st.columns(5)
    kpis = [
        (f"{pct_lib_fasa:.0f}%", "LIB FASA", "Liberación", 'accent' if pct_lib_fasa>=95 else 'orange'),
        (f"{pct_lib_fsm:.0f}%",  "LIB FSM",  "Liberación", 'accent' if pct_lib_fsm>=95  else 'orange'),
        (f"{pct_ofi_fasa:.0f}%", "OFI FASA", "Oficialización", 'accent' if pct_ofi_fasa>=95 else 'orange'),
        (f"{pct_ofi_fsm:.0f}%",  "OFI FSM",  "Oficialización", 'accent' if pct_ofi_fsm>=95  else 'orange'),
        (f"{pct_cm:.0f}%",       "CM",        "Certificados Mineros", 'accent' if pct_cm>=95 else 'orange'),
    ]
    for col, (val, lbl, sub, color) in zip([c1,c2,c3,c4,c5], kpis):
        col.markdown(metric_html(val, lbl, sub, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── TABS ──
    tab1, tab2, tab3, tab4 = st.tabs(["📦 FASA", "🏔️ FSM", "🚢 MARÍTIMOS", "📜 CERTIFICADOS MINEROS"])

    with tab1:
        render_kpi_section('FASA', lib_items, ofi_items, con_param=True)

    with tab2:
        render_kpi_section('FSM', lib_items, ofi_items, con_param=True)

    with tab3:
        st.markdown('<div class="section-header">MODAL MARÍTIMOS · TODAS LAS RAZONES SOCIALES</div>', unsafe_allow_html=True)
        mar_lib = [i for i in lib_items if i['via'] == 'MARITIMO']
        mar_ofi = [i for i in ofi_items if i['via'] == 'MARITIMO']
        pct_mar_lib, in_ml, out_ml = calcular_kpi(mar_lib, True)
        pct_mar_ofi, in_mo, out_mo = calcular_kpi(mar_ofi, True)

        c1,c2,c3,c4 = st.columns(4)
        c1.markdown(metric_html(str(len(mar_lib)), "Liberaciones", "Vía Marítima"), unsafe_allow_html=True)
        c2.markdown(metric_html(str(len(mar_ofi)), "Oficializaciones", "Vía Marítima"), unsafe_allow_html=True)
        c3.plotly_chart(chart_gauge(pct_mar_lib, "KPI IN Lib"), use_container_width=True)
        c4.plotly_chart(chart_gauge(pct_mar_ofi, "KPI IN Ofi"), use_container_width=True)

        if mar_lib:
            by_canal = Counter(i['canal'] for i in mar_lib)
            fig = go.Figure(go.Pie(
                labels=list(by_canal.keys()), values=list(by_canal.values()),
                hole=0.55, marker_colors=[COLORS.get(k.lower(), COLORS['gray']) for k in by_canal.keys()]
            ))
            fig.update_layout(**CHART_LAYOUT, height=280, title="Distribución de canales marítimos")
            st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.markdown('<div class="section-header">KPI CERTIFICADOS MINEROS</div>', unsafe_allow_html=True)
        pct_cm_pre, in_cm, out_cm = calcular_kpi(cm_pre_items, True)

        c1,c2,c3,c4 = st.columns(4)
        c1.plotly_chart(chart_gauge(pct_cm_pre, "KPI Presentación"), use_container_width=True)
        c2.markdown(metric_html(str(len(cm_pre_items)), "Total expedientes", "Presentados"), unsafe_allow_html=True)
        c3.markdown(metric_html(str(in_cm), "IN", "Dentro del rango", 'accent'), unsafe_allow_html=True)
        c4.markdown(metric_html(str(out_cm), "OUT", "Fuera del rango", 'red' if out_cm else 'accent'), unsafe_allow_html=True)

        if cm_apr_items:
            st.markdown("**APROBADOS — Distribución por tiempo (solo informativo)**")
            rangos = Counter(i['rango'] for i in cm_apr_items if i['rango'])
            labels = ['0 a 7', '8 a 15', '+15']
            values = [rangos.get(l, 0) for l in labels]
            fig = go.Figure(go.Bar(
                x=labels, y=values,
                marker_color=[COLORS['verde'], COLORS['naranja'], COLORS['rojo']],
                text=values, textposition='outside',
                textfont=dict(color='#F0F4F8', size=12)
            ))
            fig.update_layout(**CHART_LAYOUT, height=280,
                xaxis=dict(title="Rango (días corridos)"),
                yaxis=dict(title="Cantidad de expedientes"))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 3])
    with c1:
        if st.button("◀  VOLVER"):
            st.session_state.step = 2
            st.rerun()
    with c2:
        if st.button("▶  EXPORTAR"):
            st.session_state.step = 4
            st.rerun()

# ══════════════════════════════════════════════
# STEP 4 — EXPORTAR
# ══════════════════════════════════════════════
elif st.session_state.step == 4:
    st.markdown('<div class="section-header">PASO 4 · EXPORTAR</div>', unsafe_allow_html=True)

    lib_items    = st.session_state.lib_items
    ofi_items    = st.session_state.ofi_items
    cm_pre_items = st.session_state.cm_pre_items
    mes          = st.session_state.mes or 'MES'

    st.markdown("""
    <div class="alert-info">
    Descargá el Excel con todos los desvíos justificados para tu registro,
    y el PowerPoint del informe completo listo para presentar.
    </div><br>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 📊 Excel de Desvíos")
        st.markdown("Contiene todas las operaciones con desvío, sus descripciones y parámetros.")
        excel_buf = export_excel_desvios(lib_items, ofi_items, cm_pre_items)
        st.download_button(
            label="⬇  DESCARGAR EXCEL DESVÍOS",
            data=excel_buf,
            file_name=f"DESVIOS_{mes.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with c2:
        st.markdown("### 📑 PowerPoint")
        st.markdown("Generación automática del informe completo con todos los gráficos.")
        st.markdown('<div class="alert-warn">⚙️ Próximamente — en desarrollo</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("◀  VOLVER AL DASHBOARD"):
        st.session_state.step = 3
        st.rerun()
