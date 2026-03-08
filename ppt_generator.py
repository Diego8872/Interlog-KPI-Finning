"""
ppt_generator.py
Genera el PowerPoint de KPI INTERLOG a partir de los datos procesados.
Usa pptxgenjs vía Node.js para construir el archivo.
"""

import io
import json
import subprocess
import tempfile
import os
from collections import Counter


FASA = 'FINNING ARGENTINA SOCIEDAD ANO'
FSM  = 'FINNING SOLUCIONES MINERAS SA'


def calcular_kpi(items, con_parametros=True):
    total = len(items)
    if total == 0:
        return 0.0, 0, 0
    if con_parametros:
        out = sum(1 for i in items if i['desvio'] and str(i.get('parametro', '')).upper() == 'INTERLOG')
    else:
        out = sum(1 for i in items if i['desvio'])
    in_count = total - out
    return round(in_count / total * 100, 2), in_count, out


def kpi_color(pct):
    if pct >= 95:
        return "00C9A7"
    if pct >= 80:
        return "FF8C42"
    return "FF3D5E"


def preparar_datos(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes):
    """Prepara el JSON de datos que consume el script Node.js"""

    def kpi_por_nombre(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre]
        pct, inc, out = calcular_kpi(sub, True)
        return {'pct': pct, 'in': inc, 'out': out, 'total': len(sub), 'color': kpi_color(pct)}

    def kpi_via(items, nombre, via):
        sub = [i for i in items if i['nombre'] == nombre and i['via'] == via]
        if not sub:
            return None
        pct, inc, out = calcular_kpi(sub, True)
        return {'pct': pct, 'in': inc, 'out': out, 'total': len(sub), 'color': kpi_color(pct)}

    def kpi_canal(items, nombre, via, canal):
        sub = [i for i in items if i['nombre'] == nombre and i['via'] == via and i['canal'] == canal]
        if not sub:
            return None
        pct, inc, out = calcular_kpi(sub, True)
        return {'pct': pct, 'in': inc, 'out': out, 'total': len(sub), 'color': kpi_color(pct)}

    # --- KPIs principales ---
    lib_fasa = kpi_por_nombre(lib_items, 'FASA')
    lib_fsm  = kpi_por_nombre(lib_items, 'FSM')
    ofi_fasa = kpi_por_nombre(ofi_items, 'FASA')
    ofi_fsm  = kpi_por_nombre(ofi_items, 'FSM')
    pct_cm, in_cm, out_cm = calcular_kpi(cm_pre_items, True)
    cm_kpi = {'pct': pct_cm, 'in': in_cm, 'out': out_cm, 'total': len(cm_pre_items), 'color': kpi_color(pct_cm)}

    # --- Desvíos imputables ---
    def desvios_imputables(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre and i['desvio'] and str(i.get('parametro', '')).upper() == 'INTERLOG']
        total = len([i for i in items if i['nombre'] == nombre])
        return {'out': len(sub), 'total': total}

    # --- Oficialización por vía ---
    ofi_vias = {}
    for nombre in ['FASA', 'FSM']:
        ofi_vias[nombre] = {}
        for via in ['AVION', 'MARITIMO', 'CAMION']:
            v = kpi_via(ofi_items, nombre, via)
            if v:
                ofi_vias[nombre][via] = v

    # --- Canal Verde Aéreo ---
    cv_fasa = kpi_canal(lib_items, 'FASA', 'AVION', 'VERDE')
    cv_fsm  = kpi_canal(lib_items, 'FSM',  'AVION', 'VERDE')

    # --- Distribución de canales (liberaciones) ---
    def dist_canales(items, nombre=None):
        sub = [i for i in items if nombre is None or i['nombre'] == nombre]
        by_via = {}
        for via in ['AVION', 'MARITIMO', 'CAMION']:
            v_items = [i for i in sub if i['via'] == via]
            if not v_items:
                continue
            by_canal = Counter(i['canal'] for i in v_items)
            by_via[via] = {
                'total': len(v_items),
                'VERDE':   by_canal.get('VERDE', 0),
                'NARANJA': by_canal.get('NARANJA', 0),
                'ROJO':    by_canal.get('ROJO', 0),
            }
        return by_via

    # --- CM Aprobados rangos ---
    rangos = Counter(i['rango'] for i in cm_apr_items if i['rango'])
    total_apr = sum(rangos.values())
    def pct_rango(k):
        v = rangos.get(k, 0)
        return round(v / total_apr * 100) if total_apr else 0

    # --- Liberación detalle por vía/canal ---
    lib_detalle = {}
    for nombre in ['FASA', 'FSM']:
        lib_detalle[nombre] = {}
        for via in ['AVION', 'MARITIMO', 'CAMION']:
            via_items = [i for i in lib_items if i['nombre'] == nombre and i['via'] == via]
            if not via_items:
                continue
            pct_v, in_v, out_v = calcular_kpi(via_items, True)
            lib_detalle[nombre][via] = {
                'pct': pct_v, 'in': in_v, 'out': out_v,
                'total': len(via_items), 'color': kpi_color(pct_v),
                'canales': {}
            }
            for canal in ['VERDE', 'NARANJA', 'ROJO']:
                c = kpi_canal(lib_items, nombre, via, canal)
                if c:
                    lib_detalle[nombre][via]['canales'][canal] = c

    return {
        'mes': mes,
        'resumen': {
            'lib_fasa': lib_fasa, 'lib_fsm': lib_fsm,
            'ofi_fasa': ofi_fasa, 'ofi_fsm': ofi_fsm,
            'cm': cm_kpi,
            'desvios': {
                'lib_fasa': desvios_imputables(lib_items, 'FASA'),
                'lib_fsm':  desvios_imputables(lib_items, 'FSM'),
                'ofi_fasa': desvios_imputables(ofi_items, 'FASA'),
                'ofi_fsm':  desvios_imputables(ofi_items, 'FSM'),
            }
        },
        'ofi_vias': ofi_vias,
        'canal_verde': {'FASA': cv_fasa, 'FSM': cv_fsm},
        'dist_canales': {
            'general': dist_canales(lib_items),
            'FASA':    dist_canales(lib_items, 'FASA'),
            'FSM':     dist_canales(lib_items, 'FSM'),
        },
        'cm': {
            'total': len(cm_pre_items),
            'in':    in_cm,
            'out':   out_cm,
            'pct':   pct_cm,
            'color': kpi_color(pct_cm),
        },
        'cm_apr': {
            'total': total_apr,
            'r0_7':   rangos.get('0 a 7', 0),
            'r8_15':  rangos.get('8 a 15', 0),
            'r15':    rangos.get('+15', 0),
            'pct0_7':  pct_rango('0 a 7'),
            'pct8_15': pct_rango('8 a 15'),
            'pct15':   pct_rango('+15'),
        },
        'lib_detalle': lib_detalle,
    }


# ─── Node.js script que genera el PPTX ────────────────────────────────────────
NODE_SCRIPT = r"""
const pptxgen = require('pptxgenjs');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const outPath = process.argv[3];

// ── Colores ──────────────────────────────────────────────────────────────────
const C = {
  bg:      '0A1628',
  bg2:     '132236',
  bg3:     '1A2E48',
  accent:  '00C9A7',
  gold:    'FFD060',
  orange:  'FF8C42',
  red:     'FF3D5E',
  gray:    '6B8099',
  lgray:   '9AB0C4',
  white:   'F0F4F8',
  verde:   '00C9A7',
  naranja: 'FF8C42',
  rojo:    'FF3D5E',
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // 10" x 5.625"
pres.title = `KPI INTERLOG · ${data.mes}`;

// ── Helpers ──────────────────────────────────────────────────────────────────
function bg(slide) {
  slide.background = { color: C.bg };
}

function header(slide, title, subtitle) {
  // Barra superior
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: C.bg2 }, line: { color: C.bg2 }
  });
  // Acento izquierdo
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 0.75,
    fill: { color: C.accent }, line: { color: C.accent }
  });
  slide.addText('INTERLOG', {
    x: 0.2, y: 0, w: 2.5, h: 0.75,
    fontSize: 18, bold: true, color: C.accent,
    fontFace: 'Calibri', valign: 'middle', margin: 0
  });
  slide.addText(title, {
    x: 2.8, y: 0, w: 5.5, h: 0.75,
    fontSize: 14, bold: true, color: C.white,
    fontFace: 'Calibri', valign: 'middle', margin: 0, charSpacing: 1
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 8.5, y: 0, w: 1.4, h: 0.75,
      fontSize: 9, color: C.lgray,
      fontFace: 'Calibri', valign: 'middle', align: 'right', margin: 0
    });
  }
}

function footer(slide, mes) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.35, w: 10, h: 0.275,
    fill: { color: C.bg2 }, line: { color: C.bg2 }
  });
  slide.addText(`INTERLOG · KPI ${mes} · interlog.com.ar`, {
    x: 0.2, y: 5.35, w: 9.6, h: 0.275,
    fontSize: 8, color: C.gray, fontFace: 'Calibri',
    valign: 'middle', align: 'center', margin: 0
  });
}

function kpiColor(pct) {
  if (pct >= 95) return C.accent;
  if (pct >= 80) return C.orange;
  return C.red;
}

// Card de KPI grande: número % + label + sub
function kpiCard(slide, x, y, w, h, pct, label, sub, total_label) {
  const color = kpiColor(pct);
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.bg2 }, line: { color: color, pt: 1.5 },
    shadow: { type: 'outer', color: '000000', blur: 8, offset: 2, angle: 135, opacity: 0.2 }
  });
  // Barra top
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.05,
    fill: { color: color }, line: { color: color }
  });
  slide.addText(`${pct.toFixed(0)}%`, {
    x, y: y + 0.1, w, h: h * 0.5,
    fontSize: 40, bold: true, color: color,
    fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
  });
  slide.addText(label, {
    x, y: y + h * 0.58, w, h: 0.3,
    fontSize: 10, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', valign: 'middle',
    charSpacing: 1, margin: 0
  });
  if (sub) {
    slide.addText(sub, {
      x, y: y + h * 0.78, w, h: 0.22,
      fontSize: 8, color: C.lgray,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
  }
}

// Mini card stat (IN / OUT / TOTAL)
function statBox(slide, x, y, w, h, value, label, color) {
  color = color || C.accent;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.bg3 }, line: { color: 'ffffff', pt: 0.3 }
  });
  slide.addText(String(value), {
    x, y: y + 0.04, w, h: h * 0.55,
    fontSize: 22, bold: true, color: color,
    fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
  });
  slide.addText(label, {
    x, y: y + h * 0.6, w, h: h * 0.35,
    fontSize: 8, color: C.lgray, bold: true,
    fontFace: 'Calibri', align: 'center', valign: 'middle',
    charSpacing: 1, margin: 0
  });
}

// Label de sección
function sectionLabel(slide, x, y, w, text, color) {
  color = color || C.accent;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.04,
    fill: { color: color }, line: { color: color }
  });
  slide.addText(text, {
    x: x + 0.1, y: y + 0.02, w: w - 0.15, h: 0.25,
    fontSize: 9, bold: true, color: color,
    fontFace: 'Calibri', valign: 'middle', charSpacing: 1.5, margin: 0
  });
}

// Donut chart nativo pptxgenjs
function donut(slide, x, y, w, h, labels, values, colors) {
  const chartData = [{ name: 'KPI', labels, values }];
  slide.addChart(pres.charts.DOUGHNUT, chartData, {
    x, y, w, h,
    chartColors: colors,
    holeSize: 60,
    showLegend: false,
    showPercent: false,
    showValue: false,
    chartArea: { fill: { color: C.bg2 } },
    dataLabelFontSize: 0,
  });
}

// Bar chart horizontal simulado con shapes
function hBar(slide, x, y, w, h, label, pct, color, total, inc, out) {
  const bw = w;
  const bh = h;
  // Track background
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: bw, h: bh,
    fill: { color: C.bg3 }, line: { color: C.bg3 }
  });
  // Fill
  const fw = Math.max(0.02, bw * pct / 100);
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: fw, h: bh,
    fill: { color: color }, line: { color: color }
  });
  // Label
  slide.addText(`${label}  ${pct.toFixed(1)}%  (${inc} IN · ${out} OUT)`, {
    x: x + 0.08, y, w: bw - 0.1, h: bh,
    fontSize: 9, bold: true, color: C.white,
    fontFace: 'Calibri', valign: 'middle', margin: 0
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — PORTADA
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Bloque lateral izquierdo decorativo
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.3, h: 5.625,
    fill: { color: C.accent }, line: { color: C.accent }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 0, w: 0.08, h: 5.625,
    fill: { color: '007A65' }, line: { color: '007A65' }
  });

  // Centro
  s.addText('INTERLOG', {
    x: 1, y: 1.2, w: 8, h: 0.9,
    fontSize: 60, bold: true, color: C.accent,
    fontFace: 'Calibri', align: 'center', charSpacing: 8, margin: 0
  });
  s.addText('KPI FASA / FSM', {
    x: 1, y: 2.15, w: 8, h: 0.55,
    fontSize: 28, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', charSpacing: 4, margin: 0
  });

  // Línea separadora
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.5, y: 2.8, w: 5, h: 0.04,
    fill: { color: C.accent }, line: { color: C.accent }
  });

  s.addText(data.mes, {
    x: 1, y: 2.9, w: 8, h: 0.5,
    fontSize: 20, color: C.gold, bold: true,
    fontFace: 'Calibri', align: 'center', charSpacing: 3, margin: 0
  });
  s.addText('Reporte de Indicadores de Desempeño', {
    x: 1, y: 3.45, w: 8, h: 0.35,
    fontSize: 13, color: C.lgray,
    fontFace: 'Calibri', align: 'center', margin: 0
  });
  s.addText('interlog.com.ar', {
    x: 1, y: 4.9, w: 8, h: 0.35,
    fontSize: 10, color: C.gray,
    fontFace: 'Calibri', align: 'center', margin: 0
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — RESUMEN EJECUTIVO
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'RESUMEN EJECUTIVO', data.mes);
  footer(s, data.mes);

  const r = data.resumen;

  // Título sección
  sectionLabel(s, 0.3, 0.85, 9.4, 'KPIs PRINCIPALES · TARGET 95%', C.accent);

  // 5 cards KPI en fila
  const cards = [
    { d: r.lib_fasa, label: 'LIBERACIÓN\nFASA',  sub: `${r.lib_fasa.total} ops` },
    { d: r.lib_fsm,  label: 'LIBERACIÓN\nFSM',   sub: `${r.lib_fsm.total} ops` },
    { d: r.ofi_fasa, label: 'OFICIALIZ.\nFASA',  sub: `${r.ofi_fasa.total} ops` },
    { d: r.ofi_fsm,  label: 'OFICIALIZ.\nFSM',   sub: `${r.ofi_fsm.total} ops` },
    { d: r.cm,       label: 'CERT. MIN.\nPRESENTADOS', sub: `${r.cm.total} exp` },
  ];

  const cardW = 1.82, cardH = 1.55, cardY = 1.1, gap = 0.05;
  cards.forEach((c, i) => {
    kpiCard(s, 0.18 + i * (cardW + gap), cardY, cardW, cardH,
            c.d.pct, c.label, c.sub);
  });

  // Desvíos imputables a INTERLOG
  sectionLabel(s, 0.3, 2.8, 9.4, 'DESVÍOS IMPUTABLES A INTERLOG', C.red);

  const dev = r.desvios;
  const devItems = [
    { label: 'FASA · Liberación',     d: dev.lib_fasa },
    { label: 'FSM · Liberación',      d: dev.lib_fsm },
    { label: 'FASA · Oficialización', d: dev.ofi_fasa },
    { label: 'FSM · Oficialización',  d: dev.ofi_fsm },
  ];

  const dw = 2.3, dh = 0.85, dy = 3.05, dgap = 0.1;
  devItems.forEach((item, i) => {
    const dx = 0.3 + i * (dw + dgap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: dx, y: dy, w: dw, h: dh,
      fill: { color: C.bg2 }, line: { color: C.red, pt: 0.8 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: dx, y: dy, w: dw, h: 0.04,
      fill: { color: C.red }, line: { color: C.red }
    });
    s.addText(`${item.d.out}`, {
      x: dx, y: dy + 0.05, w: dw * 0.45, h: dh - 0.1,
      fontSize: 30, bold: true, color: item.d.out > 0 ? C.red : C.accent,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
    s.addText(`OUT\nde ${item.d.total} ops`, {
      x: dx + dw * 0.45, y: dy + 0.1, w: dw * 0.55, h: dh - 0.15,
      fontSize: 9, color: C.lgray,
      fontFace: 'Calibri', valign: 'middle', margin: 0
    });
    s.addText(item.label, {
      x: dx, y: dy + dh - 0.22, w: dw, h: 0.22,
      fontSize: 8, bold: true, color: C.gray,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
      charSpacing: 0.5, margin: 0
    });
  });

  // Nota target
  s.addText('Target: 95%  ·  Parámetro ≠ INTERLOG → operación IN', {
    x: 0.3, y: 5.0, w: 9.4, h: 0.25,
    fontSize: 8, color: C.gray, italic: true,
    fontFace: 'Calibri', align: 'center', margin: 0
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — OFICIALIZACIÓN FASA + FSM
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'OFICIALIZACIÓN · FASA + FSM', data.mes);
  footer(s, data.mes);

  const ofi = data.resumen;
  const nombres = ['FASA', 'FSM'];
  const kpis = [ofi.ofi_fasa, ofi.ofi_fsm];
  const limites = ['24 hs hábiles', '24 hs (48 hs Marítimo)'];

  nombres.forEach((nom, i) => {
    const kd = kpis[i];
    const x = 0.3 + i * 4.85;
    const w = 4.5;

    // Card contenedor
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w, h: 4.3,
      fill: { color: C.bg2 }, line: { color: C.bg3, pt: 0.5 }
    });

    // Header card
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w, h: 0.38,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w: 0.07, h: 0.38,
      fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText(`OFICIALIZACIÓN · ${nom}`, {
      x: x + 0.12, y: 0.88, w: w - 0.15, h: 0.38,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Calibri', valign: 'middle', charSpacing: 1, margin: 0
    });

    // KPI grande
    const color = kpiColor(kd.pct);
    s.addText(`${kd.pct.toFixed(0)}%`, {
      x, y: 1.32, w, h: 1.15,
      fontSize: 68, bold: true, color: color,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
    s.addText('KPI IN', {
      x, y: 2.5, w, h: 0.28,
      fontSize: 10, bold: true, color: C.lgray,
      fontFace: 'Calibri', align: 'center', charSpacing: 2, margin: 0
    });

    // Stats: TOTAL / IN / OUT
    const sw = (w - 0.3) / 3, sh = 0.72, sy = 2.85;
    statBox(s, x + 0.1,           sy, sw, sh, kd.total, 'TOTAL', C.white);
    statBox(s, x + 0.1 + sw,      sy, sw, sh, kd.in,    'IN',    C.accent);
    statBox(s, x + 0.1 + sw * 2,  sy, sw, sh, kd.out,   'OUT',   kd.out > 0 ? C.red : C.accent);

    // Info límite y target
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.1, y: 3.65, w: w - 0.2, h: 0.55,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addText([
      { text: 'TARGET  ', options: { bold: true, color: C.gold } },
      { text: '95%', options: { bold: true, color: C.accent } },
    ], {
      x: x + 0.15, y: 3.67, w: (w - 0.3) / 2, h: 0.25,
      fontSize: 10, fontFace: 'Calibri', valign: 'middle', margin: 0
    });
    s.addText([
      { text: 'LÍMITE  ', options: { bold: true, color: C.gold } },
      { text: limites[i], options: { color: C.lgray } },
    ], {
      x: x + 0.15, y: 3.9, w: w - 0.3, h: 0.25,
      fontSize: 9, fontFace: 'Calibri', valign: 'middle', margin: 0
    });

    // Desglose por vía
    const oviNom = data.ofi_vias[nom] || {};
    let vy = 4.28;
    Object.entries(oviNom).forEach(([via, vd]) => {
      const vc = kpiColor(vd.pct);
      hBar(s, x + 0.1, vy, w - 0.2, 0.25, `${via}`, vd.pct, vc, vd.total, vd.in, vd.out);
      vy += 0.28;
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — CANAL VERDE · VÍA AÉREA
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'CANAL VERDE · VÍA AÉREA', data.mes);
  footer(s, data.mes);

  const nombres = ['FASA', 'FSM'];
  const limite_label = '1 día hábil';

  nombres.forEach((nom, i) => {
    const kd = data.canal_verde[nom];
    const x = 0.3 + i * 4.85;
    const w = 4.5;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w, h: 4.3,
      fill: { color: C.bg2 }, line: { color: C.bg3, pt: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w, h: 0.38,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.88, w: 0.07, h: 0.38,
      fill: { color: C.verde }, line: { color: C.verde }
    });
    s.addText(`CANAL VERDE AÉREO · ${nom}`, {
      x: x + 0.12, y: 0.88, w: w - 0.15, h: 0.38,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Calibri', valign: 'middle', charSpacing: 1, margin: 0
    });

    if (!kd || kd.total === 0) {
      s.addText('Sin operaciones Canal Verde Aéreo', {
        x, y: 2, w, h: 1,
        fontSize: 12, color: C.gray, italic: true,
        fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
      });
      return;
    }

    const color = kpiColor(kd.pct);
    s.addText(`${kd.pct.toFixed(0)}%`, {
      x, y: 1.32, w, h: 1.15,
      fontSize: 68, bold: true, color: color,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
    s.addText('KPI IN', {
      x, y: 2.5, w, h: 0.28,
      fontSize: 10, bold: true, color: C.lgray,
      fontFace: 'Calibri', align: 'center', charSpacing: 2, margin: 0
    });

    const sw = (w - 0.3) / 3, sh = 0.72, sy = 2.85;
    statBox(s, x + 0.1,           sy, sw, sh, kd.total, 'TOTAL', C.white);
    statBox(s, x + 0.1 + sw,      sy, sw, sh, kd.in,    'IN',    C.accent);
    statBox(s, x + 0.1 + sw * 2,  sy, sw, sh, kd.out,   'OUT',   kd.out > 0 ? C.red : C.accent);

    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.1, y: 3.65, w: w - 0.2, h: 0.55,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addText([
      { text: 'TARGET  ', options: { bold: true, color: C.gold } },
      { text: '95%', options: { bold: true, color: C.accent } },
    ], {
      x: x + 0.15, y: 3.67, w: (w - 0.3) / 2, h: 0.25,
      fontSize: 10, fontFace: 'Calibri', valign: 'middle', margin: 0
    });
    s.addText([
      { text: 'LÍMITE  ', options: { bold: true, color: C.gold } },
      { text: limite_label, options: { color: C.lgray } },
    ], {
      x: x + 0.15, y: 3.9, w: w - 0.3, h: 0.25,
      fontSize: 9, fontFace: 'Calibri', valign: 'middle', margin: 0
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — LIBERACIÓN · DETALLE POR VÍA Y CANAL (FASA)
// ══════════════════════════════════════════════════════════════════════════════
function slideLibDetalle(nombre, color_nombre) {
  const s = pres.addSlide();
  bg(s);
  header(s, `LIBERACIÓN · ${nombre}`, data.mes);
  footer(s, data.mes);

  const det = data.lib_detalle[nombre] || {};
  const vias = Object.keys(det);

  if (vias.length === 0) {
    s.addText(`Sin operaciones de liberación para ${nombre}`, {
      x: 1, y: 2.5, w: 8, h: 1,
      fontSize: 16, color: C.gray, italic: true,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
    return;
  }

  sectionLabel(s, 0.3, 0.85, 9.4, 'LIBERACIONES · POR VÍA Y CANAL', color_nombre || C.accent);

  const viaEmoji = { AVION: '✈', CAMION: '▶', MARITIMO: '⬡' };
  const canalColor = { VERDE: C.verde, NARANJA: C.naranja, ROJO: C.rojo };

  let vy = 1.18;
  vias.forEach((via) => {
    const vd = det[via];
    const vc = kpiColor(vd.pct);

    // Via header bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: vy, w: 9.4, h: 0.3,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: vy, w: 0.05, h: 0.3,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(`${viaEmoji[via] || ''} ${via}`, {
      x: 0.42, y: vy, w: 2, h: 0.3,
      fontSize: 10, bold: true, color: C.gold,
      fontFace: 'Calibri', valign: 'middle', charSpacing: 1, margin: 0
    });
    s.addText(`${vd.pct.toFixed(1)}% IN  ·  ${vd.total} ops  ·  ${vd.in} IN · ${vd.out} OUT`, {
      x: 2.5, y: vy, w: 7.1, h: 0.3,
      fontSize: 9, color: vc, bold: true,
      fontFace: 'Calibri', valign: 'middle', align: 'right', margin: 0
    });
    vy += 0.33;

    // Canales
    const canales = Object.keys(vd.canales || {});
    const cw = 3.0, ch = 0.72, cgap = 0.17;
    const totalW = canales.length * cw + (canales.length - 1) * cgap;
    const startX = 0.3 + (9.4 - totalW) / 2;

    canales.forEach((canal, ci) => {
      const cd = vd.canales[canal];
      const cc = canalColor[canal] || C.gray;
      const cx = startX + ci * (cw + cgap);

      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: vy, w: cw, h: ch,
        fill: { color: C.bg2 }, line: { color: cc, pt: 1 }
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: vy, w: cw, h: 0.05,
        fill: { color: cc }, line: { color: cc }
      });
      s.addText(`CANAL ${canal}`, {
        x: cx, y: vy + 0.03, w: cw, h: 0.2,
        fontSize: 8, bold: true, color: cc,
        fontFace: 'Calibri', align: 'center', valign: 'middle',
        charSpacing: 1, margin: 0
      });
      s.addText(`${cd.pct.toFixed(1)}%`, {
        x: cx, y: vy + 0.2, w: cw * 0.55, h: ch - 0.25,
        fontSize: 26, bold: true, color: kpiColor(cd.pct),
        fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
      });
      s.addText([
        { text: `${cd.total} ops\n`, options: { color: C.white, fontSize: 10 } },
        { text: `✓ ${cd.in}  ✗ ${cd.out}`, options: { color: C.lgray, fontSize: 9 } },
      ], {
        x: cx + cw * 0.55, y: vy + 0.18, w: cw * 0.42, h: ch - 0.22,
        fontFace: 'Calibri', valign: 'middle', margin: 0
      });
    });

    vy += ch + 0.15;
  });
}

slideLibDetalle('FASA', C.accent);
slideLibDetalle('FSM', C.orange);

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — DISTRIBUCIÓN DE CANALES
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'DISTRIBUCIÓN DE CANALES · LIBERACIONES', data.mes);
  footer(s, data.mes);

  const dist = data.dist_canales;
  const vias = ['AVION', 'MARITIMO', 'CAMION'];
  const viaLabels = { AVION: '✈ AVIÓN', MARITIMO: '⬡ MARÍTIMO', CAMION: '▶ CAMIÓN' };
  const canales = ['VERDE', 'NARANJA', 'ROJO'];
  const cColors = { VERDE: C.verde, NARANJA: C.naranja, ROJO: C.rojo };

  sectionLabel(s, 0.3, 0.85, 9.4, 'DISTRIBUCIÓN DE CANALES · TODAS LAS RAZONES SOCIALES', C.gold);

  // Una tabla por vía
  const tw = 2.9, th_h = 0.3, th_row = 0.28, tgap = 0.15;
  let tx = 0.3;
  vias.forEach((via) => {
    const gd = dist.general[via];
    if (!gd) { tx += tw + tgap; return; }

    // Header vía
    s.addShape(pres.shapes.RECTANGLE, {
      x: tx, y: 1.1, w: tw, h: th_h,
      fill: { color: C.bg3 }, line: { color: C.gold, pt: 0.8 }
    });
    s.addText(viaLabels[via], {
      x: tx, y: 1.1, w: tw, h: th_h,
      fontSize: 10, bold: true, color: C.gold,
      fontFace: 'Calibri', align: 'center', valign: 'middle', charSpacing: 1, margin: 0
    });

    let ry = 1.1 + th_h;
    canales.forEach((canal) => {
      const cnt = gd[canal] || 0;
      const pct = gd.total > 0 ? cnt / gd.total * 100 : 0;
      const cc = cColors[canal];

      s.addShape(pres.shapes.RECTANGLE, {
        x: tx, y: ry, w: tw, h: th_row,
        fill: { color: C.bg2 }, line: { color: C.bg3 }
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: tx, y: ry, w: Math.max(0.02, tw * pct / 100), h: th_row,
        fill: { color: cc, transparency: 60 }, line: { color: cc, transparency: 60 }
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: tx, y: ry, w: 0.06, h: th_row,
        fill: { color: cc }, line: { color: cc }
      });
      s.addText(`${canal}`, {
        x: tx + 0.1, y: ry, w: tw * 0.45, h: th_row,
        fontSize: 9, bold: true, color: cc,
        fontFace: 'Calibri', valign: 'middle', margin: 0
      });
      s.addText(`${cnt} ops · ${pct.toFixed(0)}%`, {
        x: tx + tw * 0.45, y: ry, w: tw * 0.55 - 0.05, h: th_row,
        fontSize: 9, color: C.white,
        fontFace: 'Calibri', valign: 'middle', align: 'right', margin: 0
      });
      ry += th_row;
    });

    // Total
    s.addShape(pres.shapes.RECTANGLE, {
      x: tx, y: ry, w: tw, h: th_h,
      fill: { color: C.bg3 }, line: { color: C.bg3 }
    });
    s.addText(`TOTAL: ${gd.total} operaciones`, {
      x: tx, y: ry, w: tw, h: th_h,
      fontSize: 9, color: C.lgray, bold: true,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });

    tx += tw + tgap;
  });

  // Por sociedad — FASA y FSM
  ['FASA', 'FSM'].forEach((nom, ni) => {
    const nomColor = ni === 0 ? C.accent : C.orange;
    const labelY = 2.5 + ni * 1.35;
    sectionLabel(s, 0.3, labelY, 9.4, `DISTRIBUCIÓN POR RAZÓN SOCIAL · ${nom}`, nomColor);

    let tx2 = 0.3;
    vias.forEach((via) => {
      const nd = (dist[nom] || {})[via];
      if (!nd) { tx2 += tw + tgap; return; }

      let ry2 = labelY + 0.25;
      canales.forEach((canal) => {
        const cnt = nd[canal] || 0;
        const pct = nd.total > 0 ? cnt / nd.total * 100 : 0;
        const cc = cColors[canal];

        s.addShape(pres.shapes.RECTANGLE, {
          x: tx2, y: ry2, w: tw, h: th_row - 0.02,
          fill: { color: C.bg2 }, line: { color: C.bg3 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
          x: tx2, y: ry2, w: 0.05, h: th_row - 0.02,
          fill: { color: cc }, line: { color: cc }
        });
        s.addText(`${canal}: ${cnt} (${pct.toFixed(0)}%)`, {
          x: tx2 + 0.1, y: ry2, w: tw - 0.12, h: th_row - 0.02,
          fontSize: 8.5, color: C.white,
          fontFace: 'Calibri', valign: 'middle', margin: 0
        });
        ry2 += th_row - 0.02;
      });
      tx2 += tw + tgap;
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — CERTIFICADOS MINEROS
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'CERTIFICADOS MINEROS', data.mes);
  footer(s, data.mes);

  const cm = data.cm;
  const apr = data.cm_apr;

  // ── PRESENTADOS ──
  sectionLabel(s, 0.3, 0.85, 9.4, 'PRESENTADOS · KPI DE GESTIÓN', C.accent);

  // 4 stat cards
  const cards_cm = [
    { v: cm.total,            label: 'TOTAL', color: C.white },
    { v: cm.in,               label: 'IN',    color: C.accent },
    { v: cm.out,              label: 'OUT',   color: cm.out > 0 ? C.red : C.accent },
    { v: `${cm.pct.toFixed(0)}%`, label: 'KPI',   color: kpiColor(cm.pct) },
  ];
  const sw2 = 2.2, sh2 = 0.85, sy2 = 1.12, sgap2 = 0.12;
  cards_cm.forEach((c, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3 + i * (sw2 + sgap2), y: sy2, w: sw2, h: sh2,
      fill: { color: C.bg2 }, line: { color: c.color, pt: 0.8 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3 + i * (sw2 + sgap2), y: sy2, w: sw2, h: 0.04,
      fill: { color: c.color }, line: { color: c.color }
    });
    s.addText(String(c.v), {
      x: 0.3 + i * (sw2 + sgap2), y: sy2 + 0.05, w: sw2, h: sh2 * 0.6,
      fontSize: 30, bold: true, color: c.color,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
    s.addText(c.label, {
      x: 0.3 + i * (sw2 + sgap2), y: sy2 + sh2 * 0.65, w: sw2, h: sh2 * 0.3,
      fontSize: 9, bold: true, color: C.lgray,
      fontFace: 'Calibri', align: 'center', valign: 'middle',
      charSpacing: 1.5, margin: 0
    });
  });

  s.addText('Límite: 48 hs hábiles desde TAD Subido', {
    x: 0.3, y: 2.05, w: 9.4, h: 0.25,
    fontSize: 8.5, color: C.gray, italic: true,
    fontFace: 'Calibri', align: 'center', margin: 0
  });

  // ── APROBADOS ──
  sectionLabel(s, 0.3, 2.38, 9.4, 'APROBADOS · TIEMPO DE APROBACIÓN (informativo)', C.gold);

  // Barras de rangos
  const rangos_apr = [
    { label: '0 a 7 días',  v: apr.r0_7,  pct: apr.pct0_7,  color: C.verde },
    { label: '8 a 15 días', v: apr.r8_15, pct: apr.pct8_15, color: C.naranja },
    { label: '+15 días',    v: apr.r15,   pct: apr.pct15,   color: C.rojo },
  ];

  if (apr.total === 0) {
    s.addText('Sin expedientes aprobados en el período', {
      x: 0.3, y: 2.7, w: 9.4, h: 0.6,
      fontSize: 12, color: C.gray, italic: true,
      fontFace: 'Calibri', align: 'center', valign: 'middle', margin: 0
    });
  } else {
    const bw3 = 9.0, bh3 = 0.55, by3 = 2.65, bgap3 = 0.12;
    rangos_apr.forEach((rg, i) => {
      const ry3 = by3 + i * (bh3 + bgap3);
      // Track
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: ry3, w: bw3, h: bh3,
        fill: { color: C.bg3 }, line: { color: C.bg3 }
      });
      // Fill proporcional
      const fw3 = Math.max(0.05, bw3 * rg.pct / 100);
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: ry3, w: fw3, h: bh3,
        fill: { color: rg.color, transparency: 20 }, line: { color: rg.color, transparency: 20 }
      });
      // Accent left
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: ry3, w: 0.06, h: bh3,
        fill: { color: rg.color }, line: { color: rg.color }
      });
      // Texto
      s.addText(rg.label, {
        x: 0.62, y: ry3, w: 4, h: bh3,
        fontSize: 11, bold: true, color: rg.color,
        fontFace: 'Calibri', valign: 'middle', margin: 0
      });
      s.addText(`${rg.v} exp · ${rg.pct}%`, {
        x: 0.5 + bw3 - 3.5, y: ry3, w: 3.4, h: bh3,
        fontSize: 13, bold: true, color: C.white,
        fontFace: 'Calibri', align: 'right', valign: 'middle', margin: 0
      });
    });

    s.addText(`Total aprobados: ${apr.total} expedientes`, {
      x: 0.3, y: 5.05, w: 9.4, h: 0.22,
      fontSize: 9, color: C.gray,
      fontFace: 'Calibri', align: 'center', margin: 0
    });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// ESCRIBIR ARCHIVO
// ══════════════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: outPath }).then(() => {
  console.log('OK');
}).catch(err => {
  console.error('ERROR:', err);
  process.exit(1);
});
"""


def generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes='MES'):
    """
    Genera el PowerPoint y retorna un BytesIO con el archivo.
    Compatible con la llamada desde app.py:
        from ppt_generator import generar_ppt
        ppt_buf = generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes=...)
    """
    datos = preparar_datos(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes)

    with tempfile.TemporaryDirectory() as tmpdir:
        json_path = os.path.join(tmpdir, 'data.json')
        js_path   = os.path.join(tmpdir, 'gen.js')
        out_path  = os.path.join(tmpdir, 'output.pptx')

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False, default=str)

        with open(js_path, 'w', encoding='utf-8') as f:
            f.write(NODE_SCRIPT)

        result = subprocess.run(
            ['node', js_path, json_path, out_path],
            capture_output=True, text=True, timeout=60
        )

        if result.returncode != 0:
            raise RuntimeError(f"Error en Node.js:\n{result.stderr}\n{result.stdout}")

        with open(out_path, 'rb') as f:
            buf = io.BytesIO(f.read())

    buf.seek(0)
    return buf


# ── Test rápido si se ejecuta directo ────────────────────────────────────────
if __name__ == '__main__':
    print("Generando PPT de prueba con datos sintéticos...")

    from datetime import datetime, timedelta

    def fake_lib(nombre, via, canal, total, out_count):
        items = []
        limite = {'AVION': {'VERDE': 1, 'NARANJA': 3, 'ROJO': 4},
                  'MARITIMO': {'VERDE': 3, 'NARANJA': 4, 'ROJO': 5},
                  'CAMION': {'VERDE': 1, 'NARANJA': 2, 'ROJO': 3}}.get(via, {}).get(canal, 9)
        for i in range(total):
            desvio = i < out_count
            items.append({
                'razon': 'FINNING ARGENTINA SOCIEDAD ANO' if nombre == 'FASA' else 'FINNING SOLUCIONES MINERAS SA',
                'nombre': nombre, 'ref': f'REF{i:03}', 'carpeta': f'C{i:03}',
                'via': via, 'canal': canal,
                'f_ofi': datetime(2025, 12, 1), 'f_cancel': datetime(2025, 12, 3),
                'hs': limite + 2 if desvio else limite - 0.5,
                'limite': limite, 'desvio': desvio,
                'desvio_desc': 'Demora operativa' if desvio else '',
                'parametro': 'INTERLOG' if desvio else ''
            })
        return items

    def fake_ofi(nombre, via, total, out_count):
        FSM = 'FINNING SOLUCIONES MINERAS SA'
        limite = 48 if nombre == 'FSM' and via == 'MARITIMO' else 24
        items = []
        for i in range(total):
            desvio = i < out_count
            items.append({
                'razon': 'FINNING ARGENTINA SOCIEDAD ANO' if nombre == 'FASA' else FSM,
                'nombre': nombre, 'ref': f'OREF{i:03}', 'carpeta': f'OC{i:03}',
                'via': via,
                'f_ofi': datetime(2025, 12, 1), 'f_ult': datetime(2025, 12, 2),
                'hs': limite + 5 if desvio else limite - 2,
                'limite': limite, 'desvio': desvio,
                'desvio_desc': 'Demora doc' if desvio else '',
                'parametro': 'INTERLOG' if desvio else ''
            })
        return items

    lib_items = (
        fake_lib('FASA', 'AVION',    'VERDE',   20, 1) +
        fake_lib('FASA', 'AVION',    'NARANJA',  5, 0) +
        fake_lib('FASA', 'AVION',    'ROJO',     2, 0) +
        fake_lib('FASA', 'MARITIMO', 'VERDE',   10, 2) +
        fake_lib('FASA', 'CAMION',   'VERDE',    8, 0) +
        fake_lib('FSM',  'AVION',    'VERDE',   15, 0) +
        fake_lib('FSM',  'MARITIMO', 'VERDE',   12, 1) +
        fake_lib('FSM',  'MARITIMO', 'NARANJA',  3, 0)
    )
    ofi_items = (
        fake_ofi('FASA', 'AVION',    18, 1) +
        fake_ofi('FASA', 'MARITIMO', 10, 0) +
        fake_ofi('FSM',  'AVION',    14, 0) +
        fake_ofi('FSM',  'MARITIMO',  8, 2)
    )
    cm_pre = [
        {'carpeta': f'C{i}', 'exp': f'EXP{i:03}',
         'f_tad': datetime(2025, 12, 1), 'f_ult': datetime(2025, 12, 3),
         'hs': 50 if i < 2 else 30, 'desvio': i < 2,
         'desvio_desc': 'Demora' if i < 2 else '', 'parametro': 'ADUANA' if i < 2 else ''}
        for i in range(31)
    ]
    cm_apr = [
        {'carpeta': f'C{i}', 'exp': f'EXP{i:03}',
         'f_inicio': datetime(2025, 11, 1),
         'f_apro': datetime(2025, 11, 1) + timedelta(days=3 if i < 29 else 10),
         'dias': 3 if i < 29 else 10,
         'rango': '0 a 7' if i < 29 else '8 a 15'}
        for i in range(31)
    ]

    buf = generar_ppt(lib_items, ofi_items, cm_pre, cm_apr, mes='DICIEMBRE 2025')

    with open('/mnt/user-data/outputs/KPI_TEST.pptx', 'wb') as f:
        f.write(buf.read())
    print("✅ Generado: /mnt/user-data/outputs/KPI_TEST.pptx")
