"""
ppt_generator.py  —  INTERLOG KPI Dashboard
Genera el PowerPoint editando el KPI_template.pptx (mismo diseño)
e inyectando la slide 5 de Distribución de Canales con datos reales.
"""

import io, os, zipfile
from collections import Counter

FASA = 'FINNING ARGENTINA SOCIEDAD ANO'
FSM  = 'FINNING SOLUCIONES MINERAS SA'

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'KPI_template.pptx')


def calcular_kpi(items, con_parametros=True):
    total = len(items)
    if total == 0:
        return 0.0, 0, 0
    out = sum(1 for i in items if i['desvio'] and (
        not con_parametros or str(i.get('parametro', '')).upper() == 'INTERLOG'))
    return round((total - out) / total * 100, 1), total - out, out


def fmt_pct(pct):
    return f"{int(pct)}%" if pct == int(pct) else f"{pct:.1f}%"


def _nth(xml, buscar, reemplazar, n=1):
    tag_b = f'<a:t>{buscar}</a:t>'
    tag_r = f'<a:t>{reemplazar}</a:t>'
    count, idx = 0, 0
    while True:
        pos = xml.find(tag_b, idx)
        if pos == -1:
            break
        count += 1
        if count == n:
            return xml[:pos] + tag_r + xml[pos+len(tag_b):]
        idx = pos + 1
    return xml


# ── XML helpers para shapes ────────────────────────────────────────────────

def _sp_rect(sid, x, y, cx, cy, fill_hex, line_hex=None, line_w=None):
    """Genera un shape rectángulo relleno sólido."""
    ln = ''
    if line_hex:
        w_attr = f' w="{line_w}"' if line_w else ''
        ln = f'<a:ln{w_attr}><a:solidFill><a:srgbClr val="{line_hex}"/></a:solidFill></a:ln>'
    else:
        ln = '<a:ln><a:noFill/></a:ln>'
    return f'''
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{sid}" name="r{sid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
          {ln}
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>'''


def _sp_text(sid, x, y, cx, cy, text, sz, bold=False, color='1E2D35',
             align='l', valign='ctr', wrap=True):
    """Genera un textbox."""
    b = '1' if bold else '0'
    algn = align
    wrap_attr = 'wrap="square"' if wrap else 'wrap="none"'
    return f'''
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{sid}" name="t{sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr {wrap_attr} anchor="{valign}"/>
          <a:lstStyle/>
          <a:p><a:pPr algn="{algn}"/>
            <a:r><a:rPr lang="es-AR" sz="{sz}" b="{b}" i="0" dirty="0">
              <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
              <a:latin typeface="Calibri"/>
            </a:rPr><a:t xml:space="preserve">{text}</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>'''


EMU = 914400  # 1 inch in EMU
# Slide: 12192000 x 6858000 EMU  (13.33 x 7.5 inches)
# Márgenes: left=411480, top content ~900000

def _build_slide5_shapes(lib_items):
    """
    Genera todos los shapes XML para la slide 5 de Distribución de Canales.
    Layout: 3 columnas (AVIÓN, MARÍTIMO, CAMIÓN) × 3 filas (VERDE, NARANJA, ROJO)
    Para TOTAL, FASA y FSM.
    """
    shapes = ''
    sid = 50

    VERDE   = '27AE60'
    NARANJA = 'D35400'
    ROJO    = 'C0392B'
    DARK    = '1E2D35'
    TEAL    = '1A7A6E'
    LGRAY   = 'D0D8DE'
    WHITE   = 'FFFFFF'
    BG_ROW  = 'F0F4F6'
    BG_HDR  = 'E2EAEE'

    canal_colors = {'VERDE': VERDE, 'NARANJA': NARANJA, 'ROJO': ROJO}

    vias = [
        ('AVION',    'AVIÓN'),
        ('MARITIMO', 'MARÍTIMO'),
        ('CAMION',   'CAMIÓN'),
    ]
    canales = ['VERDE', 'NARANJA', 'ROJO']

    groups = [
        (None,   'TOTAL GENERAL',    DARK,  TEAL),
        ('FASA', 'FASA',             DARK,  '1A6A8A'),
        ('FSM',  'FSM',              DARK,  '7A4A1A'),
    ]

    # Layout: 3 grupos apilados verticalmente
    # Cada grupo tiene un header + 3 barras
    # Coordenadas en EMU
    left_margin = 411480
    slide_w     = 12192000
    usable_w    = slide_w - left_margin * 2   # ~11369040

    via_w   = usable_w // 3       # ancho de cada columna
    bar_h   = 228600              # 0.25 inch por barra
    hdr_h   = 274320              # header de grupo
    sep_h   = 91440               # separación entre grupos
    top_y   = 914400              # empieza después del título

    for gi, (nombre, label, txt_color, accent_color) in enumerate(groups):
        # filtrar items
        if nombre:
            sub_all = [i for i in lib_items if i['nombre'] == nombre]
        else:
            sub_all = lib_items

        # y de inicio de este grupo
        group_y = top_y + gi * (hdr_h + len(canales) * bar_h + sep_h)

        # Header del grupo (barra de fondo)
        shapes += _sp_rect(sid, left_margin, group_y, usable_w, hdr_h,
                           BG_HDR, accent_color, 9525)
        sid += 1
        shapes += _sp_text(sid, left_margin + 91440, group_y, usable_w, hdr_h,
                           label, 1100, bold=True, color=accent_color,
                           align='l', valign='ctr')
        sid += 1

        for vi, (via_key, via_label) in enumerate(vias):
            sub_via = [i for i in sub_all if i['via'] == via_key]
            total_via = len(sub_via)
            col_x = left_margin + vi * via_w

            if total_via == 0:
                # Columna vacía
                for ci in range(len(canales)):
                    row_y = group_y + hdr_h + ci * bar_h
                    shapes += _sp_rect(sid, col_x, row_y, via_w - 9144, bar_h,
                                       BG_ROW, LGRAY, 3175)
                    sid += 1
                    shapes += _sp_text(sid, col_x + 18288, row_y, via_w, bar_h,
                                       'Sin ops', 800, color='AAAAAA',
                                       align='l', valign='ctr')
                    sid += 1
                continue

            by_canal = Counter(i['canal'] for i in sub_via)

            for ci, canal in enumerate(canales):
                cnt = by_canal.get(canal, 0)
                pct = round(cnt / total_via * 100)
                cc = canal_colors[canal]
                row_y = group_y + hdr_h + ci * bar_h

                # Background
                shapes += _sp_rect(sid, col_x, row_y, via_w - 9144, bar_h,
                                   BG_ROW, LGRAY, 3175)
                sid += 1

                # Fill bar proporcional
                if cnt > 0:
                    fill_w = max(9144, int((via_w - 9144) * pct / 100))
                    shapes += _sp_rect(sid, col_x, row_y, fill_w, bar_h,
                                       cc + '60', cc, 0)  # semitransparente visual
                    # Acento sólido izquierdo
                    shapes += _sp_rect(sid+1, col_x, row_y, 18288, bar_h, cc, None, None)
                    sid += 2

                # Texto canal (izq)
                shapes += _sp_text(sid, col_x + 27432, row_y, via_w // 2, bar_h,
                                   canal, 900, bold=True, color=cc,
                                   align='l', valign='ctr')
                sid += 1

                # Texto conteo y % (der)
                label_txt = f"{cnt} · {pct}%" if cnt > 0 else '0'
                shapes += _sp_text(sid, col_x, row_y, via_w - 18288, bar_h,
                                   label_txt, 900, bold=False, color=DARK,
                                   align='r', valign='ctr')
                sid += 1

        # Encabezados de vía (encima del primer grupo, una vez)
        if gi == 0:
            for vi, (via_key, via_label) in enumerate(vias):
                col_x = left_margin + vi * via_w
                # Header de vía
                shapes += _sp_rect(sid, col_x, top_y - hdr_h - 45720,
                                   via_w - 9144, hdr_h, DARK, TEAL, 9525)
                sid += 1
                shapes += _sp_text(sid, col_x, top_y - hdr_h - 45720,
                                   via_w, hdr_h,
                                   via_label, 1100, bold=True, color=WHITE,
                                   align='c', valign='ctr')
                sid += 1

    return shapes


def generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes='MES'):

    # ── KPIs ─────────────────────────────────────────────────────────────────
    def kpi_nom(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre]
        pct, inc, out = calcular_kpi(sub, True)
        return pct, inc, out, len(sub)

    lib_fasa_pct, lib_fasa_in, lib_fasa_out, lib_fasa_tot = kpi_nom(lib_items, 'FASA')
    lib_fsm_pct,  lib_fsm_in,  lib_fsm_out,  lib_fsm_tot  = kpi_nom(lib_items, 'FSM')
    ofi_fasa_pct, ofi_fasa_in, ofi_fasa_out, ofi_fasa_tot = kpi_nom(ofi_items, 'FASA')
    ofi_fsm_pct,  ofi_fsm_in,  ofi_fsm_out,  ofi_fsm_tot  = kpi_nom(ofi_items, 'FSM')
    cm_pct, cm_in, cm_out = calcular_kpi(cm_pre_items, True)
    cm_tot = len(cm_pre_items)

    def dev_out(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre]
        out = sum(1 for i in sub if i['desvio'] and str(i.get('parametro', '')).upper() == 'INTERLOG')
        return out, len(sub)

    lib_fasa_dev_out, lib_fasa_dev_tot = dev_out(lib_items, 'FASA')
    lib_fsm_dev_out,  lib_fsm_dev_tot  = dev_out(lib_items, 'FSM')
    ofi_fasa_dev_out, ofi_fasa_dev_tot = dev_out(ofi_items, 'FASA')
    ofi_fsm_dev_out,  ofi_fsm_dev_tot  = dev_out(ofi_items, 'FSM')

    def kpi_cv(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre and i['via'] == 'AVION' and i['canal'] == 'VERDE']
        if not sub: return 0.0, 0, 0, 0
        pct, inc, out = calcular_kpi(sub, True)
        return pct, inc, out, len(sub)

    cv_fasa_pct, cv_fasa_in, cv_fasa_out, cv_fasa_tot = kpi_cv(lib_items, 'FASA')
    cv_fsm_pct,  cv_fsm_in,  cv_fsm_out,  cv_fsm_tot  = kpi_cv(lib_items, 'FSM')

    rangos = Counter(i['rango'] for i in cm_apr_items if i['rango'])
    tot_apr = sum(rangos.values())
    r0_7  = rangos.get('0 a 7', 0)
    r8_15 = rangos.get('8 a 15', 0)
    r15   = rangos.get('+15', 0)
    def pct_r(v): return round(v / tot_apr * 100) if tot_apr else 0
    def lbl_r(v, p): return 'Sin expedientes' if v == 0 else f"{v} exp · {p}%"

    # ── Template ──────────────────────────────────────────────────────────────
    candidates = [
        TEMPLATE_PATH,
        '/mnt/user-data/uploads/KPI_INTERLOG_Diciembre_2025__21_.pptx',
    ]
    tfile = next((p for p in candidates if os.path.exists(p)), None)
    if not tfile:
        raise FileNotFoundError("No se encontró KPI_template.pptx")

    with zipfile.ZipFile(tfile, 'r') as z:
        files = {n: z.read(n) for n in z.namelist()}

    def gs(n): return files[f'ppt/slides/slide{n}.xml'].decode('utf-8')
    def ss(n, c): files[f'ppt/slides/slide{n}.xml'] = c.encode('utf-8')

    # SLIDE 1
    s1 = gs(1)
    s1 = _nth(s1, 'DICIEMBRE 2025', mes)
    ss(1, s1)

    # SLIDE 2
    s2 = gs(2)
    for v in [fmt_pct(lib_fasa_pct), fmt_pct(lib_fsm_pct),
              fmt_pct(ofi_fasa_pct), fmt_pct(ofi_fsm_pct), fmt_pct(cm_pct)]:
        s2 = _nth(s2, '0%', v)
    for v in [f"{lib_fasa_tot} operaciones", f"{lib_fsm_tot} operaciones",
              f"{ofi_fasa_tot} operaciones", f"{ofi_fsm_tot} operaciones"]:
        s2 = _nth(s2, '0 operaciones', v)
    s2 = _nth(s2, '31 expedientes', f'{cm_tot} expedientes')
    for out_v, tot_v in [(lib_fasa_dev_out, lib_fasa_dev_tot),
                         (lib_fsm_dev_out,  lib_fsm_dev_tot),
                         (ofi_fasa_dev_out, ofi_fasa_dev_tot),
                         (ofi_fsm_dev_out,  ofi_fsm_dev_tot)]:
        s2 = _nth(s2, '0 OUT', f'{out_v} OUT')
        s2 = _nth(s2, 'de 0 ops', f'de {tot_v} ops')
    ss(2, s2)

    # SLIDE 3
    s3 = gs(3)
    for v in [str(ofi_fasa_tot), str(ofi_fasa_in), str(ofi_fasa_out)]:
        s3 = _nth(s3, '0', v)
    s3 = _nth(s3, '0%', fmt_pct(ofi_fasa_pct))
    for v in [str(ofi_fsm_tot), str(ofi_fsm_in), str(ofi_fsm_out)]:
        s3 = _nth(s3, '0', v)
    s3 = _nth(s3, '0%', fmt_pct(ofi_fsm_pct))
    ss(3, s3)

    # SLIDE 4
    s4 = gs(4)
    for v in [str(cv_fasa_tot), str(cv_fasa_in), str(cv_fasa_out)]:
        s4 = _nth(s4, '0', v)
    s4 = _nth(s4, '0%', fmt_pct(cv_fasa_pct))
    for v in [str(cv_fsm_tot), str(cv_fsm_in), str(cv_fsm_out)]:
        s4 = _nth(s4, '0', v)
    s4 = _nth(s4, '0%', fmt_pct(cv_fsm_pct))
    ss(4, s4)

    # SLIDE 5 — Distribución de canales (inyectar shapes reales)
    s5 = gs(5)
    new_shapes = _build_slide5_shapes(lib_items)
    s5 = s5.replace('</p:spTree>', new_shapes + '\n    </p:spTree>')
    ss(5, s5)

    # SLIDE 6
    s6 = gs(6)
    s6 = _nth(s6, '31', str(cm_out), 2)
    s6 = _nth(s6, '31', str(cm_tot), 1)
    s6 = _nth(s6, '0',  str(cm_in))
    s6 = _nth(s6, '0%', fmt_pct(cm_pct))
    s6 = _nth(s6, 'Sin expedientes', lbl_r(r0_7,  pct_r(r0_7)))
    s6 = _nth(s6, '29 exp · 94%',    lbl_r(r8_15, pct_r(r8_15)))
    s6 = _nth(s6, '2 exp · 6%',      lbl_r(r15,   pct_r(r15)))
    s6 = _nth(s6, 'Total aprobados: 31 expedientes',
              f'Total aprobados: {tot_apr} expedientes')
    ss(6, s6)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, d in files.items():
            zout.writestr(n, d)
    buf.seek(0)
    return buf


if __name__ == '__main__':
    from datetime import datetime, timedelta

    def fake_lib(nombre, via, canal, total, out_count):
        limite = {'AVION':{'VERDE':1,'NARANJA':3,'ROJO':4},
                  'MARITIMO':{'VERDE':3,'NARANJA':4,'ROJO':5},
                  'CAMION':{'VERDE':1,'NARANJA':2,'ROJO':3}}.get(via,{}).get(canal,9)
        return [{'razon': FASA if nombre=='FASA' else FSM, 'nombre': nombre,
                 'ref': f'R{i}', 'carpeta': f'C{i}', 'via': via, 'canal': canal,
                 'f_ofi': datetime(2025,12,1), 'f_cancel': datetime(2025,12,3),
                 'hs': limite+2 if i<out_count else limite-0.5, 'limite': limite,
                 'desvio': i<out_count, 'desvio_desc': 'D' if i<out_count else '',
                 'parametro': 'INTERLOG' if i<out_count else ''} for i in range(total)]

    def fake_ofi(nombre, via, total, out_count):
        limite = 48 if nombre=='FSM' and via=='MARITIMO' else 24
        return [{'razon': FASA if nombre=='FASA' else FSM, 'nombre': nombre,
                 'ref': f'O{i}', 'carpeta': f'OC{i}', 'via': via,
                 'f_ofi': datetime(2025,12,1), 'f_ult': datetime(2025,12,2),
                 'hs': limite+5 if i<out_count else limite-2, 'limite': limite,
                 'desvio': i<out_count, 'desvio_desc': '' , 'parametro': 'INTERLOG' if i<out_count else ''}
                for i in range(total)]

    lib  = (fake_lib('FASA','AVION','VERDE',20,1) + fake_lib('FASA','AVION','NARANJA',5,0) +
            fake_lib('FASA','MARITIMO','VERDE',10,2) + fake_lib('FASA','CAMION','VERDE',8,0) +
            fake_lib('FSM','AVION','VERDE',15,0) + fake_lib('FSM','MARITIMO','VERDE',12,1) +
            fake_lib('FSM','MARITIMO','NARANJA',3,0))
    ofi  = (fake_ofi('FASA','AVION',18,1) + fake_ofi('FASA','MARITIMO',10,0) +
            fake_ofi('FSM','AVION',14,0) + fake_ofi('FSM','MARITIMO',8,2))
    cm_pre = [{'carpeta':f'C{i}','exp':f'EXP{i:03}','f_tad':datetime(2025,12,1),
               'f_ult':datetime(2025,12,3),'hs':50 if i<2 else 30,'desvio':i<2,
               'desvio_desc':'','parametro':'INTERLOG' if i<2 else ''} for i in range(31)]
    cm_apr = [{'carpeta':f'C{i}','exp':f'EXP{i:03}','f_inicio':datetime(2025,11,1),
               'f_apro':datetime(2025,11,1)+timedelta(days=3 if i<29 else 10),
               'dias':3 if i<29 else 10,'rango':'0 a 7' if i<29 else '8 a 15'} for i in range(31)]

    buf = generar_ppt(lib, ofi, cm_pre, cm_apr, mes='DICIEMBRE 2025')
    with open('/mnt/user-data/outputs/KPI_FINAL.pptx', 'wb') as f:
        f.write(buf.read())
    print('✅ /mnt/user-data/outputs/KPI_FINAL.pptx')
