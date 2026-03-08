"""
ppt_generator.py
Genera el PowerPoint de KPI INTERLOG editando el template original.
Reemplaza los valores en los XMLs de las slides manteniendo el diseño exacto.
"""

import io
import os
import shutil
import tempfile
import zipfile
from collections import Counter

FASA = 'FINNING ARGENTINA SOCIEDAD ANO'
FSM  = 'FINNING SOLUCIONES MINERAS SA'

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'KPI_template.pptx')


def calcular_kpi(items, con_parametros=True):
    total = len(items)
    if total == 0:
        return 0.0, 0, 0
    if con_parametros:
        out = sum(1 for i in items if i['desvio'] and str(i.get('parametro', '')).upper() == 'INTERLOG')
    else:
        out = sum(1 for i in items if i['desvio'])
    in_count = total - out
    return round(in_count / total * 100, 1), in_count, out


def fmt_pct(pct):
    if pct == int(pct):
        return f"{int(pct)}%"
    return f"{pct:.1f}%"


def reemplazar_nth(xml_str, buscar, reemplazar, n=1):
    """Reemplaza la n-ésima ocurrencia de <a:t>buscar</a:t>"""
    tag_b = f'<a:t>{buscar}</a:t>'
    tag_r = f'<a:t>{reemplazar}</a:t>'
    count = 0
    idx = 0
    while True:
        pos = xml_str.find(tag_b, idx)
        if pos == -1:
            # Try with xml:space="preserve"
            tag_b2 = f'<a:t xml:space="preserve">{buscar}</a:t>'
            tag_r2 = f'<a:t xml:space="preserve">{reemplazar}</a:t>'
            pos = xml_str.find(tag_b2, 0)
            if pos == -1:
                break
            # count occurrences of tag_b2
            count2 = 0
            idx2 = 0
            while True:
                p2 = xml_str.find(tag_b2, idx2)
                if p2 == -1:
                    break
                count2 += 1
                if count2 == n:
                    return xml_str[:p2] + tag_r2 + xml_str[p2 + len(tag_b2):]
                idx2 = p2 + 1
            break
        count += 1
        if count == n:
            return xml_str[:pos] + tag_r + xml_str[pos + len(tag_b):]
        idx = pos + 1
    return xml_str


def generar_ppt(lib_items, ofi_items, cm_pre_items, cm_apr_items, mes='MES'):
    # ── KPIs ────────────────────────────────────────────────────────────────
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

    def desvios_out(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre]
        out = sum(1 for i in sub if i['desvio'] and str(i.get('parametro', '')).upper() == 'INTERLOG')
        return out, len(sub)

    lib_fasa_dev_out, lib_fasa_dev_tot = desvios_out(lib_items, 'FASA')
    lib_fsm_dev_out,  lib_fsm_dev_tot  = desvios_out(lib_items, 'FSM')
    ofi_fasa_dev_out, ofi_fasa_dev_tot = desvios_out(ofi_items, 'FASA')
    ofi_fsm_dev_out,  ofi_fsm_dev_tot  = desvios_out(ofi_items, 'FSM')

    def kpi_canal_verde(items, nombre):
        sub = [i for i in items if i['nombre'] == nombre and i['via'] == 'AVION' and i['canal'] == 'VERDE']
        if not sub:
            return 0.0, 0, 0, 0
        pct, inc, out = calcular_kpi(sub, True)
        return pct, inc, out, len(sub)

    cv_fasa_pct, cv_fasa_in, cv_fasa_out, cv_fasa_tot = kpi_canal_verde(lib_items, 'FASA')
    cv_fsm_pct,  cv_fsm_in,  cv_fsm_out,  cv_fsm_tot  = kpi_canal_verde(lib_items, 'FSM')

    rangos = Counter(i['rango'] for i in cm_apr_items if i['rango'])
    tot_apr = sum(rangos.values())
    r0_7  = rangos.get('0 a 7', 0)
    r8_15 = rangos.get('8 a 15', 0)
    r15   = rangos.get('+15', 0)

    def pct_r(v):
        return round(v / tot_apr * 100) if tot_apr else 0

    def label_rango(v, pct_val, empty='Sin expedientes'):
        return empty if v == 0 else f"{v} exp · {pct_val}%"

    # ── Buscar template ──────────────────────────────────────────────────────
    template_candidates = [
        TEMPLATE_PATH,
        '/mnt/user-data/uploads/KPI_INTERLOG_Diciembre_2025__21_.pptx',
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'KPI_INTERLOG_Diciembre_2025__21_.pptx'),
    ]
    template_file = next((p for p in template_candidates if os.path.exists(p)), None)
    if not template_file:
        raise FileNotFoundError("No se encontró KPI_template.pptx. Subilo junto a ppt_generator.py en el repo.")

    with zipfile.ZipFile(template_file, 'r') as z:
        files = {name: z.read(name) for name in z.namelist()}

    def get_slide(n):
        return files[f'ppt/slides/slide{n}.xml'].decode('utf-8')

    def set_slide(n, content):
        files[f'ppt/slides/slide{n}.xml'] = content.encode('utf-8')

    # ── SLIDE 1: Portada ────────────────────────────────────────────────────
    s1 = get_slide(1)
    s1 = reemplazar_nth(s1, 'DICIEMBRE 2025', mes)
    set_slide(1, s1)

    # ── SLIDE 2: Resumen Ejecutivo ──────────────────────────────────────────
    s2 = get_slide(2)
    # 5 KPIs en orden (todos eran "0%")
    for val in [fmt_pct(lib_fasa_pct), fmt_pct(lib_fsm_pct),
                fmt_pct(ofi_fasa_pct), fmt_pct(ofi_fsm_pct), fmt_pct(cm_pct)]:
        s2 = reemplazar_nth(s2, '0%', val)

    # 4x "0 operaciones"
    for val in [f"{lib_fasa_tot} operaciones", f"{lib_fsm_tot} operaciones",
                f"{ofi_fasa_tot} operaciones", f"{ofi_fsm_tot} operaciones"]:
        s2 = reemplazar_nth(s2, '0 operaciones', val)

    # CM expedientes
    s2 = reemplazar_nth(s2, '31 expedientes', f'{cm_tot} expedientes')

    # 4 bloques de desvíos
    for out_v, tot_v in [(lib_fasa_dev_out, lib_fasa_dev_tot),
                         (lib_fsm_dev_out,  lib_fsm_dev_tot),
                         (ofi_fasa_dev_out, ofi_fasa_dev_tot),
                         (ofi_fsm_dev_out,  ofi_fsm_dev_tot)]:
        s2 = reemplazar_nth(s2, '0 OUT', f'{out_v} OUT')
        s2 = reemplazar_nth(s2, 'de 0 ops', f'de {tot_v} ops')

    set_slide(2, s2)

    # ── SLIDE 3: Oficialización ──────────────────────────────────────────────
    s3 = get_slide(3)
    # FASA: TOTAL, IN, OUT (todos eran "0"), KPI% (era "0%")
    s3 = reemplazar_nth(s3, '0', str(ofi_fasa_tot))
    s3 = reemplazar_nth(s3, '0', str(ofi_fasa_in))
    s3 = reemplazar_nth(s3, '0', str(ofi_fasa_out))
    s3 = reemplazar_nth(s3, '0%', fmt_pct(ofi_fasa_pct))
    # FSM
    s3 = reemplazar_nth(s3, '0', str(ofi_fsm_tot))
    s3 = reemplazar_nth(s3, '0', str(ofi_fsm_in))
    s3 = reemplazar_nth(s3, '0', str(ofi_fsm_out))
    s3 = reemplazar_nth(s3, '0%', fmt_pct(ofi_fsm_pct))
    set_slide(3, s3)

    # ── SLIDE 4: Canal Verde Aéreo ───────────────────────────────────────────
    s4 = get_slide(4)
    s4 = reemplazar_nth(s4, '0', str(cv_fasa_tot))
    s4 = reemplazar_nth(s4, '0', str(cv_fasa_in))
    s4 = reemplazar_nth(s4, '0', str(cv_fasa_out))
    s4 = reemplazar_nth(s4, '0%', fmt_pct(cv_fasa_pct))
    s4 = reemplazar_nth(s4, '0', str(cv_fsm_tot))
    s4 = reemplazar_nth(s4, '0', str(cv_fsm_in))
    s4 = reemplazar_nth(s4, '0', str(cv_fsm_out))
    s4 = reemplazar_nth(s4, '0%', fmt_pct(cv_fsm_pct))
    set_slide(4, s4)

    # ── SLIDE 5: Distribución de canales — sin datos numéricos directos ──────
    # Se mantiene el diseño original sin cambios

    # ── SLIDE 6: Certificados Mineros ────────────────────────────────────────
    s6 = get_slide(6)
    # Template: TOTAL=31, IN=0, OUT=31, KPI=0%
    # Reemplazamos OUT (2da ocurrencia de '31') antes que TOTAL (1ra)
    s6 = reemplazar_nth(s6, '31', str(cm_out), 2)  # OUT primero
    s6 = reemplazar_nth(s6, '31', str(cm_tot), 1)  # TOTAL
    s6 = reemplazar_nth(s6, '0',  str(cm_in))      # IN
    s6 = reemplazar_nth(s6, '0%', fmt_pct(cm_pct)) # KPI

    # Rangos aprobados
    s6 = reemplazar_nth(s6, 'Sin expedientes', label_rango(r0_7,  pct_r(r0_7),  'Sin expedientes'))
    s6 = reemplazar_nth(s6, '29 exp · 94%',    label_rango(r8_15, pct_r(r8_15), 'Sin expedientes'))
    s6 = reemplazar_nth(s6, '2 exp · 6%',      label_rango(r15,   pct_r(r15),   'Sin expedientes'))
    s6 = reemplazar_nth(s6, 'Total aprobados: 31 expedientes',
                        f'Total aprobados: {tot_apr} expedientes')
    set_slide(6, s6)

    # ── Escribir PPTX ────────────────────────────────────────────────────────
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf


if __name__ == '__main__':
    from datetime import datetime, timedelta

    def fake_lib(nombre, via, canal, total, out_count):
        limite = {'AVION': {'VERDE': 1, 'NARANJA': 3, 'ROJO': 4},
                  'MARITIMO': {'VERDE': 3}, 'CAMION': {'VERDE': 1}}.get(via, {}).get(canal, 9)
        return [{'razon': FASA if nombre=='FASA' else FSM, 'nombre': nombre,
                 'ref': f'R{i}', 'carpeta': f'C{i}', 'via': via, 'canal': canal,
                 'f_ofi': datetime(2025,12,1), 'f_cancel': datetime(2025,12,3),
                 'hs': limite+2 if i<out_count else limite-0.5, 'limite': limite,
                 'desvio': i<out_count, 'desvio_desc': 'Demora' if i<out_count else '',
                 'parametro': 'INTERLOG' if i<out_count else ''} for i in range(total)]

    def fake_ofi(nombre, via, total, out_count):
        limite = 48 if nombre=='FSM' and via=='MARITIMO' else 24
        return [{'razon': FASA if nombre=='FASA' else FSM, 'nombre': nombre,
                 'ref': f'O{i}', 'carpeta': f'OC{i}', 'via': via,
                 'f_ofi': datetime(2025,12,1), 'f_ult': datetime(2025,12,2),
                 'hs': limite+5 if i<out_count else limite-2, 'limite': limite,
                 'desvio': i<out_count, 'desvio_desc': 'Demora' if i<out_count else '',
                 'parametro': 'INTERLOG' if i<out_count else ''} for i in range(total)]

    lib_items = (fake_lib('FASA','AVION','VERDE',20,1) + fake_lib('FASA','AVION','NARANJA',5,0) +
                 fake_lib('FASA','MARITIMO','VERDE',10,2) + fake_lib('FSM','AVION','VERDE',15,0) +
                 fake_lib('FSM','MARITIMO','VERDE',12,1))
    ofi_items = (fake_ofi('FASA','AVION',18,1) + fake_ofi('FASA','MARITIMO',10,0) +
                 fake_ofi('FSM','AVION',14,0) + fake_ofi('FSM','MARITIMO',8,2))
    cm_pre = [{'carpeta':f'C{i}','exp':f'EXP{i:03}','f_tad':datetime(2025,12,1),
               'f_ult':datetime(2025,12,3),'hs':50 if i<2 else 30,'desvio':i<2,
               'desvio_desc':'Demora' if i<2 else '','parametro':'ADUANA' if i<2 else ''}
              for i in range(31)]
    cm_apr = [{'carpeta':f'C{i}','exp':f'EXP{i:03}','f_inicio':datetime(2025,11,1),
               'f_apro':datetime(2025,11,1)+timedelta(days=3 if i<29 else 10),
               'dias':3 if i<29 else 10,'rango':'0 a 7' if i<29 else '8 a 15'}
              for i in range(31)]

    print("Generando PPT con diseño original...")
    buf = generar_ppt(lib_items, ofi_items, cm_pre, cm_apr, mes='DICIEMBRE 2025')
    with open('/mnt/user-data/outputs/KPI_OUTPUT.pptx', 'wb') as f:
        f.write(buf.read())
    print("✅ Listo: /mnt/user-data/outputs/KPI_OUTPUT.pptx")
