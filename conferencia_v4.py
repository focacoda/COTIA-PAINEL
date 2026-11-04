import openpyxl, shutil, os, json
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict, OrderedDict

UP  = '/sessions/happy-lucid-turing/mnt/uploads/ZONA_SECAO_EQUIPE.xlsx'
OUT = '/sessions/happy-lucid-turing/work/ZONA_SECAO_EQUIPE_CONFERENCIA.xlsx'
DATA= '/sessions/happy-lucid-turing/mnt/outputs/cotia_dados.json'

d = json.load(open(DATA, encoding='utf-8'))
secoes = {}
for s in d['secoes']:
    secoes[(str(s['zona']), str(int(s['secao'])))] = s

# votos
VOT = '/sessions/happy-lucid-turing/mnt/ELEIÇÃO 2026/votaçao 227-286.xlsx'
av, xv = defaultdict(int), defaultdict(int)
wbv = openpyxl.load_workbook(VOT, read_only=True, data_only=True)
ws = wbv['votacao_secao_2026_SP']
rows = ws.iter_rows(values_only=True)
head = next(rows)
iZ = head.index('NR_ZONA'); iS = head.index('NR_SECAO')
iV = head.index('QT_VOTOS'); iN = head.index('NM_VOTAVEL'); iC = head.index('DS_CARGO')
for r in rows:
    try:
        k = (str(int(r[iZ])), str(int(r[iS])))
    except (TypeError, ValueError):
        continue
    nome = (r[iN] or '').strip().upper(); cargo = (r[iC] or '').upper()
    v = r[iV] or 0
    if 'ATILA' in nome and 'JACOMUSSI' in nome and 'ESTADUAL' in cargo:
        av[k] += v
    elif 'ALEXANDRE PEREIRA' in nome and 'FEDERAL' in cargo:
        xv[k] += v
wbv.close()

# quem apontou (abas individuais)
wbu = openpyxl.load_workbook(UP, data_only=True)
apont = defaultdict(set)
for sn in wbu.sheetnames:
    if sn.strip().upper() in ('RESUMO',):
        continue
    w = wbu[sn]
    time_atual = None
    for r in w.iter_rows(min_row=2, values_only=True):
        if len(r) < 3:
            continue
        if r[0] and str(r[0]).strip():
            time_atual = str(r[0]).strip()
        z, s = r[1], r[2]
        if z is None or s is None:
            continue
        try:
            k = (str(int(z)), str(int(s)))
        except (TypeError, ValueError):
            continue
        apont[k].add(time_atual or sn.strip())
wbu.close()

# ler RESUMO preservando
if os.path.exists(OUT):
    os.remove(OUT)
shutil.copy(UP, OUT)
os.chmod(OUT, 0o644)
wb = openpyxl.load_workbook(OUT)
wsr = wb['RESUMO']

registros = []
for i, r in enumerate(wsr.iter_rows(min_row=2, values_only=True), start=2):
    if not r or r[0] is None:
        continue
    nome = str(r[0]).strip()
    if nome.upper().startswith('OBS'):
        continue
    z, s, time = r[1], r[2], r[3]
    if z is None or s is None:
        continue
    try:
        k = (str(int(z)), str(int(s)))
    except (TypeError, ValueError):
        continue
    info = secoes.get(k)
    n_pes = len(apont.get(k, set()))
    if info:
        registros.append([nome, time or '', int(z), int(s),
                          av.get(k, 0), xv.get(k, 0),
                          info['escola'], info.get('bairro', ''),
                          'OK', n_pes])
    else:
        registros.append([nome, time or '', int(z), int(s),
                          0, 0, 'NÃO EXISTE NA BASE', '',
                          'SEÇÃO INVÁLIDA', n_pes])

# ---------- estilos ----------
HDR   = PatternFill('solid', fgColor='1F3A5F')
HDRF  = Font(color='FFFFFF', bold=True, size=10)
BANDA = PatternFill('solid', fgColor='F4F7FB')
ALERTA= PatternFill('solid', fgColor='FDE7E7')
th = Side(style='thin', color='D8DEE8')
BORD  = Border(left=th, right=th, top=th, bottom=th)

def enfeitar(w, larguras, freeze='A2'):
    for j, wd in enumerate(larguras, 1):
        w.column_dimensions[openpyxl.utils.get_column_letter(j)].width = wd
    for c in w[1]:
        c.fill = HDR; c.font = HDRF
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    w.row_dimensions[1].height = 30
    w.freeze_panes = freeze
    w.auto_filter.ref = w.dimensions

# ---------- aba CONFERENCIA ----------
if 'CONFERENCIA' in wb.sheetnames:
    del wb['CONFERENCIA']
w = wb.create_sheet('CONFERENCIA', 0)
w.append(['Nome', 'Time / Equipe', 'Zona', 'Seção', 'Átila', 'Alexandre',
          'Escola', 'Bairro', 'Situação', 'Pessoas nesta seção'])
for reg in registros:
    w.append(reg)
for i, c in enumerate(w.iter_rows(min_row=2), start=2):
    for cell in c:
        cell.border = BORD
        cell.alignment = Alignment(vertical='center')
    if reg := registros[i-2]:
        if reg[8] != 'OK':
            for cell in c:
                cell.fill = ALERTA
        elif i % 2 == 0:
            for cell in c:
                cell.fill = BANDA
enfeitar(w, [24, 22, 8, 9, 10, 12, 40, 26, 18, 14])

# ---------- aba POR_TIME ----------
if 'POR_TIME' in wb.sheetnames:
    del wb['POR_TIME']
w2 = wb.create_sheet('POR_TIME', 1)
w2.append(['Time / Equipe', 'Apontamentos', 'Seções válidas', 'Seções inexistentes',
           'Votos Átila', 'Votos Alexandre', 'Total'])

TIME_FIX = OrderedDict()
for reg in registros:
    t = reg[1] or 'SEM TIME'
    TIME_FIX.setdefault(t, {'n': 0, 'ok': 0, 'bad': 0, 'a': 0, 'x': 0})
    d_ = TIME_FIX[t]
    d_['n'] += 1
    if reg[8] == 'OK':
        d_['ok'] += 1; d_['a'] += reg[4]; d_['x'] += reg[5]
    else:
        d_['bad'] += 1

for t, v in sorted(TIME_FIX.items(), key=lambda kv: -(kv[1]['a'] + kv[1]['x'])):
    w2.append([t, v['n'], v['ok'], v['bad'], v['a'], v['x'], v['a'] + v['x']])
enfeitar(w2, [26, 13, 14, 18, 12, 15, 11])

# ---------- aba SECOES ----------
if 'SECOES' in wb.sheetnames:
    del wb['SECOES']
w3 = wb.create_sheet('SECOES', 2)
w3.append(['Zona', 'Seção', 'Escola', 'Bairro', 'Átila', 'Alexandre', 'Total',
           'Qtd. pessoas que apontaram', 'Times'])

vistos = {}
for reg in registros:
    if reg[8] != 'OK':
        continue
    k = (reg[2], reg[3])
    if k in vistos:
        continue
    vistos[k] = True

por_sec = defaultdict(lambda: {'a': 0, 'x': 0})
for reg in registros:
    if reg[8] == 'OK':
        k = (reg[2], reg[3])
        por_sec[k]['a'] = reg[4]
        por_sec[k]['x'] = reg[5]

for k in sorted(por_sec, key=lambda kk: (-(por_sec[kk]['a'] + por_sec[kk]['x']), kk)):
    info = secoes.get((str(k[0]), str(k[1])), {})
    times = sorted(apont.get((str(k[0]), str(k[1])), set()))
    w3.append([k[0], k[1], info.get('escola', ''), info.get('bairro', ''),
               por_sec[k]['a'], por_sec[k]['x'],
               por_sec[k]['a'] + por_sec[k]['x'],
               len(times), ', '.join(times) if times else '—'])
enfeitar(w3, [8, 9, 40, 26, 10, 12, 9, 16, 46])

# ---------- aba DIVERGENCIAS ----------
if 'DIVERGENCIAS' in wb.sheetnames:
    del wb['DIVERGENCIAS']
w4 = wb.create_sheet('DIVERGENCIAS', 3)
w4.append(['Zona', 'Seção', 'Escola', 'Bairro', 'Átila', 'Alexandre',
           'Pessoas que apontaram', 'Times'])
cont = 0
for k in sorted(por_sec, key=lambda kk: (-(por_sec[kk]['a'] + por_sec[kk]['x']), kk)):
    times = sorted(apont.get((str(k[0]), str(k[1])), set()))
    if len(times) < 2:
        continue
    info = secoes.get((str(k[0]), str(k[1])), {})
    w4.append([k[0], k[1], info.get('escola', ''), info.get('bairro', ''),
               por_sec[k]['a'], por_sec[k]['x'], len(times), ', '.join(times)])
    cont += 1
enfeitar(w4, [8, 9, 40, 26, 10, 12, 16, 60])

wb.save(OUT)
print('Salvo:', OUT)
print('Abas:', wb.sheetnames)
print('Registros RESUMO:', len(registros))
print('Seções válidas:', len(por_sec))
print('Seções com 2+ times:', cont)
print('RESUMO preservada:', wsr.max_row, 'linhas')
