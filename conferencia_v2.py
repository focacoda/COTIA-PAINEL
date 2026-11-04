import openpyxl, shutil, os, json, pathlib
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, OrderedDict

UP  = '/sessions/happy-lucid-turing/mnt/uploads/ZONA_SECAO_EQUIPE.xlsx'
OUT = '/sessions/happy-lucid-turing/work/ZONA_SECAO_EQUIPE_CONFERENCIA.xlsx'
DADOS = '/sessions/happy-lucid-turing/mnt/outputs/cotia_dados.json'
VOT = '/sessions/happy-lucid-turing/mnt/ELEIÇÃO 2026/votaçao 227-286.xlsx'

if os.path.exists(OUT):
    os.remove(OUT)
shutil.copy(UP, OUT)
os.chmod(OUT, 0o644)

# ---------- base de Cotia ----------
d = json.load(open(DADOS, encoding='utf-8'))
sec_info = {}
for s in d['secoes']:
    sec_info[(str(s['zona']), str(int(s['secao'])))] = s

# votos por seção
import openpyxl as ox
wbv = ox.load_workbook(VOT, read_only=True, data_only=True)
wsv = wbv['votacao_secao_2026_SP']
rows = list(wsv.iter_rows(values_only=True))
hdr = list(rows[0])
iZ = hdr.index('NR_ZONA'); iS = hdr.index('NR_SECAO')
iV = hdr.index('QT_VOTOS'); iN = hdr.index('NM_VOTAVEL'); iC = hdr.index('DS_CARGO')

at_secao = defaultdict(int)
al_secao = defaultdict(int)
for r in rows[1:]:
    nome = str(r[iN] or '').strip().upper()
    cargo = str(r[iC] or '').strip().upper()
    k = (str(r[iZ]).strip(), str(int(r[iS])))
    v = r[iV] or 0
    if nome == 'ATILA CESAR MONTEIRO JACOMUSSI' and 'ESTADUAL' in cargo:
        at_secao[k] += v
    elif nome == 'ALEXANDRE PEREIRA DA SILVA' and 'FEDERAL' in cargo:
        al_secao[k] += v
wbv.close()

# ---------- ler RESUMO do usuário (somente leitura) ----------
wb = openpyxl.load_workbook(OUT)
wsr = wb['RESUMO']

registros = []
for r in wsr.iter_rows(min_row=2, values_only=True):
    if not r or r[0] is None:
        continue
    nome = str(r[0]).strip()
    if nome.upper().startswith('OBS') or nome.upper() == 'NOME':
        continue
    zona = str(r[1]).strip() if r[1] is not None else ''
    secao = str(r[2]).strip() if r[2] is not None else ''
    time_ = str(r[3]).strip() if r[3] is not None else ''
    if zona.replace('.0','').isdigit():
        zona = str(int(float(zona)))
    if secao.replace('.0','').isdigit():
        secao = str(int(float(secao)))
    registros.append({'nome': nome, 'zona': zona, 'secao': secao, 'time': time_})

# quantas pessoas apontaram cada seção
contagem = defaultdict(list)
for g in registros:
    contagem[(g['zona'], g['secao'])].append(g['nome'])

# ---------- cores ----------
AZUL   = '1F3864'
VERDE  = 'E2EFDA'
VERM   = 'FCE4E4'
AMAR   = 'FFF2CC'
CINZA  = 'F2F2F2'
branco = Font(color='FFFFFF', bold=True, size=10)
thin   = Side(style='thin', color='BFBFBF')
borda  = Border(left=thin, right=thin, top=thin, bottom=thin)

def cabecalho(ws, cols, row=1):
    for j, c in enumerate(cols, 1):
        cel = ws.cell(row=row, column=j, value=c)
        cel.font = branco
        cel.fill = PatternFill('solid', fgColor=AZUL)
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cel.border = borda
    ws.row_dimensions[row].height = 28
    ws.freeze_panes = ws.cell(row=row+1, column=1)

# ================= ABA 1: CONFERÊNCIA =================
if 'CONFERENCIA' in wb.sheetnames:
    del wb['CONFERENCIA']
ws1 = wb.create_sheet('CONFERENCIA')
cabecalho(ws1, ['Nome', 'Time / Equipe', 'Zona', 'Seção', 'Átila', 'Alexandre',
                'Escola', 'Bairro', 'Situação', 'Pessoas nesta seção'])

for i, g in enumerate(registros, start=2):
    k = (g['zona'], g['secao'])
    s = sec_info.get(k)
    if s:
        at, al = at_secao.get(k, 0), al_secao.get(k, 0)
        escola = s.get('escola', '')
        bairro = s.get('bairro', '')
        sit = 'OK' if (at or al) else 'SEM VOTO'
        cor = VERDE if (at or al) else AMAR
        if at == 0 and al == 0:
            sit = 'SEM VOTO'
    else:
        at, al = 0, 0
        escola = 'NÃO EXISTE NA BASE'
        bairro = ''
        sit = 'SEÇÃO INVÁLIDA'
        cor = VERM
    vals = [g['nome'], g['time'], g['zona'], g['secao'], at, al, escola, bairro,
            sit, len(contagem[k])]
    for j, v in enumerate(vals, 1):
        cel = ws1.cell(row=i, column=j, value=v)
        cel.border = borda
        cel.fill = PatternFill('solid', fgColor=cor)
        if j in (3, 4, 5, 6, 10):
            cel.alignment = Alignment(horizontal='center')

calc = {'A':108,'B':20,'C':7,'D':8,'E':8,'F':11,'G':34,'H':22,'I':17,'J':12}
for c, w in calc.items():
    ws1.column_dimensions[c].width = w

# ================= ABA 2: POR TIME =================
if 'POR_TIME' in wb.sheetnames:
    del wb['POR_TIME']
ws2 = wb.create_sheet('POR_TIME')
cabecalho(ws2, ['Time / Equipe', 'Pessoas', 'Seções distintas', 'Seções válidas',
                'Seções inválidas', 'Átila', 'Alexandre'])

por_time = OrderedDict()
for g in registros:
    t = g['time'] or '(sem time)'
    d_ = por_time.setdefault(t, {'nomes': [], 'secs': set(), 'at': 0, 'al': 0, 'inv': 0})
    d_['nomes'].append(g['nome'])
    k = (g['zona'], g['secao'])
    if k not in d_['secs']:
        d_['secs'].add(k)
        if k in sec_info:
            d_['at'] += at_secao.get(k, 0)
            d_['al'] += al_secao.get(k, 0)
        else:
            d_['inv'] += 1

i = 2
for t, v in sorted(por_time.items(), key=lambda x: -len(x[1]['secs'])):
    vals = [t, len(v['nomes']), len(v['secs']), len(v['secs']) - v['inv'], v['inv'], v['at'], v['al']]
    for j, val in enumerate(vals, 1):
        cel = ws2.cell(row=i, column=j, value=val)
        cel.border = borda
        if j > 1:
            cel.alignment = Alignment(horizontal='center')
    i += 1

for c, w in {'A':30,'B':10,'C':16,'D':15,'E':16,'F':10,'G':12}.items():
    ws2.column_dimensions[c].width = w

# ================= ABA 3: SEÇÕES =================
if 'SECOES' in wb.sheetnames:
    del wb['SECOES']
ws3 = wb.create_sheet('SECOES')
cabecalho(ws3, ['Zona', 'Seção', 'Escola', 'Bairro', 'Átila', 'Alexandre',
                'Total', 'Pessoas que apontaram', 'Quem apontou'])

secs = defaultdict(list)
for g in registros:
    secs[(g['zona'], g['secao'])].append(g['nome'])

i = 2
for k in sorted(secs, key=lambda x: (x[0], int(x[1]) if x[1].isdigit() else 9999)):
    s = sec_info.get(k)
    at = at_secao.get(k, 0) if s else 0
    al = al_secao.get(k, 0) if s else 0
    vals = [k[0], k[1],
            s['escola'] if s else 'NÃO EXISTE NA BASE',
            s.get('bairro','') if s else '',
            at, al, at + al, len(set(secs[k])),
            ', '.join(sorted(set(secs[k])))]
    cor = VERM if not s else (VERDE if (at or al) else AMAR)
    for j, v in enumerate(vals, 1):
        cel = ws3.cell(row=i, column=j, value=v)
        cel.border = borda
        cel.fill = PatternFill('solid', fgColor=cor)
        if j in (1,2,5,6,7,8):
            cel.alignment = Alignment(horizontal='center')
    i += 1

for c, w in {'A':7,'B':8,'C':40,'D':24,'E':8,'F':11,'G':8,'H':12,'I':46}.items():
    ws3.column_dimensions[c].width = w

# ================= ABA 4: DIVERGÊNCIAS =================
if 'DIVERGENCIAS' in wb.sheetnames:
    del wb['DIVERGENCIAS']
ws4 = wb.create_sheet('DIVERGENCIAS')
cabecalho(ws4, ['Zona', 'Seção', 'Problema', 'Detalhe', 'Pessoas envolvidas'])

i = 2
for k, nomes in sorted(secs.items(), key=lambda x: -len(x[1])):
    s = sec_info.get(k)
    if not s:
        vals = [k[0], k[1], 'SEÇÃO NÃO EXISTE', 'Zona/seção não encontrada em Cotia', ', '.join(sorted(set(nomes)))]
    elif len(set(nomes)) > 1:
        at, al = at_secao.get(k,0), al_secao.get(k,0)
        vals = [k[0], k[1], 'SEÇÃO COMPARTILHADA', f'{len(set(nomes))} pessoas apontaram | Átila {at} + Alexandre {al}', ', '.join(sorted(set(nomes)))]
    elif not at_secao.get(k,0) and not al_secao.get(k,0):
        vals = [k[0], k[1], 'SEM VOTO', 'Nenhum voto dos dois candidatos', ', '.join(sorted(set(nomes)))]
    else:
        continue
    cor = VERM if vals[2] == 'SEÇÃO NÃO EXISTE' else AMAR
    for j, v in enumerate(vals, 1):
        cel = ws4.cell(row=i, column=j, value=v)
        cel.border = borda
        cel.fill = PatternFill('solid', fgColor=cor)
        if j in (1,2):
            cel.alignment = Alignment(horizontal='center')
    i += 1

for c, w in {'A':7,'B':8,'C':26,'D':52,'E':62}.items():
    ws4.column_dimensions[c].width = w

wb.save(OUT)
print('Salvo:', OUT)
print('Abas novas:', [s for s in wb.sheetnames if s in ('CONFERENCIA','POR_TIME','SECOES','DIVERGENCIAS')])
print('Registros:', len(registros), '| Seções distintas:', len(secs), '| Times:', len(por_time))
