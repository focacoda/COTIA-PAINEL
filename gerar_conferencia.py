import openpyxl, json, shutil, pathlib
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

UP = '/sessions/happy-lucid-turing/mnt/uploads/ZONA_SECAO_EQUIPE.xlsx'
OUT = '/sessions/happy-lucid-turing/mnt/outputs/ZONA_SECAO_EQUIPE_CONFERENCIA.xlsx'
VOT = '/sessions/happy-lucid-turing/mnt/ELEIÇÃO 2026/votaçao 227-286.xlsx'
BASE = '/sessions/happy-lucid-turing/mnt/outputs/cotia_dados.json'

ATILA = 'ATILA CESAR MONTEIRO JACOMUSSI'
ALEX = 'ALEXANDRE PEREIRA DA SILVA'

# ---------- base Cotia: escola/bairro por (zona,secao) ----------
d = json.load(open(BASE, encoding='utf-8'))
info = {}
for s in d['secoes']:
    info[(str(s['zona']), str(int(s['secao'])))] = (s['escola'], s.get('bairro') or '')
zonas_ok = {str(s['zona']) for s in d['secoes']}

# ---------- votos por (zona,secao) ----------
av, xv = {}, {}
wbv = openpyxl.load_workbook(VOT, read_only=True, data_only=True)
ws = wbv['votacao_secao_2026_SP']
rows = ws.iter_rows(values_only=True)
hdr = next(rows)
iZ = hdr.index('NR_ZONA'); iS = hdr.index('NR_SECAO')
iV = hdr.index('QT_VOTOS'); iN = hdr.index('NM_VOTAVEL'); iC = hdr.index('DS_CARGO')
for r in rows:
    if r[iV] is None:
        continue
    k = (str(int(r[iZ])), str(int(r[iS])))
    v = r[iV] or 0
    nome = (r[iN] or '').strip().upper()
    cargo = (r[iC] or '').strip().upper()
    if nome == ATILA and 'ESTADUAL' in cargo:
        av[k] = av.get(k, 0) + v
    elif nome == ALEX and 'FEDERAL' in cargo:
        xv[k] = xv.get(k, 0) + v
wbv.close()

# ---------- ler RESUMO preservando layout ----------
shutil.copy(UP, OUT)
wb = openpyxl.load_workbook(OUT)
wsr = wb['RESUMO']
rows = list(wsr.iter_rows(values_only=True))

registros = []
obs = None
for i, r in enumerate(rows, start=1):
    if not r or r[0] is None:
        continue
    # Detectar blocos de nomes
    if isinstance(r[0], str) and r[0].strip().lower() == 'nome':
        continue
    # Linha de observação
    if isinstance(r[0], str) and r[0].strip().startswith('OBS'):
        obs = r[0]
        continue

    nome = str(r[0]).strip() if r[0] else ''
    zona = str(int(r[1])) if r[1] is not None and str(r[1]).replace('.','').isdigit() else str(r[1] or '').strip()
    secao = str(int(r[2])) if r[2] is not None and str(r[2]).replace('.','').isdigit() else str(r[2] or '').strip()
    equipe = str(r[3]).strip() if r[3] else ''

    if not nome or not zona or not secao:
        continue

    # Verificar existência
    existe = (zona, secao) in info
    escola_base, bairro_base = info.get((zona, secao), ('', ''))

    # Votos
    votos_atila = av.get((zona, secao), 0) if existe else None
    votos_alex = xv.get((zona, secao), 0) if existe else None

    # Escola e Bairro da base
    escola_real = escola_base if existe else 'NÃO EXISTE NA BASE'
    bairro_real = bairro_base if existe else ''

    registros.append({
        'nome': nome,
        'zona': zona,
        'secao': secao,
        'equipe': equipe,
        'atila_input': r[4] if len(r) > 4 else '',
        'alex_input': r[5] if len(r) > 5 else '',
        'escola_input': r[6] if len(r) > 6 else '',
        'bairro_input': r[7] if len(r) > 7 else '',
        'existe': existe,
        'votos_atila': votos_atila,
        'votos_alex': votos_alex,
        'escola_real': escola_real,
        'bairro_real': bairro_real,
        'linha_origem': i
    })

# ---------- Criar aba CONFERENCIA_IA ----------
if 'CONFERENCIA_IA' in wb.sheetnames:
    wb.remove(wb['CONFERENCIA_IA'])
wsc = wb.create_sheet('CONFERENCIA_IA')

# Cabeçalho
headers = ['Nome', 'ZONA', 'SEÇÃO', 'TIME/EQUIPE', 'ÁTILA (Base)', 'ALEXANDRE (Base)', 'ESCOLA (Base)', 'BAIRRO (Base)', 'Status', 'Linha Origem']
wsc.append(headers)

# Estilo cabeçalho
fill_header = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
font_white = Font(color="FFFFFF", bold=True)
for col_num, header in enumerate(headers, 1):
    cell = wsc.cell(row=1, column=col_num)
    cell.fill = fill_header
    cell.font = font_white
    cell.alignment = Alignment(horizontal='center')

# Preencher dados
for reg in registros:
    status = 'OK' if reg['existe'] else 'INVÁLIDO'
    row = [
        reg['nome'],
        reg['zona'],
        reg['secao'],
        reg['equipe'],
        reg['votos_atila'] if reg['votos_atila'] is not None else 'N/A',
        reg['votos_alex'] if reg['votos_alex'] is not None else 'N/A',
        reg['escola_real'],
        reg['bairro_real'],
        status,
        reg['linha_origem']
    ]
    wsc.append(row)

    # Colorir status
    idx = len(list(wsc.iter_rows()))
    cell_status = wsc.cell(row=idx, column=9)
    if status == 'INVÁLIDO':
        cell_status.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        cell_status.font = Font(color="FFFFFF", bold=True)

# Ajustar larguras
widths = [25, 8, 8, 25, 12, 12, 35, 25, 12, 12]
for i, w in enumerate(widths, 1):
    wsc.column_dimensions[get_column_letter(i)].width = w

# ---------- Criar aba RESUMO_EQUIPES_IA ----------
if 'RESUMO_EQUIPES_IA' in wb.sheetnames:
    wb.remove(wb['RESUMO_EQUIPES_IA'])
wsre = wb.create_sheet('RESUMO_EQUIPES_IA')

wsre.append(['EQUIPE', 'TOTAL APONTAMENTOS', 'SEÇÕES VÁLIDAS', 'SEÇÕES INVÁLIDAS', 'VOTOS ÁTILA (SOMA)', 'VOTOS ALEXANDRE (SOMA)'])
fill_eq = PatternFill(start_color="0E4C5A", end_color="0E4C5A", fill_type="solid")
for col_num, h in enumerate(['EQUIPE', 'TOTAL APONTAMENTOS', 'SEÇÕES VÁLIDAS', 'SEÇÕES INVÁLIDAS', 'VOTOS ÁTILA (SOMA)', 'VOTOS ALEXANDRE (SOMA)'], 1):
    c = wsre.cell(row=1, column=col_num)
    c.fill = fill_eq
    c.font = font_white

from collections import defaultdict
stats = defaultdict(lambda: {'total': 0, 'validas': 0, 'invalidas': 0, 'atila': 0, 'alex': 0})

for reg in registros:
    eq = reg['equipe'] if reg['equipe'] else 'SEM EQUIPE'
    stats[eq]['total'] += 1
    if reg['existe']:
        stats[eq]['validas'] += 1
        stats[eq]['atila'] += (reg['votos_atila'] or 0)
        stats[eq]['alex'] += (reg['votos_alex'] or 0)
    else:
        stats[eq]['invalidas'] += 1

for eq, s in sorted(stats.items()):
    wsre.append([eq, s['total'], s['validas'], s['invalidas'], s['atila'], s['alex']])

wsre.column_dimensions['A'].width = 30
for i in range(2, 7):
    wsre.column_dimensions[get_column_letter(i)].width = 18

wb.save(OUT)
print(f"Arquivo salvo: {OUT}")
print(f"Total registros processados: {len(registros)}")
print(f"Equipes encontradas: {len(stats)}")