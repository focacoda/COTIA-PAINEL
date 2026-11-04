import openpyxl, json, pathlib
from collections import defaultdict

EQ = '/sessions/happy-lucid-turing/mnt/ELEIÇÃO 2026/CADASTRO EQUIPE/ZONA_SECAO_EQUIPE.xlsx'
VOT = '/sessions/happy-lucid-turing/mnt/ELEIÇÃO 2026/votaçao 227-286.xlsx'

# 1. Equipe
wb = openpyxl.load_workbook(EQ, read_only=True, data_only=True)
assignments = []
for sn in wb.sheetnames:
    ws = wb[sn]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        continue
    nm = str(rows[1][0]).strip() if rows[1][0] else sn.strip()
    for r in rows[1:]:
        if r[1] is not None and r[2] is not None:
            assignments.append((nm, str(r[1]).strip(), str(r[2]).strip()))
wb.close()

by_sec = defaultdict(set)
for nm, z, s in assignments:
    by_sec[(z, s)].add(nm)

exclusive_secs = {k for k, v in by_sec.items() if len(v) == 1}
shared_secs = {k for k, v in by_sec.items() if len(v) > 1}

# 2. Votação
wbv = openpyxl.load_workbook(VOT, read_only=True, data_only=True)
ws = wbv['votacao_secao_2026_SP']
rows = list(ws.iter_rows(values_only=True))
h = [str(c).strip().upper() if c else '' for c in rows[0]]
zi = h.index('NR_ZONA'); si = h.index('NR_SECAO'); vi = h.index('QT_VOTOS')
ni = h.index('NM_VOTAVEL'); ci = h.index('DS_CARGO')

atila_votes = {}
alex_votes = {}
atila_total = 0
alex_total = 0
for r in rows[1:]:
    nome = str(r[ni]).strip().upper() if r[ni] else ''
    cargo = str(r[ci]).strip().upper() if r[ci] else ''
    key = (str(r[zi]).strip(), str(r[si]).strip())
    qt = int(r[vi] or 0)
    if 'JACOMUSSI' in nome and 'ESTADUAL' in cargo:
        atila_votes[key] = atila_votes.get(key, 0) + qt
        atila_total += qt
    if 'ALEXANDRE PEREIRA' in nome and 'FEDERAL' in cargo:
        alex_votes[key] = alex_votes.get(key, 0) + qt
        alex_total += qt
wbv.close()

# 3. Por equipe
team_secs = defaultdict(set)
for nm, z, s in assignments:
    team_secs[nm].add((z, s))

records = []
for nm, secs in team_secs.items():
    excl = [k for k in secs if k in exclusive_secs]
    comp = [k for k in secs if k in shared_secs]
    records.append({
        'name': nm,
        'atila_excl': sum(atila_votes.get(k, 0) for k in excl),
        'alex_excl': sum(alex_votes.get(k, 0) for k in excl),
        'atila_shared': sum(atila_votes.get(k, 0) for k in comp),
        'alex_shared': sum(alex_votes.get(k, 0) for k in comp),
        'secs_excl': len(excl),
        'secs_shared': len(comp),
        'secs': sorted(secs),
    })

for r in records:
    r['atila_total'] = r['atila_excl'] + r['atila_shared']
    r['alex_total'] = r['alex_excl'] + r['alex_shared']

records.sort(key=lambda r: (-r['atila_total'], -r['alex_total']))
for i, r in enumerate(records, 1):
    r['rank'] = i
    r.pop('secs')

out = {
    'teams': records,
    'summary': {
        'total_unique_sections': len(by_sec),
        'exclusive_sections': len(exclusive_secs),
        'shared_sections': len(shared_secs),
        'assignments': len(assignments),
        'atila_cotia': atila_total,
        'alex_cotia': alex_total,
        'atila_sections': len(atila_votes),
        'alex_sections': len(alex_votes),
    }
}

p = pathlib.Path('/sessions/happy-lucid-turing/mnt/outputs/equipe_resultados.json')
p.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')

print('seções únicas:', len(by_sec), '| exclusivas:', len(exclusive_secs), '| compartilhadas:', len(shared_secs))
print('Átila total Cotia:', atila_total, '| Alexandre total Cotia:', alex_total)
print()
for r in records:
    print(f"{r['rank']:>2}. {r['name']:<26} Átila {r['atila_total']:>5} (ex {r['atila_excl']:>3} + comp {r['atila_shared']:>3}) | Alex {r['alex_total']:>5} (ex {r['alex_excl']:>3} + comp {r['alex_shared']:>3}) | {r['secs_excl']:>2}ex/{r['secs_shared']:>2}comp")
print('\nsalvo:', p)
