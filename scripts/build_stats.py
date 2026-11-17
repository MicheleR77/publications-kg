"""Summary indicators for the CV (same Excel as the KG).

Usage: python build_stats.py <excel.xlsx> <out_dir>
Writes stats.json and stats_chart.png (language and open-access bars).
Scope: the publications in the Excel/KG (articles, chapters, conference papers, monographs).
"""
import sys, json, collections
from pathlib import Path
import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

XLS, OUT = Path(sys.argv[1]), Path(sys.argv[2]); OUT.mkdir(parents=True, exist_ok=True)
ws = openpyxl.load_workbook(XLS, read_only=True)['Publications']
it = ws.iter_rows(values_only=True); head = next(it)
P = [dict(zip(head, r)) for r in it if r and r[0]]
ME = 'Russo, Michele'
sp = lambda v: [a.strip() for a in str(v or '').split(';') if a.strip()]
auth = [sp(p['authors']) for p in P]
co = collections.Counter(a for au in auth for a in au if a != ME)
lang = collections.Counter(str(p['language'] or 'n.d.') for p in P)
oa = collections.Counter({'yes': 0, 'no': 0, 'to verify': 0})
oa.update(str(p['open_access']) if p['open_access'] in ('yes', 'no') else 'to verify' for p in P)
S = {'publications': len(P), 'coauthors': len(co), 'mean_authors': round(sum(map(len, auth)) / len(P), 1),
     'single_author': sum(1 for au in auth if au == [ME]),
     'language': dict(lang), 'open_access': dict(oa)}
(OUT / 'stats.json').write_text(json.dumps(S, ensure_ascii=False, indent=1), encoding='utf-8')

# --- chart: two 100% bars with direct labels ---------------------------------
plt.rcParams.update({'font.family': 'Carlito', 'font.size': 8.5})
INK, MUTED, SURF = '#0b0b0b', '#52514e', '#ffffff'
LBL = {'ENG': 'Inglese', 'ITA': 'Italiano', 'ITA/ENG': 'Bilingue', 'SPA': 'Spagnolo'}
COL = {'ENG': '#2a78d6', 'ITA': '#eb6834', 'ITA/ENG': '#1baf7a', 'SPA': '#eda100'}
rows = [('Lingua', [(LBL.get(k, k), lang[k], COL.get(k, '#888')) for k in ['ENG', 'ITA', 'ITA/ENG', 'SPA'] if lang[k]]),
        ('Open access', [('Aperte', oa['yes'], '#2a78d6'), ('Chiuse', oa['no'], '#52514e'), ('Da verificare', oa['to verify'], '#d9d8d4')])]
fig, ax = plt.subplots(figsize=(6.3, 1.45), dpi=200)
n = len(P)
for y, (name, segs) in enumerate(reversed(rows)):
    x = 0; small = []
    for lab, v, c in segs:
        w = v / n
        ax.barh(y, w - 0.004, left=x, height=0.5, color=c, edgecolor=SURF, linewidth=0)
        txt = f'{lab} {v}'
        if w > 0.13:
            ax.text(x + 0.008, y, txt, va='center', ha='left', color='white' if c not in ('#d9d8d4', '#eda100') else INK, fontsize=8)
        else:
            small.append(txt)
        x += w
    if small:
        ax.text(1, y + 0.3, ' · '.join(small), ha='right', va='bottom', fontsize=7.5, color=MUTED)
    ax.text(-0.01, y, name, va='center', ha='right', color=INK, fontsize=8.5, fontweight='bold')
ax.set_xlim(0, 1); ax.set_ylim(-0.5, len(rows) - 0.2); ax.axis('off')
fig.subplots_adjust(left=0.15, right=0.99, top=0.95, bottom=0.03)
fig.savefig(OUT / 'stats_chart.png', facecolor=SURF)
print(json.dumps(S, ensure_ascii=False))
