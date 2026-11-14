"""Textual affinity between publications (title + abstract) -> affinity.json for build_kg.py.

Usage:  python build_affinity.py <excel.xlsx> <model_dir> <affinity_decisions.xlsx> <out_dir>

- Recomputes ALL pairs every run (language-mean centering changes slightly when publications are added).
- Model: multilingual-e5-small, int8 ONNX (model_quantized.onnx + tokenizer.json in <model_dir>).
  SHA-256 model_quantized.onnx = f80102d3f2a1229f387d3c81909990d8945513e347b0eab049f7de3c6f98c193
- Pairs with similarity >= THRESHOLD:
    * sharing at least one keyword  -> accepted automatically (basis "shared keywords"), unless decision = NO
    * no shared keyword             -> accepted only with decision = YES; without a decision -> pending review
- Pairs approved earlier (YES) that fall below the threshold stay accepted and are reported.
- Abstract text and vectors are never written to the output (only pair ids and scores).
Outputs: <out_dir>/affinity.json (accepted pairs) and <out_dir>/affinity_pending.xlsx (to review, if any).
Needs: numpy, openpyxl, onnxruntime, tokenizers.
"""
import sys, re, json, hashlib
from pathlib import Path
import numpy as np, openpyxl

THRESHOLD = 0.25
MODEL_SHA = 'f80102d3f2a1229f387d3c81909990d8945513e347b0eab049f7de3c6f98c193'
XLS, MDIR, DEC, OUT = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]), Path(sys.argv[4])
OUT.mkdir(parents=True, exist_ok=True)

h = hashlib.sha256(); h.update((MDIR / 'model_quantized.onnx').read_bytes())
if h.hexdigest() != MODEL_SHA:
    sys.exit('Model file differs from the pinned one: scores would not be comparable. Stop.')

import onnxruntime as ort
from tokenizers import Tokenizer
tok = Tokenizer.from_file(str(MDIR / 'tokenizer.json')); tok.enable_truncation(512); tok.enable_padding()
sess = ort.InferenceSession(str(MDIR / 'model_quantized.onnx'), providers=['CPUExecutionProvider'])
inp = [i.name for i in sess.get_inputs()]

def embed(texts, bs=8):
    out = []
    for i in range(0, len(texts), bs):
        enc = tok.encode_batch(texts[i:i + bs])
        ids = np.array([e.ids for e in enc], dtype=np.int64); am = np.array([e.attention_mask for e in enc], dtype=np.int64)
        feed = {'input_ids': ids, 'attention_mask': am}
        if 'token_type_ids' in inp: feed['token_type_ids'] = np.zeros_like(ids)
        hs = sess.run(None, feed)[0]; m = am[..., None]
        v = (hs * m).sum(1) / m.sum(1); out.append(v / np.linalg.norm(v, axis=1, keepdims=True))
    return np.vstack(out)

wb = openpyxl.load_workbook(XLS, read_only=True)
it = wb['Publications'].iter_rows(values_only=True); head = next(it)
pubs = [dict(zip(head, r)) for r in it if r and r[0]]
def txt(v): v = str(v or '').strip(); return '' if v in ('None', 'not in source', 'to verify') else v
P = [p for p in pubs if len(txt(p.get('abstract'))) > 80]
ids = [p['publication_id'] for p in P]
lang = np.array(['ITA' if str(p.get('language') or '').startswith('ITA') else 'ENG' for p in P])  # SPA (2 items) grouped with ENG
kw = [set(k.strip() for k in str(p.get('keywords_normalized') or '').split(';') if k.strip()) for p in P]

V = embed(['query: ' + txt(p['title']) + '. ' + txt(p['abstract']) for p in P])
for l in set(lang): V[lang == l] -= V[lang == l].mean(0)
V /= np.linalg.norm(V, axis=1, keepdims=True); S = V @ V.T

# decisions: id_a, id_b, decision (YES/NO), note
dec = {}
if DEC.exists():
    ws = openpyxl.load_workbook(DEC, read_only=True).active; it = ws.iter_rows(values_only=True); hd = next(it)
    for r in it:
        r = dict(zip(hd, r))
        if r.get('id_a') and r.get('decision'): dec[frozenset((r['id_a'], r['id_b']))] = str(r['decision']).strip().upper()

idx = {p: i for i, p in enumerate(ids)}
accepted, pending, below = [], [], []
for i in range(len(P)):
    for j in range(i + 1, len(P)):
        s = float(S[i, j]); key = frozenset((ids[i], ids[j])); d = dec.get(key); shared = kw[i] & kw[j]
        if s >= THRESHOLD:
            if d == 'NO': continue
            if shared: accepted.append(dict(a=ids[i], b=ids[j], score=round(s, 3), basis='shared keywords'))
            elif d == 'YES': accepted.append(dict(a=ids[i], b=ids[j], score=round(s, 3), basis='reviewed'))
            else: pending.append((round(s, 3), ids[i], P[i]['year'], P[i]['title'], ids[j], P[j]['year'], P[j]['title']))
        elif d == 'YES':
            accepted.append(dict(a=ids[i], b=ids[j], score=round(s, 3), basis='reviewed'))
            below.append((round(s, 3), ids[i], ids[j]))
for key, d in dec.items():
    if d == 'YES' and not all(k in idx for k in key): print('WARNING approved pair with missing publication/abstract:', sorted(key))

(OUT / 'affinity.json').write_text(json.dumps({'model': 'multilingual-e5-small (int8 ONNX)', 'model_sha256': MODEL_SHA,
    'threshold': THRESHOLD, 'pairs': accepted}, ensure_ascii=False, indent=1), encoding='utf-8')
pf = OUT / 'affinity_pending.xlsx'
if pending:
    w = openpyxl.Workbook(); ws = w.active; ws.title = 'pending'
    ws.append(['similarity', 'id_a', 'year_a', 'title_a', 'id_b', 'year_b', 'title_b', 'decision', 'note'])
    for r in sorted(pending, reverse=True): ws.append(list(r) + ['', ''])
    w.save(pf)
elif pf.exists(): pf.unlink()
print(f'publications with abstract={len(P)} accepted={len(accepted)} '
      f'(auto={sum(a["basis"]=="shared keywords" for a in accepted)}, reviewed={sum(a["basis"]=="reviewed" for a in accepted)}) '
      f'pending={len(pending)} approved-below-threshold={len(below)}')
for b in below: print('  below threshold but approved:', b)
