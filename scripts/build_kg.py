"""Excel -> Knowledge Graph (JSON-LD + Turtle + data for the web viewer).

Usage:  python build_kg.py <excel.xlsx> <out_dir>
Rules: the `notes` column is never exported; `not in source` abstracts become empty.
"""
import sys, re, json, unicodedata
from pathlib import Path
import openpyxl

XLS, OUT = Path(sys.argv[1]), Path(sys.argv[2])
# Base IRI of the graph: replace with the public URL (e.g. https://<user>.github.io/<repo>/) when published.
BASE = sys.argv[3] if len(sys.argv) > 3 else 'urn:mrusso-kg:'
OUT.mkdir(parents=True, exist_ok=True)
wb = openpyxl.load_workbook(XLS, read_only=True)

def rows(sheet):
    it = wb[sheet].iter_rows(values_only=True)
    head = next(it)
    return [dict(zip(head, r)) for r in it if r and r[0]]

def slug(s):
    s = unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode()
    return re.sub(r'[^A-Za-z0-9]+', '-', s).strip('-').lower()

def split(v, sep=';'):
    return [x.strip() for x in str(v or '').split(sep) if x.strip()]

def clean(v):
    if v is None: return None
    v = str(v).strip()
    return None if v in ('', 'None', 'to verify', 'not in source', 'n/a', 'DA COMPLETARE') else v

pubs = rows('Publications')
authors_v = rows('Authors_Vocabulary')
kw_v = rows('Keywords_Vocabulary')
scale_v = rows('Scale_Vocabulary')

TYPE = {'journal_article': 'ScholarlyArticle', 'conference_paper': 'ScholarlyArticle',
        'book_chapter': 'Chapter', 'monograph': 'Book'}
LABEL_TYPE = {'journal_article': 'Journal article', 'conference_paper': 'Conference paper',
              'book_chapter': 'Book chapter', 'monograph': 'Monograph'}

graph = []
nodes = {}   # id -> viewer node
links = []

def vnode(i, label, kind, **extra):
    if i not in nodes:
        nodes[i] = dict(id=i, label=label, kind=kind, **extra)
    return i

# --- vocabularies -------------------------------------------------------
themes = sorted({k['macro_theme'] for k in kw_v if k.get('macro_theme')})
for t in themes:
    tid = f'theme/{slug(t)}'
    graph.append({'@id': tid, '@type': 'skos:Concept', 'skos:prefLabel': t, 'skos:inScheme': {'@id': 'scheme/macro-themes'}})
    vnode(tid, t, 'theme')
kw_index = {}
for k in kw_v:
    kid = f'keyword/{slug(k["canonical_keyword"])}'
    kw_index[k['canonical_keyword']] = kid
    item = {'@id': kid, '@type': 'skos:Concept', 'skos:prefLabel': k['canonical_keyword'],
            'skos:inScheme': {'@id': 'scheme/keywords'}}
    alts = split(k.get('variants_found'))
    if alts: item['skos:altLabel'] = alts
    if k.get('macro_theme'):
        item['skos:broader'] = {'@id': f'theme/{slug(k["macro_theme"])}'}
    graph.append(item)
for s in scale_v:
    sid = f'scale/{slug(s["scale_value"])}'
    graph.append({'@id': sid, '@type': 'skos:Concept', 'skos:prefLabel': s['scale_value'],
                  'skos:definition': s.get('description'), 'skos:inScheme': {'@id': 'scheme/case-study-scale'}})
graph += [{'@id': 'scheme/keywords', '@type': 'skos:ConceptScheme', 'skos:prefLabel': 'Keywords'},
          {'@id': 'scheme/macro-themes', '@type': 'skos:ConceptScheme', 'skos:prefLabel': 'Macro-themes'},
          {'@id': 'scheme/case-study-scale', '@type': 'skos:ConceptScheme', 'skos:prefLabel': 'Case study scale'}]
for a in authors_v:
    pid = f'person/{slug(a["canonical_form"])}'
    item = {'@id': pid, '@type': 'Person', 'name': a['canonical_form']}
    alts = split(a.get('variants_found'))
    if alts: item['alternateName'] = alts
    graph.append(item)

# --- publications --------------------------------------------------------
venues = {}
places = {}
for p in pubs:
    pid = f'publication/{p["publication_id"]}'
    t = p['pub_type']
    item = {'@id': pid, '@type': TYPE.get(t, 'CreativeWork'), 'identifier': p['publication_id'],
            'name': p['title'], 'datePublished': str(p['year']), 'kg:publicationType': LABEL_TYPE.get(t, t)}
    au = split(p['authors'])
    item['author'] = {'@list': [{'@id': f'person/{slug(a)}'} for a in au]}
    oa = clean(p.get('open_access')) == 'yes'
    # abstracts are published only for open-access publications (decision M. Russo 13/11/2026)
    if oa and clean(p.get('abstract')): item['abstract'] = clean(p['abstract'])
    for k, prop in [('language', 'inLanguage'), ('pages', 'pagination'),
                    ('isbn', 'isbn'), ('issn', 'issn'), ('publisher', 'publisher'), ('full_text_url', 'url'),
                    ('volume', 'volumeNumber'), ('number', 'issueNumber')]:
        v = clean(p.get(k))
        if v: item[prop] = v
    same = []
    if clean(p.get('doi')): same.append('https://doi.org/' + clean(p['doi']))
    if clean(p.get('iris_url')): same.append(clean(p['iris_url']))
    if same: item['sameAs'] = same
    for k, prop in [('scopus_id', 'kg:scopusId'), ('wos_id', 'kg:wosId'), ('open_access', 'kg:openAccess'),
                    ('anvur_scientific_area08', 'kg:anvurScientificArea08'), ('anvur_class_a', 'kg:anvurClassA')]:
        v = clean(p.get(k))
        if v: item[prop] = v
    # venue
    vname, vtype = None, None
    if clean(p.get('journal')): vname, vtype = p['journal'], 'Periodical'
    elif clean(p.get('book_title')): vname, vtype = p['book_title'], 'Book'
    elif clean(p.get('conference_name')): vname, vtype = p['conference_name'], 'Event'
    if vname:
        vid = f'venue/{slug(vname)[:80]}'
        if vid not in venues:
            venues[vid] = {'@id': vid, '@type': vtype, 'name': vname}
            if vtype == 'Book' and clean(p.get('editors')): venues[vid]['editor'] = clean(p['editors'])
        item['isPartOf'] = {'@id': vid}
        vnode(vid, vname, 'venue', vtype=vtype)
        links.append(dict(source=pid, target=vid, rel='published in'))
    if clean(p.get('conference_name')):
        item['kg:conference'] = {k: clean(p.get(c)) for k, c in
                                 [('name', 'conference_name'), ('location', 'conference_location'), ('year', 'conference_year')] if clean(p.get(c))}
    # keywords / scale / places
    kws = split(p['keywords_normalized'])
    item['about'] = [{'@id': kw_index[k]} for k in kws if k in kw_index]
    sc = split(p.get('case_study_scale'))
    if sc: item['kg:caseStudyScale'] = [{'@id': f'scale/{slug(s)}'} for s in sc]
    pl = split(p.get('case_study_place'), '|')
    if pl:
        item['contentLocation'] = []
        for x in pl:
            plid = f'place/{slug(x)}'
            places[plid] = {'@id': plid, '@type': 'Place', 'name': x}
            item['contentLocation'].append({'@id': plid})
            vnode(plid, x, 'place')
            links.append(dict(source=pid, target=plid, rel='case study in'))
    graph.append(item)
    # viewer
    vnode(pid, p['title'], 'publication', year=p['year'], ptype=LABEL_TYPE.get(t, t),
          authors=au, venue=vname, pages=clean(p.get('pages')), doi=clean(p.get('doi')),
          url=clean(p.get('full_text_url')), iris=clean(p.get('iris_url')), abstract=clean(p.get('abstract')) if oa else None, oa=oa, lang=clean(p.get('language')),
          keywords=kws, scale=sc, places=pl, classA=clean(p.get('anvur_class_a')) == 'yes')
    for a in au:
        aid = f'person/{slug(a)}'
        vnode(aid, a, 'author')
        links.append(dict(source=pid, target=aid, rel='written by'))
    for k in kws:
        if k in kw_index:
            vnode(kw_index[k], k, 'keyword')
            links.append(dict(source=pid, target=kw_index[k], rel='about'))

for k in kw_v:
    kid = kw_index[k['canonical_keyword']]
    if kid in nodes and k.get('macro_theme'):
        nodes[kid]['theme'] = k['macro_theme']
        links.append(dict(source=kid, target=f'theme/{slug(k["macro_theme"])}', rel='broader'))

graph += list(venues.values()) + list(places.values())

# --- dataset description -------------------------------------------------
dataset = {'@id': 'dataset', '@type': 'Dataset',
           'name': 'Knowledge Graph of the publications of Michele Russo',
           'creator': {'@id': 'person/russo-michele'},
           'description': 'Publications (journal articles, book chapters, conference papers, monographs) with authors, keywords, macro-themes, venues, case-study places and scales. Abstracts are included for open-access publications only.',
           'license': 'https://creativecommons.org/licenses/by/4.0/'}
context = {'@vocab': 'https://schema.org/', 'skos': 'http://www.w3.org/2004/02/skos/core#',
           'kg': BASE + 'vocab#', '@base': BASE}
doc = {'@context': context, '@graph': [dataset] + graph}
(OUT / 'kg.jsonld').write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding='utf-8')

# Turtle (optional, needs rdflib)
try:
    import rdflib
    g = rdflib.Graph().parse(data=json.dumps(doc), format='json-ld')
    g.serialize(OUT / 'kg.ttl', format='turtle')
    ntriples = len(g)
except Exception as e:
    ntriples = f'skipped ({e})'

# viewer data
viewer = {'nodes': list(nodes.values()), 'links': links}
(OUT / 'kg_viewer_data.json').write_text(json.dumps(viewer, ensure_ascii=False), encoding='utf-8')
print(f'publications={len(pubs)} nodes={len(nodes)} links={len(links)} triples={ntriples}')

# web viewer: viewer_template.html (same folder as this script) + data -> index.html
tpl = Path(__file__).with_name('viewer_template.html')
if tpl.exists():
    html = tpl.read_text(encoding='utf-8').replace('/*__DATA__*/null', json.dumps(viewer, ensure_ascii=False).replace('</', '<\\/'))
    (OUT / 'index.html').write_text(html, encoding='utf-8')
    print('index.html written')
