# Knowledge Graph of the publications of Michele Russo

Interactive knowledge graph of the scientific publications of Michele Russo
(Sapienza Università di Roma, Department of History, Representation and Restoration of Architecture — sector CEAR-10/A Disegno, 08/E1).

- **Interactive viewer:** https://micheler77.github.io/publications-kg/
- **Linked data:** [`kg.jsonld`](kg.jsonld) (JSON-LD) and [`kg.ttl`](kg.ttl) (Turtle), using schema.org, SKOS and Dublin Core-compatible terms.

## Contents
Journal articles, book chapters, conference papers and monographs, with authors, keywords (controlled vocabulary grouped in macro-themes), venues, case-study places and scales, DOI and IRIS Sapienza links.
Every relation in the graph is verified against the publications or the IRIS institutional repository,
except the **textual affinities**, which are inferred and labelled as such (see below).
Abstracts are included only for open-access publications.

## Textual affinities (inferred)
Pairs of publications whose title + abstract are close in content, computed offline with the multilingual
model *multilingual-e5-small* (int8 ONNX, pinned by SHA-256) and cosine similarity (threshold 0.25),
after removing the language component (Italian/English). Only publication ids and scores are published
([`affinity.json`](affinity.json)); abstract texts and vectors are not. A pair is kept when it shares at
least one keyword (basis `shared keywords`) or when the author reviewed and approved it (basis `reviewed`).
In the viewer they appear as a dashed layer, off by default, and under "Connessioni possibili" in each publication card.

## Update
The data are maintained in an Excel workbook (not published here) and regenerated with
`python scripts/build_affinity.py <excel.xlsx> <model_dir> <affinity_decisions.xlsx> <work_dir>` (all pairs recomputed; new pairs without shared keywords go to `affinity_pending.xlsx` for review), then
`python scripts/build_kg.py <excel.xlsx> <output_dir> https://micheler77.github.io/publications-kg/ affinity.json`.

## Licence
Data: [CC BY 4.0](LICENSE).
