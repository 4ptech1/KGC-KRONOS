# Benchmark results (`vclc bench`, D136)

One folder per run: the candidates as the writer gave them (`try1/<id>.vcl`, the first try; `final/<id>.vcl`, after
KRONOS's up to 3 tries with the toolchain's own messages, never the hidden TESTs) and their scores (`try1.json`,
`final.json`). Score a folder again with `python compiler/vclc.py bench <folder>/final --today <the run's date>`.

**These recipes answer held-out requests: nothing here is ever trained on, paraphrased, repaired or put in the recipe
finder's library** (D59; `train/make_dataset.py`, `vclc datagen` and `kronos_finder.py` read only `data/verified`).

| Run | Writer | First try: right (H) | After 3 tries: right (H) | Right and strong (strict) |
| --- | --- | --- | --- | --- |
| `2026-11-25-claude-baseline` | Claude (Sonnet helpers) as the Orchestrator, VCL 0.7 prompt | 9 of 15 | 10 of 15 | 4 of 15 |
| `2026-11-26-claude-documents` | the same, on the 5 KGC document items added by batch 9 (D140), with the layouts KRONOS attaches | 0 of 5 | 3 of 5 | 2 of 5 |

`2026-11-26-claude-documents` (D141) scores only c162, c168, c173, c177 and c181 (its json files list the other 15 as
"no candidate"). Its first run showed that the requests of c162 and c168 did not say what their hidden TESTs check (the
count returned and blank cells kept; a row added on every run); each request gained one sentence, and a fresh writer
wrote those two again. The first run's recipes for them are kept in `unclear-wording/` (0 of 2 right).
