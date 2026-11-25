# Benchmark results (`vclc bench`, D136)

One folder per run: the candidates as the writer gave them (`try1/<id>.vcl`, the first try; `final/<id>.vcl`, after
KRONOS's up to 3 tries with the toolchain's own messages, never the hidden TESTs) and their scores (`try1.json`,
`final.json`). Score a folder again with `python compiler/vclc.py bench <folder>/final --today <the run's date>`.

**These recipes answer held-out requests: nothing here is ever trained on, paraphrased, repaired or put in the recipe
finder's library** (D59; `train/make_dataset.py`, `vclc datagen` and `kronos_finder.py` read only `data/verified`).

| Run | Writer | First try: right (H) | After 3 tries: right (H) | Right and strong (strict) |
| --- | --- | --- | --- | --- |
| `2026-11-25-claude-baseline` | Claude (Sonnet helpers) as the Orchestrator, VCL 0.7 prompt | 9 of 15 | 10 of 15 | 4 of 15 |
