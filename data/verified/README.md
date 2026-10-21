# data/verified: the proven pairs

One folder per pair, `c001`, `c002`, ...: a plain-English request and the VCL spec that the whole
toolchain proved (D64, `compiler/vcl_corpus.py`).

| File | What |
| --- | --- |
| `task.txt` | the request, in the words Boss would type |
| `spec.vcl` | the proven spec: VCL 0.6, canonical (`vclc fmt`), with its GOAL lines |
| `meta.json` | id, kind (TASK or BUILD), registers, family, ops used, steps, source, the mutation score it was stored with |
| `paraphrase_N.txt` | the same request worded differently (optional) |

The VBA is not stored: it is derived and goes stale. `python compiler/vclc.py export` compiles it fresh
into `dist/dataset.jsonl`.

- Check every pair: `python compiler/vclc.py corpus` (`--mutation` runs the mutation again, `--report` shows coverage).
- Store a new pair: `python compiler/vclc.py corpus-add --id c013 --request FILE --spec FILE --registers bom,grn --family reconcile`.
  It stores nothing unless every gate passes.
- `c001` to `c012` are the 12 specs in `specs/` (with a GOAL line added); the plan for the rest is `docs/corpus-plan.md`.

Synthetic data only: invented sheets, headers and figures, never a real KGC file or name (CLAUDE.md, D25).
