# data/repair: failing specs and their fixes

Pairs of a spec that fails a check (with the coded messages or the failing TEST) and the spec that fixes
it. Two sources, both made by the toolchain (Phase 7 of `docs/vep-0.6-stage-a.md`): mutants of the proven
pairs, and the real failed attempts of the corpus helpers. `python compiler/vclc.py datagen` builds the pairs
(`compiler/vcl_datagen.py`, D103) fresh into `dist/datagen/` (git-ignored, never stored here): `repairs.jsonl` (the
broken spec, the messages KRONOS would send back, the proven spec), `contrast.jsonl` (a spec one step away that
passes its own TESTs yet ends differently) and `summary.json`. Only data/verified is used, never data/eval.
`attempts/<id>/` keeps the helpers' earlier tries (`tryN.vcl` with the output it got in `tryN.txt`; 40
stored pairs so far, D65, D66, D82, D85; some tries passed every gate but were weak, so read each `tryN.txt`).
`attempts/c074/` (nine tries, batches 4 to 6), `c094/`, `c098/`, `c106/`, `c119/` and `c120/` are failed tries with no fix yet, and `c107/` tries that
passed every gate but were one TASK for a build request; `c137/`, `c148/` and `c152/` (batch 7) too: none of them is stored (D85, D93, D95,
D96); c074, c098, c106 and c120 are retired. The KRONOS app keeps its own attempts in its data folder, outside the repo
(D69); on the new PC, from March 2027, that is `KRONOS\work\runs` (D91).
