# data/repair: failing specs and their fixes

Pairs of a spec that fails a check (with the coded messages or the failing TEST) and the spec that fixes
it. Two sources, both made by the toolchain (Phase 7 of `docs/vep-0.6-stage-a.md`): mutants of the proven
pairs, and the real failed attempts of the corpus helpers. Phase 7 builds the pairs. Until then
`attempts/<id>/` keeps the helpers' earlier tries (`tryN.vcl` with the output it got in `tryN.txt`; 40
stored pairs so far, D65, D66, D82, D85; some tries passed every gate but were weak, so read each `tryN.txt`).
`attempts/c074/` (nine tries, batches 4 to 6), `c094/`, `c098/`, `c106/`, `c119/` and `c120/` are failed tries with no fix yet, and `c107/` tries that
passed every gate but were one TASK for a build request: none of them is stored (D85, D93, D95); c074, c098 and c106 are retired. The KRONOS app keeps its own attempts in its data folder, outside the repo
(D69); on the new PC, from March 2027, that is `KRONOS\work\runs` (D91).
