# data/repair: failing specs and their fixes

Pairs of a spec that fails a check (with the coded messages or the failing TEST) and the spec that fixes
it. Two sources, both made by the toolchain (Phase 7 of `docs/vep-0.6-stage-a.md`): mutants of the proven
pairs, and the real failed attempts of the corpus helpers. Phase 7 builds the pairs. Until then
`attempts/<id>/` keeps the helpers' failed tries (`tryN.vcl` with the output it got in `tryN.txt`; 19
pairs so far, D65, D66). The KRONOS app keeps its own attempts in its data folder, outside the repo
(D69).
