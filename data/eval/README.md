# data/eval: the held-out benchmark

About 10% of the corpus (15 of the 150 tasks), chosen when the corpus is made and **never used for
training, paraphrasing or repair**. Same layout as `data/verified/`, plus a contract and hidden reference
TESTs for `vclc bench` (Phase 8 of `docs/vep-0.6-stage-a.md`). It holds 4 pairs so far (c022, c030,
c036, c045, D65, D66); the contracts and `vclc bench` come in Phase 8. A pair proven through the KRONOS
app is never put here (D70).
