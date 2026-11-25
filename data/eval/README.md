# data/eval: the held-out benchmark

About 10% of the corpus (15 of the 150 tasks), chosen when the corpus is made and **never used for
training, paraphrasing or repair**. Same layout as `data/verified/`, plus a contract and hidden reference
TESTs for `vclc bench` (Phase 8 of `docs/vep-0.6-stage-a.md`). It holds 15 pairs (c022 to c149: 12 TASKs and 3 BUILDs over
the families and registers). `vclc bench` (Phase 8, D136) takes each item's contract (a TASK's PROC, KIND, ARGS and
RET; a BUILD's NAME and modMain.RunHeadless SIG) from its recipe, and its FIXTURE and TEST lines are the hidden
reference TESTs; `vclc bench --prompts DIR` writes what a writer sees: the request and the contract only. A pair proven through the KRONOS
app is never put here (D70).
