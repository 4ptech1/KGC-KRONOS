# train/: the code that will tune the model that writes VCL (D124)

Code only. Nothing here downloads, loads or runs a model, and no training starts before the new PC arrives in March 2027.
The prepared datasets, the checkpoints and the weights never go into the repo; they live on the new PC's 300 GB, in
`KRONOS\train` and `KRONOS\models` (D91). The plan is `docs/ai-plan.md`, "Stage C in detail".

| File | What it does | State |
| --- | --- | --- |
| `prompts.py` | The short fixed prompts the model is trained on and run with: write a recipe, repair a recipe | done |
| `make_dataset.py` | `python train/make_dataset.py`: the chat dataset (`train.jsonl`, `val.jsonl`, `manifest.json`) from `data/verified` and `vclc datagen`, into `dist/train/` (git-ignored) | done |
| grammar conversion | `vcl/grammar.lark` into the offline runner's grammar, so the model can write only text that parses as VCL | next |
| QLoRA script | the tuning run, for March | later |
| 4-bit export | the tuned model stored small, for March | later |

**What is in the dataset.** `write` examples: a request (the original and each of its 3 paraphrases) to the proven
recipe. `repair` examples: the request, a recipe that did not pass and the toolchain's messages, to the proven recipe;
every real failed try is kept and at most 10 mutants per pair (`--repairs-per-pair`). The contrast examples are not
training data: they score the Codex meaning check.

**What is never in it.** The held-out pairs in `data/eval` (the export stops if one turns up), the real KGC samples, and
anything not made by this repo's tools. The split into train and validation is by pair, so validation is never a recipe
the model saw in another wording.

`python -m unittest tests.test_train_dataset` runs its tests.
