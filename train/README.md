# train/: the code that will tune the model that writes VCL (D124)

Code only. Nothing here downloads, loads or runs a model, and no training starts before the new PC arrives in March 2027.
The prepared datasets, the checkpoints and the weights never go into the repo; they live on the new PC's 300 GB, in
`KRONOS\train` and `KRONOS\models` (D91). The plan is `docs/ai-plan.md`, "Stage C in detail".

| File | What it does | State |
| --- | --- | --- |
| `prompts.py` | The short fixed prompts the model is trained on and run with: write a recipe, repair a recipe | done |
| `make_dataset.py` | `python train/make_dataset.py`: the chat dataset (`train.jsonl`, `val.jsonl`, `manifest.json`) from `data/verified` and `vclc datagen`, into `dist/train/` (git-ignored) | done |
| `grammar_gbnf.py`, `vcl.gbnf` | `python train/grammar_gbnf.py`: `vcl/grammar.lark` and the parser's key order as a GBNF grammar for llama.cpp, so the model can only write text shaped like a VCL recipe; `vcl.gbnf` is committed and a test keeps it current | done (D127) |
| `gbnf_check.py` | An independent reader of that grammar, so it is tested without a runner or a model | done (D127) |
| `qlora_train.py` | The QLoRA tuning run: checks the dataset, masks the prompt, trains LoRA adapters on the 4-bit model, writes `adapter/` and `run.json`. Dry run unless `--go` | written, never run (D128) |
| `export_model.py` | Merges the adapter, converts to GGUF, stores it in 4 bits, writes `model.json` (base model, licence, card, SHA-256). Dry run unless `--go` | written, never run (D128) |
| `requirements-train.txt` | The libraries the last two need, on the new PC only | done |

**The first things to do in March**, on the new PC: install the libraries and llama.cpp from the offline mirror, download
a base model with IT's approval into `KRONOS\models\base`, run `python train/qlora_train.py ... --smoke --go` (24
examples, 10 steps: it proves the script and the card work), then the full run, then `export_model.py`, then score the
model on the benchmark beside candidate A (D102).

**The grammar.** The model's reply is generated with `--grammar-file train/vcl.gbnf`. It is the canonical form
`vclc fmt` writes (keys in column 9, steps indented two spaces a level), with keys in order, balanced brackets and a body
under every block opener, at most 10 levels deep. It accepts all 154 VCL 0.6 and 0.7 recipes in the repo and gives the
same verdict as the real parser on about 98% of 23,000 mutated lines; where they differ, the grammar allows only one space
between tokens, and it will accept two words with no space between them, which the validator catches. It lists no op
names or option names: the validator's catalogue does.

**What is in the dataset.** `write` examples: a request (the original and each of its 3 paraphrases) to the proven
recipe. `repair` examples: the request, a recipe that did not pass and the toolchain's messages, to the proven recipe;
every real failed try is kept and at most 10 mutants per pair (`--repairs-per-pair`). The contrast examples are not
training data: they score the Codex meaning check.

**What is never in it.** The held-out pairs in `data/eval` (the export stops if one turns up), the real KGC samples, and
anything not made by this repo's tools. The split into train and validation is by pair, so validation is never a recipe
the model saw in another wording.

`python -m unittest tests.test_train_dataset` runs its tests.
