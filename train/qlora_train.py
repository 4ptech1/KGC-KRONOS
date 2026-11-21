"""QLoRA tuning of a small open model on the VCL dataset (D128; candidate B of D102, plan `docs/ai-plan.md`).

    python train/qlora_train.py --base KRONOS\\models\\base\\<model> --data dist\\train --out KRONOS\\train\\runs\\<name>

CODE ONLY BEFORE MARCH 2027. It has never been run: no model is downloaded, loaded or trained before the new PC
arrives (CLAUDE.md). Without `--go` it only prints the plan and checks everything it can check without a graphics card
or a model, so it is safe anywhere. The first run in March is `--smoke` (24 examples, 10 steps): if that finishes and
writes an adapter, the full run is the same command without `--smoke`.

What it does with `--go`:
1. checks the dataset (`train/make_dataset.py` wrote it): its files match the manifest's SHA-256, its prompt is the
   prompt in `train/prompts.py` today, and no held-out item (`data/eval`) is in it;
2. refuses an output folder inside the repo (weights and checkpoints never go into the repo, D91) and a drive that
   would fall below its 30 GB reserve (`docs/kronos-storage-plan.md`);
3. loads the base model in 4-bit (NF4, double quantisation), adds LoRA adapters to every linear layer and trains them
   with the model's own chat template, the loss on the recipe only (the prompt is masked), never on the prompt;
4. skips, and counts, an example longer than `--max-seq-len` tokens: a recipe cut short would teach the wrong ending;
5. writes `adapter/` (the LoRA weights, small), the tokenizer, and `run.json` (the settings, the dataset's SHA-256, the
   library versions, the date, the losses), which `train/export_model.py` reads next.

Needs (on the new PC only, pinned there with hashes, `train/requirements-train.txt`): torch, transformers, peft,
bitsandbytes, accelerate. They are imported only when a run starts.
"""
import argparse
import dataclasses
import datetime as dt
import importlib.util
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "compiler"))
from make_dataset import digest  # noqa: E402
from prompts import SYSTEM  # noqa: E402

LIBRARIES = ("torch", "transformers", "peft", "bitsandbytes", "accelerate")
CHARS_PER_TOKEN = 2.5     # a deliberately low guess, so the dry run over-counts the long examples; March counts tokens
KEEP_FREE_GB = 30         # the new PC's reserve (D91)
SMOKE_EXAMPLES, SMOKE_STEPS = 24, 10


class TrainError(Exception):
    """Something that must be fixed before a run starts."""


@dataclasses.dataclass
class Config:
    base: str                        # the base model's folder (KRONOS\models\base\<model>)
    data: str                        # the folder make_dataset.py wrote
    out: str                         # this run's folder (KRONOS\train\runs\<name>)
    max_seq_len: int = 8192          # tokens; 24 GB card. Use 4096 on 16 GB (more examples are skipped)
    epochs: float = 3.0
    learning_rate: float = 2e-4
    warmup_ratio: float = 0.03
    batch_size: int = 1
    grad_accum: int = 16             # 16 recipes per step
    lora_r: int = 16
    lora_alpha: int = 32
    lora_dropout: float = 0.05
    target_modules: tuple = ("q_proj", "k_proj", "v_proj", "o_proj", "gate_proj", "up_proj", "down_proj")
    eval_steps: int = 50
    save_steps: int = 50
    seed: int = 42
    load_4bit: bool = True
    smoke: bool = False

    def as_dict(self):
        return dataclasses.asdict(self)


# ---------------------------------------------------------------- what can be checked without a model

def sha256_file(path):
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def load_examples(path):
    out = []
    for number, line in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        row = json.loads(line)
        roles = [m.get("role") for m in row.get("messages", [])]
        if roles != ["system", "user", "assistant"]:
            raise TrainError(f"{path} line {number}: the messages are {roles}, not system, user, assistant")
        out.append(row)
    return out


def check_dataset(data_dir):
    """(train examples, val examples, manifest) when the dataset is whole and current, else TrainError."""
    data = Path(data_dir)
    for name in ("train.jsonl", "val.jsonl", "manifest.json"):
        if not (data / name).is_file():
            raise TrainError(f"{data / name} is missing: run python train/make_dataset.py")
    manifest = json.loads((data / "manifest.json").read_text(encoding="utf-8"))
    for name, want in manifest.get("files", {}).items():
        if sha256_file(data / name) != want:
            raise TrainError(f"{name} is not what the manifest says (its SHA-256 differs): make the dataset again")
    if manifest.get("prompt_sha256") != digest(SYSTEM):
        raise TrainError("the dataset was made with another prompt than train/prompts.py has now: make it again")
    train, val = load_examples(data / "train.jsonl"), load_examples(data / "val.jsonl")
    if not train or not val:
        raise TrainError("the dataset has no training or no validation examples")
    held = held_out_ids()
    for row in train + val:
        if row["id"].split("/")[0] in held:
            raise TrainError(f"{row['id']} is a held-out item (data/eval): it must never be trained on")
    return train, val, manifest


def held_out_ids():
    from vcl_corpus import items
    return {folder.name for folder in items("eval")}


def too_long(examples, max_seq_len):
    """How many examples are probably over max_seq_len tokens, by characters: a guess that errs high."""
    limit = max_seq_len * CHARS_PER_TOKEN
    return sum(1 for e in examples if sum(len(m["content"]) for m in e["messages"]) > limit)


def mask_labels(prompt_ids, full_ids):
    """(input_ids, labels): the loss is on the answer only, so the prompt's labels are -100. The prompt's tokens must
    be the start of the full text's: if the chat template changes them, the answer can't be told from the prompt."""
    if list(full_ids[:len(prompt_ids)]) != list(prompt_ids):
        raise TrainError("the chat template does not give the same tokens for the prompt alone and for the whole "
                         "exchange, so the answer can't be separated; check the model's template")
    return list(full_ids), [-100] * len(prompt_ids) + list(full_ids[len(prompt_ids):])


def encode_example(tokenizer, example, max_seq_len):
    """{input_ids, labels, attention_mask} for one example, or None when it is longer than max_seq_len tokens.
    tokenizer needs apply_chat_template(messages, tokenize=False, add_generation_prompt=...) and tokenizer(text,
    add_special_tokens=False)["input_ids"], which every Hugging Face chat tokenizer has."""
    messages = example["messages"]
    prompt = tokenizer.apply_chat_template(messages[:-1], tokenize=False, add_generation_prompt=True)
    full = tokenizer.apply_chat_template(messages, tokenize=False)
    ids, labels = mask_labels(tokenizer(prompt, add_special_tokens=False)["input_ids"],
                              tokenizer(full, add_special_tokens=False)["input_ids"])
    if len(ids) > max_seq_len:
        return None
    return {"input_ids": ids, "labels": labels, "attention_mask": [1] * len(ids)}


def refuse_inside_repo(path):
    """Weights and checkpoints never go into the repo (D91)."""
    target = Path(path).resolve()
    if target == ROOT or ROOT in target.parents:
        raise TrainError(f"{path} is inside the repo: weights and checkpoints stay on the new PC's drive (D91)")


def free_gb(path):
    folder = Path(path)
    while not folder.exists() and folder != folder.parent:
        folder = folder.parent
    return shutil.disk_usage(folder).free / 1e9


def folder_gb(path):
    return sum(f.stat().st_size for f in Path(path).rglob("*") if f.is_file()) / 1e9 if Path(path).is_dir() else 0.0


def need_gb(base):
    """What a run needs on the drive: checkpoints and caches, about the size of the base model."""
    return max(folder_gb(base), 1.0)


def missing_libraries():
    return [name for name in LIBRARIES if importlib.util.find_spec(name) is None]


def plan(cfg):
    """The plan as lines of text, and the problems that would stop a run. Reads files; loads no model."""
    problems, lines = [], []
    try:
        train, val, manifest = check_dataset(cfg.data)
        lines.append(f"dataset: {len(train)} train and {len(val)} validation examples "
                     f"({manifest['examples']['train']['write']} write and {manifest['examples']['train']['repair']} repair "
                     f"in train); its SHA-256 are in the manifest and match")
        over = too_long(train + val, cfg.max_seq_len)
        lines.append(f"length: about {over} of {len(train) + len(val)} examples are probably over {cfg.max_seq_len} tokens "
                     f"(counted by characters, {CHARS_PER_TOKEN} per token) and would be skipped, not cut; the run counts "
                     "the real tokens")
    except TrainError as err:
        problems.append(str(err))
    try:
        refuse_inside_repo(cfg.out)
    except TrainError as err:
        problems.append(str(err))
    base = Path(cfg.base)
    if not base.is_dir():
        problems.append(f"the base model folder {cfg.base} does not exist (downloaded once, with IT's approval, in March)")
    elif not (base / "config.json").is_file():
        problems.append(f"{cfg.base} has no config.json: it is not a Hugging Face model folder")
    need = need_gb(cfg.base)
    free = free_gb(cfg.out)
    lines.append(f"drive: the run needs about {need:.0f} GB and the drive has {free:.0f} GB free; "
                 f"{KEEP_FREE_GB} GB must stay free")
    if free - need < KEEP_FREE_GB:
        problems.append(f"the drive would fall below its {KEEP_FREE_GB} GB reserve")
    steps = "{} steps (smoke)".format(SMOKE_STEPS) if cfg.smoke else f"{cfg.epochs:g} epochs"
    lines.append(f"settings: LoRA r={cfg.lora_r} alpha={cfg.lora_alpha} dropout={cfg.lora_dropout} on "
                 f"{', '.join(cfg.target_modules)}; learning rate {cfg.learning_rate:g}; batch {cfg.batch_size} x "
                 f"{cfg.grad_accum}; {steps}; {'4-bit NF4' if cfg.load_4bit else 'full precision'}; seed {cfg.seed}")
    lacking = missing_libraries()
    if lacking:
        lines.append(f"libraries not installed here: {', '.join(lacking)} (they are installed on the new PC)")
    return lines, problems


def run_record(cfg, manifest, counts, versions, losses, started, finished):
    """What `run.json` holds, so a model can be traced back to its settings and its data."""
    return {"config": cfg.as_dict(), "dataset": {"files": manifest.get("files", {}), "prompt_sha256": manifest["prompt_sha256"],
                                                 "pairs": manifest["pairs"]},
            "examples": counts, "libraries": versions, "losses": losses, "started": started, "finished": finished,
            "smoke": cfg.smoke}


# ---------------------------------------------------------------- the run (needs the new PC)

def run_training(cfg):
    """Train and write adapter/, the tokenizer and run.json into cfg.out. Needs the graphics libraries."""
    lacking = missing_libraries()
    if lacking:
        raise TrainError(f"not installed: {', '.join(lacking)}. This runs on the new PC (train/requirements-train.txt)")
    import inspect

    import torch
    import transformers
    from peft import LoraConfig, get_peft_model, prepare_model_for_kbit_training
    from transformers import (AutoModelForCausalLM, AutoTokenizer, BitsAndBytesConfig, Trainer, TrainingArguments)

    train_rows, val_rows, manifest = check_dataset(cfg.data)
    if cfg.smoke:
        train_rows, val_rows = train_rows[:SMOKE_EXAMPLES], val_rows[:SMOKE_EXAMPLES // 4]
    out = Path(cfg.out)
    out.mkdir(parents=True, exist_ok=True)
    started = dt.datetime.now().isoformat(timespec="seconds")

    tokenizer = AutoTokenizer.from_pretrained(cfg.base)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token
    encoded = {}
    for name, rows in (("train", train_rows), ("val", val_rows)):
        done = [encode_example(tokenizer, row, cfg.max_seq_len) for row in rows]
        encoded[name] = [e for e in done if e is not None]
        print(f"{name}: {len(encoded[name])} examples, {len(done) - len(encoded[name])} skipped as longer than "
              f"{cfg.max_seq_len} tokens")
    if not encoded["train"] or not encoded["val"]:
        raise TrainError("every example is longer than --max-seq-len: raise it, or use a card with more memory")

    use_bf16 = torch.cuda.is_available() and torch.cuda.is_bf16_supported()
    compute = torch.bfloat16 if use_bf16 else torch.float16
    quant = BitsAndBytesConfig(load_in_4bit=True, bnb_4bit_quant_type="nf4", bnb_4bit_use_double_quant=True,
                               bnb_4bit_compute_dtype=compute) if cfg.load_4bit else None
    model = AutoModelForCausalLM.from_pretrained(cfg.base, quantization_config=quant, torch_dtype=compute,
                                                 device_map={"": 0})
    model.config.use_cache = False
    if cfg.load_4bit:
        model = prepare_model_for_kbit_training(model, use_gradient_checkpointing=True)
    else:   # a 16-bit model with gradient checkpointing needs this, which the 4-bit helper does
        model.enable_input_require_grads()
    model = get_peft_model(model, LoraConfig(r=cfg.lora_r, lora_alpha=cfg.lora_alpha, lora_dropout=cfg.lora_dropout,
                                             target_modules=list(cfg.target_modules), bias="none", task_type="CAUSAL_LM"))
    model.print_trainable_parameters()

    def collate(batch):
        width = max(len(b["input_ids"]) for b in batch)
        pad = tokenizer.pad_token_id
        return {"input_ids": torch.tensor([b["input_ids"] + [pad] * (width - len(b["input_ids"])) for b in batch]),
                "labels": torch.tensor([b["labels"] + [-100] * (width - len(b["labels"])) for b in batch]),
                "attention_mask": torch.tensor([b["attention_mask"] + [0] * (width - len(b["attention_mask"]))
                                                for b in batch])}

    params = inspect.signature(TrainingArguments).parameters
    eval_key = "eval_strategy" if "eval_strategy" in params else "evaluation_strategy"   # renamed between versions
    args = TrainingArguments(
        output_dir=str(out / "checkpoints"), per_device_train_batch_size=cfg.batch_size,
        per_device_eval_batch_size=cfg.batch_size, gradient_accumulation_steps=cfg.grad_accum,
        learning_rate=cfg.learning_rate, lr_scheduler_type="cosine", warmup_ratio=cfg.warmup_ratio,
        num_train_epochs=cfg.epochs, max_steps=SMOKE_STEPS if cfg.smoke else -1, bf16=use_bf16, fp16=not use_bf16,
        optim="paged_adamw_8bit" if cfg.load_4bit else "adamw_torch", logging_steps=5, save_steps=cfg.save_steps,
        eval_steps=cfg.eval_steps, save_total_limit=3, gradient_checkpointing=True,
        gradient_checkpointing_kwargs={"use_reentrant": False}, seed=cfg.seed, data_seed=cfg.seed, report_to="none",
        remove_unused_columns=False, **{eval_key: "steps"})
    trainer = Trainer(model=model, args=args, train_dataset=encoded["train"], eval_dataset=encoded["val"],
                      data_collator=collate)
    trainer.train()
    final = trainer.evaluate()

    model.save_pretrained(str(out / "adapter"))
    tokenizer.save_pretrained(str(out / "adapter"))
    losses = {"eval_loss": final.get("eval_loss"),
              "train_log": [{k: v for k, v in entry.items() if k in ("step", "loss", "eval_loss")}
                            for entry in trainer.state.log_history]}
    versions = {"torch": torch.__version__, "transformers": transformers.__version__}
    for name in ("peft", "bitsandbytes", "accelerate"):
        versions[name] = importlib.import_module(name).__version__
    record = run_record(cfg, manifest, {"train": len(encoded["train"]), "val": len(encoded["val"])}, versions, losses,
                        started, dt.datetime.now().isoformat(timespec="seconds"))
    (out / "run.json").write_text(json.dumps(record, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    print(f"done: {out / 'adapter'} and {out / 'run.json'}; eval loss {final.get('eval_loss')}")
    return record


def parse_args(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--base", required=True, help="the base model's folder")
    ap.add_argument("--data", default=str(ROOT / "dist" / "train"), help="the folder make_dataset.py wrote")
    ap.add_argument("--out", required=True, help="this run's folder, outside the repo")
    ap.add_argument("--go", action="store_true", help="really train (without it: print the plan and check)")
    ap.add_argument("--smoke", action="store_true", help=f"{SMOKE_EXAMPLES} examples and {SMOKE_STEPS} steps, to see that it works")
    ap.add_argument("--no-4bit", action="store_true", help="load the model in 16 bits (a small model on a big card)")
    defaults = Config("", "", "")
    for field in dataclasses.fields(Config):
        if field.name in ("base", "data", "out", "smoke", "load_4bit", "target_modules"):
            continue
        kind = type(getattr(defaults, field.name))
        ap.add_argument("--" + field.name.replace("_", "-"), type=kind, default=getattr(defaults, field.name))
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    values = {f.name: getattr(args, f.name) for f in dataclasses.fields(Config)
              if hasattr(args, f.name) and f.name not in ("load_4bit", "smoke")}
    cfg = Config(**values, load_4bit=not args.no_4bit, smoke=args.smoke)
    lines, problems = plan(cfg)
    print("\n".join(lines))
    for problem in problems:
        print("PROBLEM:", problem)
    if problems:
        return 1
    if not args.go:
        print("dry run: nothing was loaded or trained. Add --go to train (on the new PC, March 2027).")
        return 0
    try:
        run_training(cfg)
    except TrainError as err:
        print("STOP:", err)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
