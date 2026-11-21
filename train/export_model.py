"""Merge a tuned adapter into its base model and store the result in 4-bit for the offline runner (D128).

    python train/export_model.py --base KRONOS\\models\\base\\<model> --adapter KRONOS\\train\\runs\\<run>\\adapter \\
        --out KRONOS\\models\\<name> --llama-cpp KRONOS\\mirror\\llama.cpp --base-name <name> --base-version <v> \\
        --licence Apache-2.0

CODE ONLY BEFORE MARCH 2027, never run: no model is downloaded, loaded or converted before the new PC arrives. Without
`--go` it prints the plan (every command it would run) and checks what it can without a model, so it is safe anywhere.

The steps, each its own command so a failure says which:
1. merge     the LoRA adapter into the base model, in 16 bits (merging into 4-bit weights would lose accuracy), to
             <out>/work/merged
2. convert   llama.cpp's convert_hf_to_gguf.py makes a 16-bit GGUF file
3. quantize  llama.cpp's llama-quantize stores it in 4 bits (Q4_K_M: about a quarter of the size, about 4.5 GB for a
             7-billion-parameter model)
4. record    <out>/<name>-<quant>.gguf, <out>/vcl.gbnf (the grammar the runner locks the output with) and
             <out>/model.json, which D102 asks for: the base model's name, version, licence, model card and SHA-256,
             the adapter's run record, the dataset's SHA-256, the prompt's, the grammar's and the model file's

Guards: the output is never inside the repo (weights never go into the repo, D91); the drive keeps its 30 GB reserve; the
adapter must have been tuned from this base model; a model tuned on real data (`--real-data`) is marked in model.json
as never leaving the PC (D91). llama.cpp (MIT licence) is built on the new PC from its own source, from the offline
mirror, with IT's approval; it is not part of this repo.
"""
import argparse
import dataclasses
import datetime as dt
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "compiler"))
from make_dataset import digest  # noqa: E402
from prompts import SYSTEM  # noqa: E402
from qlora_train import (KEEP_FREE_GB, TrainError, folder_gb, free_gb, refuse_inside_repo,  # noqa: E402
                         sha256_file)

QUANTS = ("Q4_K_M", "Q4_K_S", "Q5_K_M", "Q5_K_S", "Q6_K", "Q8_0")
EXPORT_FACTOR = 2.5   # the work folder holds a 16-bit merged copy and a 16-bit GGUF, about 2.5 x the base model


@dataclasses.dataclass
class Export:
    base: str
    adapter: str
    out: str
    llama_cpp: str
    base_name: str
    base_version: str
    licence: str
    quant: str = "Q4_K_M"
    model_card: str = ""           # the base model's card (README.md in its folder) unless another file is given
    real_data: bool = False        # tuned on real data: model.json says it never leaves this PC (D91)

    @property
    def name(self):
        return Path(self.out).name

    @property
    def work(self):
        return Path(self.out) / "work"

    @property
    def merged(self):
        return self.work / "merged"

    @property
    def f16(self):
        return self.work / f"{self.name}-f16.gguf"

    @property
    def gguf(self):
        return Path(self.out) / f"{self.name}-{self.quant}.gguf"

    @property
    def convert_script(self):
        return Path(self.llama_cpp) / "convert_hf_to_gguf.py"

    @property
    def quantize_bin(self):
        return Path(self.llama_cpp) / "build" / "bin" / ("llama-quantize.exe" if os.name == "nt" else "llama-quantize")


def steps(cfg):
    """The commands, in order: {name, argv}. argv is a list (never a shell string), or None for the merge, which is Python."""
    return [
        {"name": "merge", "argv": None},
        {"name": "convert", "argv": [sys.executable, str(cfg.convert_script), str(cfg.merged), "--outfile",
                                     str(cfg.f16), "--outtype", "f16"]},
        {"name": "quantize", "argv": [str(cfg.quantize_bin), str(cfg.f16), str(cfg.gguf), cfg.quant]},
    ]


def try_command(cfg):
    """How to try the model with the grammar, to print (the flag names are llama.cpp's and move between versions)."""
    return ["llama-cli", "-m", str(cfg.gguf), "--grammar-file", str(Path(cfg.out) / "vcl.gbnf"), "-p", "<a request>"]


def check(cfg):
    """The problems that would stop an export, found without loading a model."""
    problems = []
    try:
        refuse_inside_repo(cfg.out)
    except TrainError as err:
        problems.append(str(err))
    if cfg.quant not in QUANTS:
        problems.append(f"--quant {cfg.quant} is not one of {', '.join(QUANTS)}")
    for what, path in (("the base model", cfg.base), ("the adapter", cfg.adapter)):
        if not Path(path).is_dir():
            problems.append(f"{what} folder {path} does not exist")
    if Path(cfg.adapter).is_dir():
        if not (Path(cfg.adapter) / "adapter_config.json").is_file():
            problems.append(f"{cfg.adapter} has no adapter_config.json: it is not a LoRA adapter")
        record = Path(cfg.adapter).parent / "run.json"
        if not record.is_file():
            problems.append(f"{record} is missing: the adapter's run record (train/qlora_train.py writes it)")
        else:
            tuned_from = json.loads(record.read_text(encoding="utf-8"))["config"]["base"]
            if Path(tuned_from).resolve() != Path(cfg.base).resolve():
                problems.append(f"this adapter was tuned from {tuned_from}, not from {cfg.base}")
    for what, path in (("convert_hf_to_gguf.py", cfg.convert_script), ("llama-quantize", cfg.quantize_bin)):
        if not path.is_file():
            problems.append(f"{what} is not at {path}: llama.cpp is built from the offline mirror on the new PC")
    if not (cfg.licence or "").strip() or not cfg.base_name.strip() or not cfg.base_version.strip():
        problems.append("the base model's name, version and licence must be given: D102 writes them beside the model")
    need = folder_gb(cfg.base) * EXPORT_FACTOR
    free = free_gb(cfg.out)
    if free - need < KEEP_FREE_GB:
        problems.append(f"the export needs about {need:.0f} GB and the drive has {free:.0f} GB free; "
                        f"{KEEP_FREE_GB} GB must stay free")
    return problems


def base_files(base):
    """{relative path: SHA-256} of every file in the base model's folder: its weights, tokenizer, model card and
    licence (hidden folders such as the download cache are left out)."""
    out = {}
    for f in sorted(Path(base).rglob("*")):
        rel = f.relative_to(base)
        if f.is_file() and not any(part.startswith(".") for part in rel.parts):
            out[rel.as_posix()] = sha256_file(f)
    return out


def model_record(cfg, run, files, gguf_sha, grammar_sha, today=None):
    """model.json: what D102 asks to be written beside the model, plus what traces it to its data."""
    return {
        "name": cfg.name,
        "date": (today or dt.date.today()).isoformat(),
        "base_model": {"name": cfg.base_name, "version": cfg.base_version, "licence": cfg.licence,
                       "model_card": cfg.model_card or "README.md in the base folder", "files_sha256": files},
        "tuned_with": {"method": "QLoRA", "run": {"config": run["config"], "examples": run["examples"],
                                                    "losses": {"eval_loss": run["losses"].get("eval_loss")},
                                                    "libraries": run["libraries"], "smoke": run.get("smoke", False)}},
        "dataset": run["dataset"],
        "prompt_sha256": digest(SYSTEM),
        "grammar_sha256": grammar_sha,
        "file": {"name": cfg.gguf.name, "quantisation": cfg.quant, "sha256": gguf_sha,
                 "bytes": cfg.gguf.stat().st_size if cfg.gguf.is_file() else None},
        "trained_on_real_data": cfg.real_data,
        "note": ("Tuned on real KGC data: this model never leaves this PC (D91)." if cfg.real_data else
                 "Tuned on the repo's invented data only."),
    }


def merge_adapter(cfg):
    """The adapter merged into the 16-bit base model, saved to cfg.merged. Needs the graphics libraries."""
    lacking = [n for n in ("torch", "transformers", "peft") if importlib.util.find_spec(n) is None]
    if lacking:
        raise TrainError(f"not installed: {', '.join(lacking)}. This runs on the new PC (train/requirements-train.txt)")
    import torch
    from peft import PeftModel
    from transformers import AutoModelForCausalLM, AutoTokenizer

    dtype = torch.bfloat16 if torch.cuda.is_available() and torch.cuda.is_bf16_supported() else torch.float16
    model = AutoModelForCausalLM.from_pretrained(cfg.base, torch_dtype=dtype)
    model = PeftModel.from_pretrained(model, cfg.adapter).merge_and_unload()
    cfg.merged.mkdir(parents=True, exist_ok=True)
    model.save_pretrained(str(cfg.merged), safe_serialization=True)
    AutoTokenizer.from_pretrained(cfg.adapter).save_pretrained(str(cfg.merged))


def run(cfg, today=None):
    """Do every step, then write the record. Stops at the first step that fails."""
    problems = check(cfg)
    if problems:
        raise TrainError("; ".join(problems))
    Path(cfg.out).mkdir(parents=True, exist_ok=True)
    cfg.work.mkdir(parents=True, exist_ok=True)
    for step in steps(cfg):
        print(f"== {step['name']}")
        if step["argv"] is None:
            merge_adapter(cfg)
            continue
        done = subprocess.run(step["argv"])
        if done.returncode:
            raise TrainError(f"{step['name']} failed (exit {done.returncode}): {' '.join(step['argv'])}")
    if not cfg.gguf.is_file():
        raise TrainError(f"{cfg.gguf} was not written")
    shutil.copyfile(HERE / "vcl.gbnf", Path(cfg.out) / "vcl.gbnf")
    run_record = json.loads((Path(cfg.adapter).parent / "run.json").read_text(encoding="utf-8"))
    record = model_record(cfg, run_record, base_files(cfg.base), sha256_file(cfg.gguf),
                          sha256_file(Path(cfg.out) / "vcl.gbnf"), today)
    (Path(cfg.out) / "model.json").write_text(json.dumps(record, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    shutil.rmtree(cfg.work)   # the 16-bit copies are about 30 GB and can be made again
    return record


def parse_args(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--base", required=True)
    ap.add_argument("--adapter", required=True, help="the run's adapter folder (the run.json is beside it)")
    ap.add_argument("--out", required=True, help="the model's folder under KRONOS\\models, outside the repo")
    ap.add_argument("--llama-cpp", required=True, help="a built llama.cpp folder")
    ap.add_argument("--base-name", required=True)
    ap.add_argument("--base-version", required=True)
    ap.add_argument("--licence", required=True)
    ap.add_argument("--quant", default="Q4_K_M")
    ap.add_argument("--model-card", default="")
    ap.add_argument("--real-data", action="store_true", help="the adapter was tuned on real KGC data")
    ap.add_argument("--go", action="store_true", help="really export (without it: print the plan and check)")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    cfg = Export(base=args.base, adapter=args.adapter, out=args.out, llama_cpp=args.llama_cpp,
                 base_name=args.base_name, base_version=args.base_version, licence=args.licence, quant=args.quant,
                 model_card=args.model_card, real_data=args.real_data)
    problems = check(cfg)
    print("plan:")
    for step in steps(cfg):
        print(f"  {step['name']}: " + (" ".join(step["argv"]) if step["argv"] else "merge the adapter into the base (Python)"))
    print(f"  record: {cfg.gguf}, {Path(cfg.out) / 'vcl.gbnf'}, {Path(cfg.out) / 'model.json'}")
    print("  to try it afterwards:", " ".join(try_command(cfg)))
    for problem in problems:
        print("PROBLEM:", problem)
    if problems:
        return 1
    if not args.go:
        print("dry run: nothing was merged or converted. Add --go to export (on the new PC, March 2027).")
        return 0
    try:
        record = run(cfg)
    except TrainError as err:
        print("STOP:", err)
        return 2
    print(f"done: {cfg.gguf} ({record['file']['bytes']} bytes), sha256 {record['file']['sha256']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
