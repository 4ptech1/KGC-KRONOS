"""The training dataset for the model that writes VCL (D124; candidate B and candidate A of D102 train on the same text).

    python train/make_dataset.py [-o DIR] [--datagen DIR] [--today YYYY-MM-DD] [--repairs-per-pair N]

Writes, into DIR (default dist/train, git-ignored, made fresh and never stored):
    train.jsonl, val.jsonl   one chat example per line: {"id", "type", "source", "messages": [system, user, assistant]}
    manifest.json            counts, the split, the caps, the prompt's SHA-256, lengths, the guards that ran

Two kinds of example, both made from data/verified by this repo's own tools:
    write    the request (and each of its paraphrases) -> the proven recipe
    repair   the request, a recipe that did not pass, the toolchain's messages -> the proven recipe (`vclc datagen`)
The contrast examples (`contrast.jsonl`) are not training data: they are for scoring the meaning check (D124).

The rules, each checked here and in tests/test_train_dataset.py:
- Only data/verified is read. The held-out pairs in data/eval are never used, and the export stops if one of them (its
  id, its request or its recipe) turns up anywhere in the examples (D59).
- The split is by pair, from a hash of the id: a pair's recipe, its paraphrases and its repairs are all in train or all in
  val, so validation is never a recipe the model has already seen in another wording.
- Everything is deterministic: the same corpus, date and caps give the same files, byte for byte.
- Standard library only, and nothing here loads, downloads or runs a model.
"""
import argparse
import datetime as dt
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "compiler"))
from prompts import SYSTEM, repair_prompt, write_prompt  # noqa: E402
from vcl_corpus import load  # noqa: E402

VAL_ONE_IN = 10          # a pair goes to validation when the hash of its id is divisible by this
REPAIRS_PER_PAIR = 10    # mutant repairs kept per pair (the helpers' real failed tries are all kept)


class LeakError(Exception):
    """A held-out item (or a duplicate of one) is in the examples."""


def digest(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def in_validation(pair_id):
    return int(digest(pair_id)[:8], 16) % VAL_ONE_IN == 0


def write_examples(pair):
    """The writing examples of one proven pair: its request and each paraphrase, to the recipe."""
    out = []
    for k, request in enumerate([pair["request"]] + pair["paraphrases"]):
        out.append({"id": f"{pair['id']}/write{k}", "type": "write", "source": pair.get("source", ""),
                    "messages": write_prompt(request) + [{"role": "assistant", "content": pair["spec"]}]})
    return out


def keep(row):
    """Which mutant repairs are kept: a fixed order from a hash, so the choice never depends on the run."""
    return digest(f"{row['id']}|{row['change']}|{row['broken']}")


def repair_examples(pair_id, rows, per_pair):
    """The repair examples of one pair: every failed try a person or helper really made, and at most per_pair mutants."""
    tries = [r for r in rows if r["source"] == "attempt"]
    mutants = sorted((r for r in rows if r["source"] == "mutant"), key=keep)[:per_pair]
    out = []
    for row in sorted(tries, key=lambda r: r["change"]) + sorted(mutants, key=lambda r: r["change"]):
        out.append({"id": f"{pair_id}/repair/{row['source']}/{row['change']}", "type": "repair", "source": row["source"],
                    "messages": repair_prompt(row["request"], row["broken"], row["feedback"]) +
                    [{"role": "assistant", "content": row["fixed"]}]})
    return out


def check_no_leaks(examples, held):
    """Stop if a held-out item, or a copy of one, is in an example (D59)."""
    texts = [(e["id"], m["content"]) for e in examples for m in e["messages"]]
    for item in held:
        for what, text in (("id", item["id"]), ("request", item["request"]), ("recipe", item["spec"].strip())):
            if not text:
                continue
            if what == "id":
                clash = [i for i, _ in texts if i.startswith(item["id"] + "/")]
            else:
                clash = [i for i, t in texts if text in t]
            if clash:
                raise LeakError(f"held-out {item['id']} ({what}) is in the examples: {clash[0]}")


def lengths(examples):
    sizes = [sum(len(m["content"]) for m in e["messages"]) for e in examples]
    return {"examples": len(sizes), "max_chars": max(sizes, default=0),
            "mean_chars": round(sum(sizes) / len(sizes)) if sizes else 0}


def read_repairs(path):
    by_pair = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if line.strip():
            row = json.loads(line)
            by_pair.setdefault(row["id"], []).append(row)
    return by_pair


def build(repairs_path, per_pair=REPAIRS_PER_PAIR):
    """(train, val, manifest) from data/verified and the data generator's repairs."""
    pairs = load("verified")
    held = load("eval")
    clash = {p["id"] for p in pairs} & {h["id"] for h in held}
    if clash:
        raise LeakError(f"in both data/verified and data/eval: {sorted(clash)}")
    by_pair = read_repairs(repairs_path)
    unknown = sorted(set(by_pair) - {p["id"] for p in pairs})
    if unknown:
        raise LeakError(f"repairs for pairs that are not in data/verified: {unknown}")
    train, val = [], []
    for pair in pairs:
        rows = write_examples(pair) + repair_examples(pair["id"], by_pair.get(pair["id"], []), per_pair)
        (val if in_validation(pair["id"]) else train).extend(rows)
    check_no_leaks(train + val, held)
    kinds = lambda rows, t: sum(1 for r in rows if r["type"] == t)   # noqa: E731
    manifest = {
        "prompt_sha256": digest(SYSTEM),
        "pairs": {"train": sorted(p["id"] for p in pairs if not in_validation(p["id"])),
                  "val": sorted(p["id"] for p in pairs if in_validation(p["id"]))},
        "examples": {"train": {"write": kinds(train, "write"), "repair": kinds(train, "repair")},
                     "val": {"write": kinds(val, "write"), "repair": kinds(val, "repair")}},
        "repairs_per_pair": per_pair,
        "held_out_checked": len(held),
        "lengths": {"train": lengths(train), "val": lengths(val)},
        "guards": ["only data/verified is read", "no held-out id, request or recipe in any example",
                   "split by pair", "no model is loaded or downloaded"],
    }
    return train, val, manifest


def write_all(dest, train, val, manifest):
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    files = {}
    for name, rows in (("train.jsonl", train), ("val.jsonl", val)):
        text = "".join(json.dumps(r, sort_keys=True, ensure_ascii=False) + "\n" for r in rows)
        (dest / name).write_text(text, encoding="utf-8")
        files[name] = digest(text)
    manifest = dict(manifest, files=files)
    (dest / "manifest.json").write_text(json.dumps(manifest, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    return manifest


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("-o", "--out", default=str(ROOT / "dist" / "train"))
    ap.add_argument("--datagen", help="a folder `vclc datagen` already wrote; made fresh when not given")
    ap.add_argument("--today", help="the date the TESTs run with (default today)")
    ap.add_argument("--repairs-per-pair", type=int, default=REPAIRS_PER_PAIR)
    args = ap.parse_args(argv)
    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()
    datagen = Path(args.datagen) if args.datagen else Path(args.out) / "datagen"
    if not (datagen / "repairs.jsonl").is_file():
        from vcl_datagen import generate
        print(f"making the repairs in {datagen} (a few minutes)...")
        generate(datagen, today)
    try:
        train, val, manifest = build(datagen / "repairs.jsonl", args.repairs_per_pair)
    except LeakError as err:
        print(f"STOP: {err}")
        return 1
    manifest = write_all(args.out, train, val, manifest)
    e = manifest["examples"]
    print(f"train: {e['train']['write']} write + {e['train']['repair']} repair; "
          f"val: {e['val']['write']} write + {e['val']['repair']} repair; written to {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
