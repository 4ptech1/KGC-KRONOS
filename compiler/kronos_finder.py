"""KRONOS's recipe finder: the proven recipe closest to a request, adapted to the person's workbook (D129).

    python compiler/kronos_finder.py --build      write data/library.json from data/verified and the demo recipes
    python compiler/kronos_finder.py --check      exit 1 when data/library.json is not what the sources give
    python compiler/kronos_finder.py "request"    the five closest recipes, for a quick look

This is KRONOS's own AI until the model arrives in March 2027, and it needs none: the memory of D90 (Phase 2 of
`docs/ai-vba-plan.md`), retrieval over our own proven recipes only (D25 bans retrieval over outside sources, not over
ours). Nothing goes online.

- The library (`data/library.json`) holds every pair in data/verified (its request, its three paraphrases, its GOAL,
  its recipe) and the three demo recipes in docs/demo. Never data/eval: the held-out items stay held out, so the finder
  can be scored on them like any other candidate (D59).
- search() ranks the library for a request with BM25 over the request, the paraphrases and the GOAL, word forms folded
  together and shop words mapped by SYNONYMS (an in-house table). Given the person's workbook headings, a recipe that
  reads columns the workbook has ranks higher.
- adapt() renames the columns a recipe reads to the workbook's own headings, only where the match is sure, and proves
  the renamed recipe again; if the proof fails the recipe is given back unchanged. Nothing is renamed without the
  person's tick on the page.
Standard library only, so it runs in KRONOS Offline's browser Python as it runs here.
"""
import datetime as dt
import difflib
import json
import math
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from resources import ROOT  # noqa: E402

LIBRARY = ROOT / "data" / "library.json"
DEMO = HERE.parent / "docs" / "demo"   # read only when the library is built, from a checkout: not bundled
BUILD_DAY = dt.date(2026, 11, 21)   # the date the library's recipes are proven with when it is built (fixed: same bytes)
K1, B = 1.5, 0.75                  # BM25
WEAK = 0.34                        # under this share of the request's words matched, the best recipe is "not a good match"
CLOSE = 0.85                       # difflib ratio for "a close spelling"

STOP = set("""a an the and or of to in on for from with by at as is are be it its this that these those every each all
any some my our your their i we you me us them which what when where who whose how if then else than so not no do does
did done make makes made put puts write writes written sheet sheets column columns row rows cell cells value values list
lists one two three macro please want need give gives into onto out up there here also only just per new same other
""".split())
# Shop words that mean the same thing, as one group each (in-house, D25): the first word names the group.
SYNONYMS = [
    "ncr nonconformance nonconformity nonconforming",
    "nde ndt radiography rt ut mt pt ultrasonic xray",
    "grn receipt received receive goods",
    "mtc certificate cert mill heat",
    "qty quantity",
    "cal calibration calibrate calibrated instrument gauge",
    "wps pqr procedure",
    "weld welding welded seam joint",
    "welder qualification qual quals",
    "bom bill",
    "itp inspection",
    "dispatch shipping shipment delivery deliver shipped delivered",
    "consumable electrode filler wire rod flux",
    "highlight colour color mark flag fill",
    "summary summarise summarize total count report tally",
    "hide filter",
    "duplicate dupe repeat repeated twice",
    "tidy clean trim blank empty",
    "overdue late expire expiry expired due",
    "number no num",
    "date day",
    "job project order",
    "reject rejected fail failed repair",
    "copy duplicate",
    "compare reconcile check against versus vs",
]
ABBREVIATIONS = {"no": "number", "num": "number", "nr": "number", "qty": "quantity", "desc": "description",
                 "dt": "date", "rev": "revision", "cert": "certificate", "cal": "calibration", "insp": "inspection",
                 "ref": "reference", "pct": "percent", "%": "percent", "#": "number"}


class FinderError(Exception):
    pass


# ---------------------------------------------------------------- words

def stem(word):
    for end in ("ies", "es", "s", "ing", "ed"):
        if len(word) > len(end) + 2 and word.endswith(end):
            return word[:-len(end)] + ("y" if end == "ies" else "")
    return word


_GROUP = {}
for _line in SYNONYMS:
    _words = _line.split()
    for _w in _words:
        _GROUP.setdefault(stem(_w), stem(_words[0]))


def words(text):
    """The words of a text that carry meaning: lower case, word forms folded, shop words mapped to their group."""
    out = []
    for raw in re.findall(r"[a-z0-9]+", (text or "").lower().replace("non-conform", "nonconform")):
        if raw in STOP or len(raw) < 2 and not raw.isdigit():
            continue
        w = stem(raw)
        out.append(_GROUP.get(w, w))
    return out


def norm_heading(text):
    """A heading as compared: lower case, punctuation gone, abbreviations written out."""
    parts = re.findall(r"[a-z0-9]+|#|%", (text or "").lower())
    return " ".join(ABBREVIATIONS.get(p, p) for p in parts)


# ---------------------------------------------------------------- the library

def reads_of(spec_text):
    """[{sheet, headers}] of the sheets a recipe reads: its FIXTURE sheets that have data rows (a sheet that starts empty is
    one it writes), each sheet once, the headers in the order given. Template FIXTUREs give no headers in the text."""
    from parser import parse_spec
    seen = {}
    for fx in parse_spec(spec_text).fixtures:
        if not fx.headers or getattr(fx, "rows", 0) == 0:
            continue
        entry = seen.setdefault(fx.sheet, [])
        for _, header in fx.headers:
            if header not in entry:
                entry.append(header)
    return [{"sheet": sheet, "headers": headers} for sheet, headers in seen.items()]


def _entry(item_id, source, request, paraphrases, spec_text, family="", registers=()):
    from kronos_engine import prove, summary
    proof = prove(spec_text, BUILD_DAY, offline=True)
    if not proof.ok:
        raise FinderError(f"{item_id} is not proven today: {[f.message for f in proof.errors][:2]}")
    words_of = summary(proof)
    return {"id": item_id, "source": source, "kind": proof.kind, "family": family, "registers": list(registers),
            "request": request, "paraphrases": list(paraphrases), "goal": words_of["does"],
            "changes": words_of["changes"], "assumes": words_of["assumes"], "reads": reads_of(spec_text),
            "spec": spec_text}


def build_library():
    """The library's entries, from data/verified and docs/demo; never data/eval."""
    from vcl_corpus import items, load
    held = {folder.name for folder in items("eval")}
    out = []
    for pair in load("verified"):
        if pair["id"] in held:
            raise FinderError(f"{pair['id']} is held out (data/eval): it can't be in the library")
        out.append(_entry(pair["id"], "corpus", pair["request"], pair["paraphrases"], pair["spec"], pair["family"],
                          pair["registers"]))
    for path in sorted(DEMO.glob("recipe_*.vcl")):
        text = path.read_text(encoding="utf-8")
        goal = next(line.split(None, 1)[1] for line in text.splitlines() if line.startswith("GOAL "))
        out.append(_entry("demo-" + path.stem.split("_", 2)[1], "demo", goal, [], text, "", ["ncr"]))
    return out


def library_text(entries=None):
    entries = build_library() if entries is None else entries
    return json.dumps({"version": 1, "built_with": BUILD_DAY.isoformat(), "entries": entries}, indent=1,
                      sort_keys=True, ensure_ascii=False) + "\n"


_LOADED = None


def load_library(path=LIBRARY):
    global _LOADED
    if _LOADED is None or path != LIBRARY:
        data = json.loads(Path(path).read_text(encoding="utf-8"))["entries"]
        if path != LIBRARY:
            return data
        _LOADED = data
    return _LOADED


# ---------------------------------------------------------------- search

class Index:
    """BM25 over each entry's text: request, paraphrases, GOAL, the registers and the columns it reads."""

    def __init__(self, entries, leave_out=None):
        self.entries = entries
        self.docs = []
        for e in entries:
            texts = [e["request"], e["goal"]] + [p for p in e["paraphrases"] if p != leave_out]
            if e["request"] == leave_out:
                texts = texts[1:]
            texts += [r.replace("_", " ") for r in e["registers"]]
            texts += [h for r in e["reads"] for h in r["headers"]]
            self.docs.append(words(" ".join(texts)))
        self.n = len(self.docs)
        self.avg = sum(len(d) for d in self.docs) / max(self.n, 1)
        self.df = {}
        for d in self.docs:
            for w in set(d):
                self.df[w] = self.df.get(w, 0) + 1
        self.tf = [self._counts(d) for d in self.docs]

    @staticmethod
    def _counts(doc):
        c = {}
        for w in doc:
            c[w] = c.get(w, 0) + 1
        return c

    def idf(self, w):
        n = self.df.get(w, 0)
        return math.log(1 + (self.n - n + 0.5) / (n + 0.5))

    def score(self, query, i):
        tf, length = self.tf[i], len(self.docs[i])
        total = 0.0
        for w in set(query):
            f = tf.get(w, 0)
            if f:
                total += self.idf(w) * f * (K1 + 1) / (f + K1 * (1 - B + B * length / self.avg))
        return total


def fit(entry, headings):
    """(found, total): how many of the columns the recipe reads the person's workbook has (any sheet, any spelling the
    adapter would accept)."""
    if not headings:
        return 0, 0
    have = {norm_heading(h) for hs in headings.values() for h in hs}
    want = [h for r in entry["reads"] for h in r["headers"]]
    found = sum(1 for h in want if norm_heading(h) in have or _close(norm_heading(h), have))
    return found, len(want)


def _close(name, candidates):
    best = difflib.get_close_matches(name, list(candidates), n=2, cutoff=CLOSE)
    return len(best) == 1


def search(request, headings=None, k=5, entries=None, leave_out=None):
    """The k recipes closest to the request: [{id, goal, score, strength, matched, fit, changes, kind, source}], best first.
    headings: {sheet: [headings]} of the person's workbook, or None. strength is the share of the request's words the
    recipe matches; under WEAK the page says nothing in the library matches well."""
    entries = load_library() if entries is None else entries
    query = words(request)
    if not query:
        return []
    index = Index(entries, leave_out)
    scored = []
    for i, e in enumerate(entries):
        s = index.score(query, i)
        if s <= 0:
            continue
        found, total = fit(e, headings)
        if total:
            s *= 1 + 0.5 * found / total
        scored.append((s, i))
    scored.sort(key=lambda t: (-t[0], entries[t[1]]["id"]))
    out = []
    for s, i in scored[:k]:
        e = entries[i]
        doc = set(index.docs[i])
        matched = sorted({w for w in query if w in doc})
        found, total = fit(e, headings)
        out.append({"id": e["id"], "goal": e["goal"], "kind": e["kind"], "source": e["source"], "score": round(s, 3),
                    "strength": round(len(set(matched)) / len(set(query)), 2), "matched": matched,
                    "fit": {"found": found, "total": total}, "changes": e["changes"], "assumes": e["assumes"],
                    "reads": e["reads"]})
    return out


def entry(item_id, entries=None):
    for e in load_library() if entries is None else entries:
        if e["id"] == item_id:
            return e
    raise FinderError(f"there is no recipe {item_id} in the library")


# ---------------------------------------------------------------- adapting a recipe to a workbook

def _match(header, headings, taken):
    """(the workbook heading, why) the recipe's header is surely the same as, or (None, why not)."""
    free = [h for h in headings if h not in taken]
    if header in free:
        return header, "the same name"
    want = norm_heading(header)
    same = [h for h in free if norm_heading(h) == want]
    if len(same) == 1:
        plain = re.sub(r"[^a-z0-9]", "", header.lower()) == re.sub(r"[^a-z0-9]", "", same[0].lower())
        return same[0], "the same name written differently" if plain else "a known abbreviation"
    if len(same) > 1:
        return None, "more than one column could be it"
    close = difflib.get_close_matches(want, [norm_heading(h) for h in free], n=2, cutoff=CLOSE)
    if len(close) == 1:
        hits = [h for h in free if norm_heading(h) == close[0]]
        if len(hits) == 1:
            return hits[0], "a close spelling"
    return None, "not found in your workbook"


def proposal(spec_text, headings):
    """What adapt() would rename: {renames: [{sheet, old, new, why}], missing: [{sheet, old, why}], sheet_map: {recipe
    sheet: workbook sheet}}. headings: {sheet: [headings]} of the person's workbook."""
    renames, missing, sheet_map = [], [], {}
    reads = reads_of(spec_text)
    all_headers = {h for r in reads for h in r["headers"]}
    for r in reads:
        best, best_hits = None, -1
        for sheet, hs in (headings or {}).items():
            hits = sum(1 for h in r["headers"] if _match(h, hs, set())[0] is not None)
            if hits > best_hits:
                best, best_hits = sheet, hits
        if best is None or best_hits <= 0:
            missing += [{"sheet": r["sheet"], "old": h, "why": "no sheet in your workbook has these columns"}
                        for h in r["headers"]]
            continue
        sheet_map[r["sheet"]] = best
        taken = set()
        for header in r["headers"]:
            new, why = _match(header, headings[best], taken)
            if new is None:
                missing.append({"sheet": r["sheet"], "old": header, "why": why})
                continue
            taken.add(new)
            if new == header:
                continue
            if any(c in new for c in ',"]') or new in all_headers or not new.strip():
                missing.append({"sheet": r["sheet"], "old": header,
                                "why": f"your column {new!r} can't be written into a recipe as it is; change it by hand"})
                continue
            renames.append({"sheet": r["sheet"], "old": header, "new": new, "why": why})
    # a header written once in the recipe but read from two sheets is one piece of text: it can take one new name only
    targets = {}
    for r in renames:
        targets.setdefault(r["old"], set()).add(r["new"])
    torn = {old for old, news in targets.items() if len(news) > 1}
    for old in sorted(torn):
        missing.append({"sheet": ", ".join(sorted({r["sheet"] for r in renames if r["old"] == old})), "old": old,
                        "why": f"your sheets spell it differently ({', '.join(sorted(targets[old]))}); the recipe uses "
                               "one name for both, so change it by hand"})
    renames = [r for r in renames if r["old"] not in torn]
    seen, unique = set(), []
    for r in renames:   # the same rename found on two sheets is shown once
        if r["old"] not in seen:
            seen.add(r["old"])
            unique.append(r)
    return {"renames": unique, "missing": missing, "sheet_map": sheet_map}


def rename(spec_text, pairs):
    """The recipe with each old header replaced by its new one: as a whole quoted string, in a [Header] reference, and in a
    FIXTURE's cols=. Two steps through placeholders, so a rename never feeds another."""
    text = spec_text
    marks = {}
    for k, (old, new) in enumerate(pairs):
        mark = f"\x00{k}\x00"
        marks[mark] = new
        q_old = old.replace('"', '""')
        text = text.replace(f'"{q_old}"', f'"{mark}"')
        text = text.replace(f"[{old}]", f"[{mark}]")
        lines = []
        for line in text.split("\n"):
            if line.startswith("FIXTURE") and "cols=" in line:
                line = re.sub(rf"(?<=:){re.escape(old)}(?=,[A-Z]{{1,3}}:| rows=)", lambda m: mark, line)
            lines.append(line)
        text = "\n".join(lines)
    for mark, new in marks.items():
        text = text.replace(mark, new)
    return text


def adapt(spec_text, headings, accepted=None, today=None):
    """{ok, text, renames, missing, message}: the recipe with the accepted renames (all of proposal()'s when accepted is
    None), proven again. When the renamed recipe is not proven, ok is False and text is the original recipe."""
    from kronos_engine import prove
    plan = proposal(spec_text, headings)
    chosen = [r for r in plan["renames"] if accepted is None or r["old"] in accepted]
    if not chosen:
        return {"ok": True, "text": spec_text, "renames": [], "missing": plan["missing"],
                "message": "Nothing to rename: the recipe is used as it is."}
    text = rename(spec_text, [(r["old"], r["new"]) for r in chosen])
    proof = prove(text, today or dt.date.today(), offline=True)
    if not proof.ok:
        return {"ok": False, "text": spec_text, "renames": [], "missing": plan["missing"],
                "message": "With your column names the recipe didn't pass its own tests, so KRONOS kept the original "
                           "names; change them by hand in the recipe if you need to."}
    return {"ok": True, "text": proof.spec, "renames": chosen, "missing": plan["missing"],
            "message": f"{len(chosen)} column name{'' if len(chosen) == 1 else 's'} changed to your workbook's, "
                       "and the recipe passed its own tests again."}


def main(argv):
    if argv[:1] == ["--build"]:
        LIBRARY.write_text(library_text(), encoding="utf-8")
        print(f"wrote {LIBRARY} ({len(load_library(LIBRARY))} recipes)")
        return 0
    if argv[:1] == ["--check"]:
        if not LIBRARY.is_file() or LIBRARY.read_text(encoding="utf-8") != library_text():
            print(f"{LIBRARY} is stale: run python compiler/kronos_finder.py --build")
            return 1
        print(f"{LIBRARY} is up to date")
        return 0
    for r in search(" ".join(argv)):
        print(f"{r['id']:10} {r['score']:6.2f} {r['strength']:.2f}  {r['goal'][:90]}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
