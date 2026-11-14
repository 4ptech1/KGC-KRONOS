"""Build KRONOS Offline (D86, docs/kronos-offline.md): the folder KRONOS-Offline/ and its zip.

    python packaging/build_offline.py [--out build/offline] [--cache build/offline-cache] [--date DD-MM-YYYY]

KRONOS Offline is KRONOS in the browser, with nothing installed: Boss unzips the folder on the work PC, double-clicks
"Start KRONOS (Windows)", and a tiny local server (Node.js, Python or Windows' own PowerShell) shows the page, which
runs this repo's toolchain through Pyodide. This script assembles it, using only Python's standard library:

  KRONOS-Offline/Start KRONOS (Windows).bat   packaging/offline/, Windows line endings (CRLF)
  KRONOS-Offline/server/serve.{mjs,py,ps1}    packaging/offline/server/, CRLF, with the version in the start message
  KRONOS-Offline/app/                         web/ (the version in the page), and:
    toolchain/   every compiler/*.py, vcl/semantics.py and the data files of packaging/kronos.spec, byte for byte as
                 in the repo, with manifest.json (what worker.js writes into Pyodide)
    pyodide/, wheels/, fonts/   the files in packaging/offline_pins.json, downloaded once into the cache and refused
                 unless their SHA-256 matches
    licences/    the licence texts in packaging/offline/licences/, CRLF

The version is the last row of the version list in docs/kronos-offline.md, and the date is today in India unless
--date gives it. The zip is KRONOS-Offline-v<N>-<DD-MM-YYYY>.zip, written with sorted names and fixed times, so the
same inputs give the same bytes; beside it go <zip>.sha256.txt and build-info.json. Exit 0 when built, 1 otherwise.
"""
import argparse
import datetime
import hashlib
import io
import json
import re
import shutil
import sys
import tarfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PINS = ROOT / "packaging" / "offline_pins.json"
OFFLINE = ROOT / "packaging" / "offline"
LICENCES = OFFLINE / "licences"
WEB = ROOT / "web"
SPEC = ROOT / "packaging" / "kronos.spec"
DOC = ROOT / "docs" / "kronos-offline.md"
FOLDER = "KRONOS-Offline"   # the folder inside the zip keeps its name whatever the version, so Boss's folder does too
LAUNCHER = "Start KRONOS (Windows).bat"
SERVERS = ("serve.mjs", "serve.ps1", "serve.py")
PAGE_MARK, SERVER_MARK = "Version dev", "version dev"   # replaced by the version and date, exactly once per file
INDIA = datetime.timezone(datetime.timedelta(hours=5, minutes=30))   # India has kept +05:30 since 1945
DEFAULT_OUT = ROOT / "build" / "offline"
DEFAULT_CACHE = ROOT / "build" / "offline-cache"


class Unreachable(Exception):
    """A download could not be made (no network, or the site is down): the tests skip on it, CI fails."""


# ---------- what goes in ----------

def versions(doc=DOC):
    """The version list in docs/kronos-offline.md: [(version, date, built from, what changed)], numbered 1, 2, 3 ..."""
    text = doc.read_text(encoding="utf-8")
    section = text.split("\n## Versions\n", 1)[1].split("\n## ", 1)[0] if "\n## Versions\n" in text else ""
    rows = [(int(n), date, commit.strip(), words.strip()) for n, date, commit, words
            in re.findall(r"^\| (\d+) \| (\d\d-\d\d-\d{4}) \|([^|]*)\|([^|]*)\|\s*$", section, re.M)]
    numbers = [row[0] for row in rows]
    if not numbers or numbers != list(range(1, len(numbers) + 1)):
        raise ValueError(f"{doc.name}: the version list must number its rows 1, 2, 3 ..., got {numbers}")
    return rows


def latest_version(doc=DOC):
    """The version being built: the last row of the version list."""
    return versions(doc)[-1][0]


def spec_data_files():
    """The data files packaging/kronos.spec bundles into KRONOS.exe, as repo paths: the browser gets the same ones."""
    found = re.findall(r'\(str\(root((?: / "[^"]+")+)\), "[^"]+"\)', SPEC.read_text(encoding="utf-8"))
    return sorted("/".join(re.findall(r'"([^"]+)"', part)) for part in found)


def toolchain_files():
    """Every file the page's Python gets, as repo paths, sorted: the toolchain exactly as KRONOS.exe has it."""
    names = {f"compiler/{p.name}" for p in (ROOT / "compiler").glob("*.py")} | {"vcl/semantics.py"}
    return sorted(names | set(spec_data_files()))


def load_pins(path=PINS):
    return json.loads(path.read_text(encoding="utf-8"))


def crlf(data):
    """Windows line endings, for every text file a person may open in Notepad (rule 3 of the recipe)."""
    return data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")


def stamp(data, mark, words, name):
    text = data.decode("utf-8")
    if text.count(mark) != 1:
        raise ValueError(f"{name} must hold {mark!r} exactly once, to be replaced by the version")
    return text.replace(mark, words).encode("utf-8")


# ---------- the downloads ----------

def sha256(data):
    return hashlib.sha256(data).hexdigest()


def download(url, cache, want):
    """The file at url, from the cache when it is there and matches, else downloaded; refused if its hash differs."""
    cache.mkdir(parents=True, exist_ok=True)
    path = cache / url.rsplit("/", 1)[1]
    if path.is_file() and sha256(path.read_bytes()) == want:
        return path.read_bytes()
    try:
        with urllib.request.urlopen(url, timeout=120) as response:   # noqa: S310 (a pinned https URL)
            data = response.read()
    except (urllib.error.URLError, OSError) as exc:
        raise Unreachable(f"could not download {url}: {exc}") from None
    got = sha256(data)
    if got != want:
        raise ValueError(f"{url} has SHA-256 {got}, not the pinned {want}: refused")
    path.write_bytes(data)
    return data


def pinned_files(pins, cache):
    """{path under KRONOS-Offline/: bytes} for every download in the pins, each checked against its own hash."""
    out = {}
    for item in pins["downloads"]:
        data = download(item["url"], cache, item["sha256"])
        if "to" in item:
            out[item["to"]] = data
            continue
        with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as archive:
            for part in item["extract"]:
                member = archive.extractfile(part["member"])
                if member is None:
                    raise ValueError(f"{item['name']}: {part['member']} is not a file in {item['url']}")
                body = member.read()
                if sha256(body) != part["sha256"]:
                    raise ValueError(f"{item['name']}: {part['member']} does not match its pinned SHA-256: refused")
                out[part["to"]] = body
    return out


# ---------- the folder and the zip ----------

def contents(version, date, cache, pins=None):
    """{path under KRONOS-Offline/: bytes}: everything the zip holds."""
    pins = pins or load_pins()
    files = {}
    launcher = (OFFLINE / LAUNCHER).read_bytes()
    if crlf(launcher) != launcher:
        raise ValueError(f"packaging/offline/{LAUNCHER} must be saved with Windows line endings (CRLF)")
    files[LAUNCHER] = launcher
    for name in SERVERS:
        body = stamp((OFFLINE / "server" / name).read_bytes(), SERVER_MARK, f"version {version}, built {date}", name)
        files[f"server/{name}"] = crlf(body)
    for path in sorted(p for p in WEB.rglob("*") if p.is_file()):
        rel = path.relative_to(WEB).as_posix()
        body = path.read_bytes()
        if rel == "index.html":
            body = stamp(body, PAGE_MARK, f"Version {version} · {date}", "web/index.html")
        files[f"app/{rel}"] = body
    names = toolchain_files()
    for name in names:
        files[f"app/toolchain/{name}"] = (ROOT / name).read_bytes()
    downloaded = pinned_files(pins, cache)
    wheels = sorted(Path(to).name for to in downloaded if to.startswith("app/wheels/"))
    manifest = {"files": names, "wheels": wheels, "version": version, "date": date}
    files["app/toolchain/manifest.json"] = (json.dumps(manifest, indent=1) + "\n").encode("utf-8")
    files.update(downloaded)
    for path in sorted(LICENCES.glob("*.txt")):
        files[f"app/licences/{path.name}"] = crlf(path.read_bytes())
    return files


def zip_bytes(files, date):
    """The zip: KRONOS-Offline/ and everything in it, sorted, every entry dated the build day at noon."""
    day, month, year = (int(n) for n in date.split("-"))
    when = (year, month, day, 12, 0, 0)
    folders = {""}
    for name in files:
        parts = name.split("/")[:-1]
        folders.update("/".join(parts[:i]) for i in range(1, len(parts) + 1))
    entries = [(f"{FOLDER}/{d}/" if d else f"{FOLDER}/", None) for d in folders]
    entries += [(f"{FOLDER}/{name}", data) for name, data in files.items()]
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in sorted(entries):
            info = zipfile.ZipInfo(name, when)
            info.create_system = 0   # written as Windows writes it, whatever computer builds it
            if data is None:
                info.external_attr = 0x10   # a folder
                archive.writestr(info, b"")
            else:
                info.external_attr = 0x20
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, data, compresslevel=9)
    return buffer.getvalue()


def build(out=DEFAULT_OUT, cache=DEFAULT_CACHE, date=None, say=print):
    """Build the folder and the zip into out. Returns build-info as a dict."""
    out, cache = Path(out), Path(cache)
    version = latest_version()
    date = date or datetime.datetime.now(INDIA).strftime("%d-%m-%Y")
    if not re.fullmatch(r"\d\d-\d\d-\d{4}", date):
        raise ValueError(f"the date must be DD-MM-YYYY, got {date!r}")
    datetime.datetime.strptime(date, "%d-%m-%Y")
    files = contents(version, date, cache)
    data = zip_bytes(files, date)
    out.mkdir(parents=True, exist_ok=True)
    folder = out / FOLDER
    if folder.exists():
        shutil.rmtree(folder)
    for name, body in files.items():
        (folder / name).parent.mkdir(parents=True, exist_ok=True)
        (folder / name).write_bytes(body)
    name = f"{FOLDER}-v{version}-{date}.zip"
    for old in out.glob(f"{FOLDER}-v*.zip*"):
        old.unlink()
    (out / name).write_bytes(data)
    digest = sha256(data)
    (out / f"{name}.sha256.txt").write_bytes(f"{digest}  {name}\n".encode("ascii"))
    info = {"version": version, "date": date, "zip": name, "sha256": digest, "bytes": len(data), "files": len(files)}
    (out / "build-info.json").write_text(json.dumps(info, indent=1) + "\n", encoding="utf-8")
    say(f"KRONOS Offline version {version}, built {date}: {len(files)} files")
    say(f"{(out / name).relative_to(ROOT) if ROOT in (out / name).parents else out / name}  "
        f"({len(data) / 1e6:.1f} MB, SHA-256 {digest})")
    return info


def main(argv):
    ap = argparse.ArgumentParser(prog="build_offline.py", description="Build KRONOS Offline: its folder and its zip.")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="where to put KRONOS-Offline/ and the zip")
    ap.add_argument("--cache", default=str(DEFAULT_CACHE), help="where the downloads are kept")
    ap.add_argument("--date", help="the build date, DD-MM-YYYY (default: today in India)")
    args = ap.parse_args(argv)
    try:
        build(args.out, args.cache, args.date)
    except (Unreachable, ValueError, OSError) as exc:
        print(f"build_offline.py: {exc}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
