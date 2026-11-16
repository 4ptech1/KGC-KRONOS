"""Recipes read from a file (D89): a .vcl, .txt or .md file, or a Word .docx, that holds one or more VCL recipes.

    read_recipes(data, name)   # {"ok", "name", "recipes": [{"title", "text"}], "message"}

KRONOS Offline's page calls it through kronos_web.read_recipe_file, so a person can open the file a recipe came in
(an AI's whole reply saved as .md, the demo guide, a recipe saved from GitHub) instead of copying and pasting it.
Only the standard library: a .docx is a zip of XML, read with zipfile and xml.etree.

Where the recipes are:
  - every fenced ```vcl block, when the text has any (as an AI's reply has, rule 1);
  - otherwise every run of recipe lines: from a "VCL 0.x" line, through TASK or BUILD, "# ASSUME:" lines, the line
    keys of compiler/parser.py (and any other word in capitals at the start of a line, such as a mistyped key, so the
    checker can name it) and a BUILD's "---" lines (each later section starts with "VCL 0.x" again), to the first line
    that is none of these. Blank lines inside a recipe are dropped; anything else ends it.
What a word processor typed is undone first: curly quotes become straight ones, non-breaking spaces plain ones, and
invisible characters go.
"""
import io
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kronos_engine import FENCE, QUESTION, normalize  # noqa: E402
from parser import BUILD_KEYS, TASK_KEYS  # noqa: E402

TYPES = (".vcl", ".txt", ".md", ".markdown", ".docx")
MAX_BYTES = 2 * 1024 * 1024
KEYS = set(TASK_KEYS) | set(BUILD_KEYS) | {"TASK", "BUILD"}
VERSION_LINE = re.compile(r"VCL \d+\.\d+")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TYPED = {   # what Word or a phone typed for the plain character
    "“": '"', "”": '"', "„": '"', "‟": '"', "″": '"', "«": '"', "»": '"',
    "‘": "'", "’": "'", "‚": "'", "‛": "'", "′": "'",
    " ": " ", " ": " ", " ": " ", " ": " ", " ": " ", " ": " ",
    "​": "", "‌": "", "‍": "", "⁠": "", "﻿": "", "­": "",
}


class FileError(Exception):
    """A file KRONOS can't read recipes from, in words for the person."""


# ---------- the file's text ----------

def decode(data):
    """A text file's text: UTF-8 (with or without a BOM), UTF-16 with a BOM, else Windows' cp1252."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def _paragraph(p):
    """One Word paragraph as text: its runs' text, a tab for w:tab, a line break for w:br and w:cr."""
    out = []
    for el in p.iter():
        if el.tag == W + "t":
            out.append(el.text or "")
        elif el.tag == W + "tab":
            out.append("\t")
        elif el.tag in (W + "br", W + "cr"):
            out.append("\n")
    return "".join(out)


def _blocks(parent, lines):
    for child in parent:
        if child.tag == W + "p":
            lines.append(_paragraph(child))
        elif child.tag == W + "sdt":   # a content control: its paragraphs, in order
            content = child.find(W + "sdtContent")
            if content is not None:
                _blocks(content, lines)
        # tables (w:tbl) and section settings hold no recipe lines


def docx_text(data):
    """The body text of a Word .docx, one line per paragraph (tables left out)."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as ex:
        raise FileError("This doesn't look like a Word document (.docx). In Word, choose File, Save As, Word Document "
                        "(.docx).") from ex
    body = ElementTree.fromstring(xml).find(W + "body")
    lines = []
    if body is not None:
        _blocks(body, lines)
    return "\n".join(lines)


def clean(text):
    """Undo what a word processor typed: straight quotes, plain spaces, no invisible characters, \\n line ends."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return "".join(TYPED.get(ch, ch) for ch in text)


# ---------- the recipes in it ----------

def _fenced(lines):
    """The ```vcl blocks of the text, as lists of lines."""
    blocks, i = [], 0
    while i < len(lines):
        m = FENCE.fullmatch(lines[i].strip())
        if not m or lines[i].strip() == "```":
            i += 1
            continue
        end = i + 1
        while end < len(lines) and lines[end].strip() != "```":
            end += 1
        if m.group(1).lower() == "vcl":
            blocks.append(lines[i + 1:end])
        i = end + 1
    return blocks


KEY_LIKE = re.compile(r"[A-Z][A-Z0-9_]{1,15}(\s|$)")   # a key, or a mistyped one (GOLA): kept, so the checker names it


def _is_recipe_line(line):
    words = line.split()
    if not words:
        return False
    return (words[0] in KEYS or bool(KEY_LIKE.match(line)) or line.startswith("# ASSUME:") or line == "---"
            or bool(VERSION_LINE.fullmatch(line)))


def _unfenced(lines):
    """Every run of recipe lines that starts with a "VCL 0.x" line, as lists of lines."""
    found, current, after_rule = [], None, False
    for raw in lines:
        line = raw.strip()
        if current is None:
            if VERSION_LINE.fullmatch(line):
                current, after_rule = [line], False
            continue
        if not line:
            continue   # a blank line (an empty paragraph) inside a recipe is dropped
        if VERSION_LINE.fullmatch(line) and not after_rule:
            found.append(current)   # a new recipe starts; a BUILD's sections start after its "---" lines
            current, after_rule = [line], False
            continue
        if not _is_recipe_line(line):
            found.append(current)
            current = None
            continue
        current.append(line)
        after_rule = line == "---"
    if current is not None:
        found.append(current)
    return found


def title_of(text):
    """What a person calls the recipe: its first GOAL, else its PROC or NAME."""
    keys = {}
    for line in text.splitlines():
        key, _, rest = line.partition(" ")
        if key in ("GOAL", "PROC", "NAME") and key not in keys:
            keys[key] = rest.strip()
    return keys.get("GOAL") or keys.get("PROC") or keys.get("NAME") or "A recipe"


def find_recipes(text):
    """[{title, text}] of every recipe in a text."""
    lines = clean(text).split("\n")
    blocks = _fenced(lines)
    if not blocks:
        blocks = _unfenced(lines)
    out = []
    for block in blocks:
        body = normalize("\n".join(block))
        if not body or body.lstrip().startswith(QUESTION):
            continue
        out.append({"title": title_of(body), "text": body})
    return out


def read_recipes(data, name):
    """The recipes in a file a person chose, or a message in plain words."""
    suffix = Path(str(name or "")).suffix.lower()
    out = {"ok": False, "name": str(name or ""), "recipes": [], "message": ""}
    try:
        if suffix == ".doc":
            raise FileError("This is an old Word .doc file. In Word, choose File, Save As, Word Document (.docx).")
        if suffix not in TYPES:
            raise FileError("KRONOS reads recipes from .vcl, .txt, .md and Word .docx files.")
        if len(data) > MAX_BYTES:
            raise FileError(f"This file is {len(data) // 1024} KB. KRONOS reads recipe files up to {MAX_BYTES // 1024} KB.")
        text = docx_text(data) if suffix == ".docx" else decode(data)
    except FileError as ex:
        out["message"] = str(ex)
        return out
    out["recipes"] = find_recipes(text)
    if not out["recipes"]:
        out["message"] = ("KRONOS found no recipe in this file. A recipe starts with a line such as VCL 0.6, or sits in a "
                          "```vcl block.")
        return out
    out["ok"] = True
    count = len(out["recipes"])
    out["message"] = "KRONOS found 1 recipe." if count == 1 else f"KRONOS found {count} recipes."
    return out
