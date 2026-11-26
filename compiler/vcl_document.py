"""The DOCUMENT checks (VCL 0.7 part 3f, docs/vep-0.7-documents.md, D115): a spec that fills a Word template.

    check_document(text, spec) -> [Problem]

spec is the line parser's Spec (parser.py), for GOAL, ARGS, FIXTURE and TEST lines; FIELD, REPEAT and KEEP lines are
read with the grammar (vcl_parse.py). The spec is checked against its template (data/templates/<TEMPLATE>): every
{{field}} is filled exactly once, every section is repeated or kept, every table number exists, and a repeat inside a
section matches its rows to the section's row.
"""
import re

from errcodes import code_of, coded
from lark import Token, Tree
from parser import Problem

NAME_RE = re.compile(r"[A-Z][A-Za-z0-9]*")
DOC_ARG_TYPES = {"Workbook", "String", "Long", "Double", "Boolean", "Date"}
DOC_FUNCS = {"NUMBER_TEXT", "DATE_TEXT"}   # a DOCUMENT's own value functions
DOC_ASSERTS = {"FIELD_TEXT", "TABLE_ROWS", "TABLE_COLS", "TABLE_CELL", "SECTION_COUNT", "TEXT_HAS"}
CONSTANTS = {"TRUE", "FALSE", "TODAY", "NOW"}


class _Checker:
    """What validator.check_goal and check_fixtures need from a checker."""

    def __init__(self, spec):
        self.spec, self.problems, self.line = spec, [], spec.line

    def error(self, message, line=None):
        self.problems.append(Problem(line or self.line, "error", message))

    def warning(self, message, line=None):
        self.problems.append(Problem(line or self.line, "warning", message))

    def at_least(self, version):
        from validator import VERSION_RANK
        return VERSION_RANK.get(self.spec.version, 1) >= VERSION_RANK[version]


def _calls(tree):
    return [t for t in tree.iter_subtrees() if t.data == "call"] if isinstance(tree, Tree) else []


def _tokens(tree):
    if isinstance(tree, Token):
        return [tree]
    return list(tree.scan_values(lambda v: isinstance(v, Token))) if isinstance(tree, Tree) else []


def check_document(text, spec):
    from validator import (FIXTURE_FILE, VALUE_FUNCS, check_fixtures, check_folders, check_goal)
    from vcl_docx import DocxError, template_fields
    from vcl_parse import parse as grammar_parse
    from vcl_template import TemplateError, load_docx_template
    c = _Checker(spec)
    if not c.at_least("VCL 0.7"):
        c.error(coded("DOCUMENT is VCL 0.7: start the spec with 'VCL 0.7' to use it", "V346"))
    check_goal(spec, True, c.error, c.warning, "NAME")
    name = spec.value("NAME")
    if not name:
        c.error(coded("missing NAME: the DOCUMENT's name, PascalCase (FillEstimateSheet)", "V347"))
    elif not NAME_RE.fullmatch(name):
        c.error(coded(f"NAME {name} must be PascalCase, letters and digits", "V348"))
    for d in spec.args:
        if d.type not in DOC_ARG_TYPES:
            c.error(coded(f"${d.name} As {d.type}: a DOCUMENT takes only " + ", ".join(sorted(DOC_ARG_TYPES)), "V349"), d.line)
    tree, found = grammar_parse(text)
    for line, message in found:
        c.problems.append(Problem(line, "error", coded(f"grammar: {message}", code_of(message))))
    if found or tree is None:
        return c.problems
    keys = tree["keys"]
    templates = [k for k in keys if k["key"] == "TEMPLATE"]
    if len(templates) != 1:
        c.error(coded("a DOCUMENT names exactly one TEMPLATE (a Word .docx in data/templates/)", "V350"),
                templates[1]["line"] if templates else spec.line)
        return c.problems
    try:
        shape = template_fields(load_docx_template(str(templates[0]["tree"].children[0])))
    except (TemplateError, DocxError) as exc:
        c.error(str(exc), templates[0]["line"])
        return c.problems
    args = {d.name: d.type for d in spec.args}
    used = set()   # the arguments a FIELD, REPEAT or KEEP line reads (D139)

    def check_expr(t, line):
        for call in _calls(t):
            fname = str(call.children[0])
            if fname not in VALUE_FUNCS and fname not in DOC_FUNCS:
                c.error(coded(f"unknown function {fname} in a DOCUMENT", "V351"), line)
            pos = call.children[1].children if len(call.children) > 1 else []
            if fname in DOC_FUNCS and len(pos) == 2:
                from semantics import DATE_FORMATS_TEXT, NUMBER_FORMATS
                allowed = NUMBER_FORMATS if fname == "NUMBER_TEXT" else DATE_FORMATS_TEXT
                fmt = pos[1]
                if not (isinstance(fmt, Token) and fmt.type == "STRING" and fmt[1:-1] in allowed):
                    c.error(coded(f"{fname} needs one of " + ", ".join(f'"{f}"' for f in allowed)
                                  + " as its second argument", "V352"), line)
            elif fname in DOC_FUNCS:
                c.error(coded(f"{fname} takes 2 arguments, got {len(pos)}", "V353"), line)
            if fname == "AI_TEXT" and not (len(pos) == 1 and isinstance(pos[0], Token) and pos[0].type == "STRING"):
                c.error(coded('AI_TEXT needs one "quoted" text, written in place', "V354"), line)
        for tok in _tokens(t):
            if tok.type == "ARG":
                used.add(str(tok)[1:])
            if tok.type == "ARG" and str(tok)[1:] not in args:
                c.error(coded(f"undeclared argument {tok} (declare it in ARGS)", "V355"), line)
            elif tok.type == "NAME" and str(tok) not in CONSTANTS and not _is_call_name(t, tok):
                c.error(coded(f"{tok} is not an argument or a constant: a DOCUMENT has no variables (write ${tok})", "V356"),
                        line)

    fields, repeats, keeps = {}, [], {}
    for k in keys:
        t, line = k["tree"], k["line"]
        if k["key"] == "FIELD":
            fname = str(t.children[0])
            if fname in fields:
                c.error(coded(f"FIELD {fname} is written twice", "V357"), line)
            fields[fname] = line
            check_expr(t.children[1], line)
        elif k["key"] == "REPEAT":
            kind, target = str(t.children[0]), t.children[1]
            target = int(target.children[1]) if len(target.children) == 2 else str(target.children[0])
            rest, match = t.children[3:], None
            if rest and isinstance(rest[0], Tree) and rest[0].data == "match_opt":
                match = (int(rest[0].children[1]), str(rest[0].children[2]))
                rest = rest[1:]
            binds = {}
            for b in rest:
                bname, col = str(b.children[0]), int(b.children[1])
                if bname in binds:
                    c.error(coded(f"REPEAT binds {bname} twice", "V358"), line)
                if col < 1:
                    c.error(coded(f"{bname}={col}: columns are counted from 1", "V359"), line)
                binds[bname] = col
            source = t.children[2]
            empty = isinstance(source, Tree) and source.data == "empty_list"   # from=[]: no rows, the part is removed (D138)
            if empty and match is not None:
                c.error(coded("match= says which rows go into each copy of a section, and from=[] has no rows", "V393"), line)
            if not binds and not (empty and kind == "section"):
                c.error(coded(f"REPEAT {kind} names its fields: add <field>=<column> (only a REPEAT section from=[] "
                              "may leave them out)", "V394"), line)
            if not empty:
                check_expr(source, line)
            if isinstance(source, Tree) and source.data == "list_lit":
                for row in source.children:
                    items = row.children if isinstance(row, Tree) and row.data == "list_lit" else [row]
                    for bname, col in binds.items():
                        if col > len(items):
                            c.error(coded(f"{bname}={col}: a row written in place has {len(items)} values", "V360"), line)
            if (kind == "section") != isinstance(target, str):
                c.error(coded("REPEAT row and column take table=<n>; REPEAT section takes a section name", "V361"), line)
                continue
            repeats.append((line, kind, target, binds, match, empty))
        elif k["key"] == "KEEP":
            sname = str(t.children[0])
            if sname in keeps:
                c.error(coded(f"KEEP {sname} is written twice", "V362"), line)
            keeps[sname] = line
            check_expr(t.children[1], line)

    for d in spec.args:   # every argument is a question or an upload for the person, so each one must be read (D139)
        if d.name not in used:
            c.error(coded(f"${d.name} is never read: a DOCUMENT asks the person for every argument, so drop it from ARGS "
                          "and from the TEST calls", "V396"), d.line)
    # the spec against its template
    tables, table_section, sections = shape["tables"], shape["table_section"], shape["sections"]
    section_repeats = {r[2]: r for r in repeats if r[1] == "section"}
    covered = set()   # (place, field): place is ("table", n) or ("section", name)
    for line, kind, target, binds, match, empty in repeats:
        if kind == "section":
            if target not in sections:
                c.error(coded(f"the template has no section {target}", "V363"), line)
                continue
            if target in keeps:
                c.error(coded(f"section {target} is both repeated and kept: use one", "V364"), line)
            if [r for r in repeats if r[1] == "section" and r[2] == target][0][0] != line:
                c.error(coded(f"section {target} is repeated twice", "V365"), line)
            if empty and binds:
                c.error(coded(f"REPEAT section {target} from=[] removes the section: it has no fields to bind", "V395"), line)
            for bname in binds:
                if bname not in sections[target]:
                    c.error(coded(f"section {target} has no field {bname}", "V366"), line)
            continue
        if not 1 <= target <= len(tables):
            c.error(coded(f"the template has no table {target} (it has {len(tables)})", "V367"), line)
            continue
        rows = tables[target - 1]
        if kind == "row":
            ok = any(set(binds) <= set().union(*row) for row in rows)
            where = "row"
        else:
            ncols = max(len(row) for row in rows)
            ok = any(set(binds) <= set().union(*(row[i] for row in rows if i < len(row))) for i in range(ncols))
            where = "column"
        if not ok:
            c.error(coded(f"table {target} has no {where} holding {', '.join(sorted(binds))}", "V368"), line)
        inside = table_section[target - 1]
        if inside in section_repeats:
            if match is None:
                c.error(coded(f"table {target} is inside section {inside}, which repeats: say which rows belong to "
                              f"each copy with match=<n>=<field of {inside}>", "V369"), line)
            elif match[1] not in section_repeats[inside][3]:
                c.error(coded(f"match={match[0]}={match[1]}: {match[1]} is not a field section {inside} binds", "V370"), line)
        elif match is not None:
            c.error(coded(f"match= is only for a table inside a repeated section; table {target} isn't", "V371"), line)
        covered |= {(target, b) for b in binds}
    for sname in sections:
        if sname not in section_repeats and sname not in keeps:
            c.error(coded(f"the template's section {sname} needs a REPEAT section {sname} ... or KEEP {sname} IF ...", "V372"))
    for sname, line in keeps.items():
        if sname not in sections:
            c.error(coded(f"the template has no section {sname}", "V373"), line)
    # every field is filled exactly once
    need = set(shape["body"]) | set(shape["header"])
    for sname in keeps:
        need |= sections.get(sname, set())
    in_repeated_tables = {f for (n, f) in covered}
    loose = set()   # fields a FIELD line must fill: not bound by a repeat at their place
    for f in need:
        places = [n for n, rows in enumerate(tables, 1) if any(f in cell for row in rows for cell in row)
                  and table_section[n - 1] not in section_repeats]
        if f in in_repeated_tables and places and all((n, f) in covered for n in places):
            continue
        loose.add(f)
    for sname, (line, _, _, binds, _, empty) in section_repeats.items():
        if sname not in sections or empty:   # from=[]: the section and the tables inside it are removed (D138)
            continue
        nested = {f for (n, f) in covered if table_section[n - 1] == sname}
        for f in sorted(sections[sname] - set(binds) - nested):
            c.error(coded(f"section {sname}'s field {f} is not bound: add {f}=<column> to its REPEAT", "V374"), line)
    for f in sorted(loose - set(fields)):
        c.error(coded(f"the template's field {{{{{f}}}}} has no FIELD line", "V375"))
    for f, line in fields.items():
        if f not in loose:
            c.error(coded(f"FIELD {f}: the template has no {{{{{f}}}}} outside a repeated row, column or section", "V376"),
                    line)
    # FIXTUREs and TESTs
    check_fixtures(c, spec.fixtures)
    check_folders(c, spec.folders, spec.fixtures)
    files = {fx.file for fx in spec.fixtures}
    if not spec.tests:
        c.error(coded("a DOCUMENT needs at least one TEST", "V377"))
    for test in spec.tests:
        c.line = test.line
        if test.procedure != name:
            c.error(coded(f"TEST calls {test.procedure}, but NAME is {name}", "V378"))
        if test.fixture not in files:
            c.error(coded(f"TEST uses fixture {test.fixture}, but no FIXTURE line builds it", "V379"))
        if test.read:
            c.error(coded("read= is for a TASK: a DOCUMENT reads every workbook it is given", "V380"))
        if test.runs != 1:
            c.error(coded("runs= is for a TASK: a DOCUMENT is filled once from its template", "V381"))
        if len(test.args) != len(spec.args):
            c.error(coded(f"TEST passes {len(test.args)} arguments, but ARGS has {len(spec.args)}", "V382"))
        else:
            for i, (d, v) in enumerate(zip(spec.args, test.args)):
                if i in test.quoted:   # D138
                    c.error(coded(f"TEST passes '{v}' for ${d.name}: a sheet name in quotes is for a TASK's Worksheet "
                                  "argument, and a DOCUMENT takes none", "V392"))
                elif d.type == "Workbook" and v != test.fixture:
                    c.error(coded(f"TEST passes {v} for ${d.name}, a Workbook: pass this TEST's fixture file, "
                                  f"{test.fixture}", "V383"))
                elif d.type == "Workbook" and not FIXTURE_FILE.fullmatch(v):
                    c.error(coded(f"TEST passes {v} for ${d.name}: a Workbook is a fixture file", "V384"))
        for fname in re.findall(r"\b([A-Z][A-Z_]+)\(", test.assertion):
            if fname not in DOC_ASSERTS:
                c.error(coded(f"a DOCUMENT's TEST checks the filled document: {fname} is not one of "
                              + ", ".join(sorted(DOC_ASSERTS)), "V385"))
        for f in re.findall(r"FIELD_TEXT\((\w+)\)", test.assertion):
            if f not in fields:
                c.error(coded(f"FIELD_TEXT({f}): {f} has no FIELD line", "V386"))
        for s in re.findall(r"SECTION_COUNT\((\w+)\)", test.assertion):
            if s not in sections:
                c.error(coded(f"SECTION_COUNT({s}): the template has no section {s}", "V387"))
    if spec.tests and not any(re.search(r"\bERROR_COUNT = 0\b", t.assertion) for t in spec.tests):
        c.error(coded("no TEST checks ERROR_COUNT = 0 (rule 5)", "V388"), spec.tests[0].line)
    c.line = spec.line
    c.warning(coded("KRONOS Offline only: a DOCUMENT fills a Word file in KRONOS Offline; no macro is made (D115)", "V389"))
    return c.problems


def _is_call_name(tree, tok):
    """True when tok is the name of a call in tree (CELL in CELL(...)), not a bare name."""
    return any(c.children[0] is tok for c in _calls(tree))
