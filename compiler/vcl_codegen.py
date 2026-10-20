"""Compile a VCL 0.1 to 0.4 TASK spec to an Excel 2016 VBA procedure (MVP compiler).

House style: Option Explicit at module level; every VCL meaning goes through ForgeLib; line
labels are spec_line * 10 + k, so Erl \\ 10 is the spec line; Excel state is saved first and
always restored; PRE failure returns -1, runtime errors return -2 (Functions).
"""
import sys
from pathlib import Path

from lark import Token, Tree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "vcl"))
import semantics as S  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import Checker, canon, mentions  # noqa: E402

VBA_RESERVED = set("""and as boolean byref byte byval call case const currency date decimal declare dim do
double each else elseif empty end enum eqv erase error event exit false for friend function get global
gosub goto if imp implements in integer is let lib like long loop lset me mod new next not nothing null
object on optional or paramarray preserve private property public raiseevent redim rem resume return rset
select set single static step stop string sub then to true type typeof until variant wend while with
withevents xor""".split())
DIM_TYPE = {"Long": "Long", "Double": "Double", "String": "String", "Boolean": "Boolean", "Date": "Date",
            "Variant": "Variant", "Range": "Range", "Worksheet": "Worksheet", "Workbook": "Workbook",
            "Dictionary": "Object", "Collection": "Collection"}
KIND_OF = {"Long": "num", "Double": "num", "String": "text", "Boolean": "bool", "Date": "date",
           "Variant": "raw", "Range": "obj", "Worksheet": "obj", "Workbook": "obj",
           "Dictionary": "dict", "Collection": "coll"}
OBJECT_KINDS = {"obj", "dict", "coll"}


class CompileError(Exception):
    pass


def vba_name(n):
    return n + "_v" if n.lower() in VBA_RESERVED else n


def pos_opts(call):
    args = call.children[1].children if len(call.children) > 1 else []
    pos = [a for a in args if not (isinstance(a, Tree) and a.data == "option")]
    opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
    return pos, opts


def color(tok):
    h = str(tok).lstrip("#")
    return f"RGB({int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)})"


class Gen:
    def __init__(self, spec, source_name=""):
        self.spec = spec
        self.source = source_name
        self.args, self.vars, self.loopvars = [], {}, {}
        self.proc, self.kind, self.ret = "", "Sub", None
        self.pre, self.perf = [], []
        for k in spec["keys"]:
            t = k["tree"]
            if t is None:
                raise CompileError(f"line {k['line']} doesn't parse")
            key = k["key"]
            if key == "PROC":
                self.proc = str(t.children[0])
            elif key == "KIND":
                self.kind = str(t.children[0])
            elif key == "ARGS":
                self.args = [(str(p.children[0]), str(p.children[1])) for p in t.children]
            elif key == "RET":
                self.ret = str(t.children[0])
            elif key == "VAR":
                self.vars[str(t.children[0])] = str(t.children[1])
            elif key == "PRE":
                self.pre.append((k["line"], k["text"], t.children))
            elif key == "PERF":
                self.perf = [str(x) for x in t.children]
        self.argtypes = dict(self.args)
        self.lines = []
        self.label_k = {}
        self.facts = {}
        self.loops = []   # innermost last: [SKIP_IF label, whether a SKIP_IF jumps to it]
        self.selectors = []   # SELECT_CASE selector variables, innermost last
        self.temps = {}   # compiler-made Variant variables (the selectors): name -> kind
        self.checker = Checker(spec)

    # ---------- output ----------
    def emit(self, text, depth, line=None):
        label = ""
        if line is not None:
            k = self.label_k.get(line, 0)
            if k > 9:
                raise CompileError(f"spec line {line} expands to more than 10 numbered VBA lines")
            self.label_k[line] = k + 1
            label = str(line * 10 + k)
        self.lines.append(f"{label:<6}{'    ' * depth}{text}")

    # ---------- expressions: return (code, kind) ----------
    def kind_of_name(self, name):
        if name in self.vars:
            k = KIND_OF[self.vars[name]]
            return self.facts.get(name, k) if k == "raw" else k
        if name in self.loopvars:
            return self.loopvars[name]
        if name in self.temps:
            return self.temps[name]
        return None

    def ex(self, t):
        if isinstance(t, Token):
            kind = t.type
            if kind == "NUMBER":
                return str(t), "num"
            if kind == "STRING":
                return str(t), "text"
            if kind == "COLOR":
                return color(t), "color"
            if kind == "ARG":
                name = str(t)[1:]
                return vba_name(name), KIND_OF[self.argtypes[name]]
            name = str(t)
            const = {"TODAY": ("Date", "date"), "NOW": ("Now", "date"), "TRUE": ("True", "bool"),
                     "FALSE": ("False", "bool"), "THIS_BOOK": ("ThisWorkbook", "obj"),
                     "ACTIVE_BOOK": ("ActiveWorkbook", "obj"), "ACTIVE_SHEET": ("ActiveSheet", "obj"),
                     "NOTHING": ("Nothing", "obj")}
            if name in const:
                return const[name]
            k = self.kind_of_name(name)
            if k is None:
                return name, "word"
            return vba_name(name), k
        d = t.data
        if d == "call":
            return self.ex_call(t)
        if d == "not_op":
            c, _ = self.ex(t.children[0])
            return f"(Not {c})", "bool"
        if d in ("and_expr", "or_expr"):
            op = " And " if d == "and_expr" else " Or "
            return "(" + op.join(self.ex(c)[0] for c in t.children) + ")", "bool"
        if d == "neg":
            c, k = self.as_num(t.children[0])
            return f"(-{c})", k
        if d == "concat":
            parts = []
            for c in t.children:
                code, k = self.ex(c)
                parts.append(code if k == "text" else f"ForgeText({code})")
            return "(" + " & ".join(parts) + ")", "text"
        if d in ("sum", "product"):
            code, k = self.as_num(t.children[0])
            for i in range(1, len(t.children), 2):
                op = str(t.children[i])
                rc, rk = self.as_num(t.children[i + 1])
                if k == "date" and rk == "num":
                    k = "date"
                elif k == "date" and rk == "date" and op == "-":
                    k = "num"
                elif k == "num" and rk == "date":
                    k = "date"
                code = f"({code} {op} {rc})"
            return code, k
        if d == "comparison":
            return self.ex_compare(t)
        raise CompileError(f"can't compile expression {d}")

    def as_num(self, t, prefer=None):
        """Code for t in a numeric or date context: raw values go through ForgeNum or CDate."""
        code, k = self.ex(t)
        if k in ("raw", "word"):
            fact = self.facts.get(canon(t))
            if fact == "Date" or (fact is None and prefer == "date"):
                return f"CDate({code})", "date"
            return f"ForgeNum({code})", "num"
        return code, k

    def ex_compare(self, t):
        a_t, op, b_t = t.children
        op = str(op)
        ac, ak = self.ex(a_t)
        bc, bk = self.ex(b_t)
        if op in ("=", "<>"):
            if "text" in (ak, bk) and ("raw" in (ak, bk) or "word" in (ak, bk) or ak == bk == "text"):
                if ak == bk == "text":
                    expr = f"({ac} = {bc})"
                else:
                    expr = f"ForgeEq({ac}, {bc})"
                return (expr if op == "=" else f"(Not {expr})"), "bool"
            if "raw" in (ak, bk):
                other = bk if ak == "raw" else ak
                prefer = "date" if other == "date" else None
                ac, _ = self.as_num(a_t, prefer) if ak == "raw" else (ac, ak)
                bc, _ = self.as_num(b_t, prefer) if bk == "raw" else (bc, bk)
            return f"({ac} {op} {bc})", "bool"
        prefer = "date" if "date" in (ak, bk) else None
        if ak == bk == "text":
            return f"({ac} {op} {bc})", "bool"
        ac, _ = self.as_num(a_t, prefer)
        bc, _ = self.as_num(b_t, prefer)
        return f"({ac} {op} {bc})", "bool"

    def rng(self, t):
        """Code for a range used as a target."""
        if isinstance(t, Tree) and t.data == "call":
            name = str(t.children[0])
            pos, _ = pos_opts(t)
            if name == "CELL":
                return f"{self.ex(pos[0])[0]}.Cells({self.as_num(pos[1])[0]}, {self.as_num(pos[2])[0]})"
            if name == "ROW_RANGE":
                return "ForgeRowRange(" + ", ".join(self.ex(p)[0] if i == 0 else self.as_num(p)[0] for i, p in enumerate(pos)) + ")"
            if name == "USED_RANGE":
                return f"ForgeUsedRange({self.ex(pos[0])[0]})"
            if name == "RANGE":
                return f"{self.ex(pos[0])[0]}.Range({self.ex(pos[1])[0]})"
        code, k = self.ex(t)
        if k == "obj":
            return code
        raise CompileError(f"{canon(t)} is not a range")

    def ex_call(self, t):
        name = str(t.children[0])
        pos, opts = pos_opts(t)
        e = lambda i: self.ex(pos[i])[0]
        n = lambda i: self.as_num(pos[i])[0]
        if name == "CELL":
            return f"{e(0)}.Cells({n(1)}, {n(2)}).Value", "raw"
        if name == "COL":
            return f"ForgeCol({e(0)}, {e(1)})", "num"
        if name == "LASTROW":
            return f"ForgeLastRow({e(0)}, {n(1)})", "num"
        if name == "LASTCOL":
            return f"ForgeLastCol({e(0)}, {n(1)})", "num"
        if name in ("USED_RANGE", "ROW_RANGE", "RANGE"):
            return self.rng(t), "obj"
        if name == "TRIM":
            return f"ForgeTrim({e(0)})", "text"
        if name == "UPPER":
            return f"UCase$(ForgeText({e(0)}))", "text"
        if name == "LOWER":
            return f"LCase$(ForgeText({e(0)}))", "text"
        if name == "LEN":
            return f"Len(ForgeText({e(0)}))", "num"
        if name == "LEFT":
            return f"ForgeLeft({e(0)}, {n(1)})", "text"
        if name == "RIGHT":
            return f"ForgeRight({e(0)}, {n(1)})", "text"
        if name == "IS_EMPTY":
            return f"ForgeIsEmpty({e(0)})", "bool"
        if name == "IS_NUMBER":
            return f"ForgeIsNumber({e(0)})", "bool"
        if name == "IS_DATE":
            return f"ForgeIsDate({e(0)})", "bool"
        if name == "FIND":   # S22: Nothing when no cell matches (D39)
            whole = "False" if str(opts.get("match", "whole")) == "part" else "True"
            return f"ForgeFind({self.rng(pos[0])}, {e(1)}, {whole})", "obj"
        if name == "IS_NOTHING":
            return f"({e(0)} Is Nothing)", "bool"
        if name == "SHEET_EXISTS":   # VCL 0.4 (D40)
            return f"ForgeSheetExists({e(0)}, {e(1)})", "bool"
        if name == "SAME":
            return f"ForgeSame({e(0)}, {e(1)})", "bool"
        if name == "CLEAN_TEXT":
            return f"ForgeCleanText({e(0)})", "text"
        if name == "IS_ERROR":
            return f"ForgeIsError({e(0)})", "bool"
        if name == "PARSE_DATE":
            return f"ForgeParseDate({e(0)}, {e(1)})", "raw"
        if name == "IIF":
            (ac, ak), (bc, bk) = self.ex(pos[1]), self.ex(pos[2])
            return f"ForgeIif({e(0)}, {ac}, {bc})", ak if ak == bk else "raw"
        if name == "MOD":
            return f"ForgeMod({n(0)}, {n(1)})", "num"
        if name == "SAFE_NAME":
            return f"ForgeSafeName({e(0)})", "text"
        if name == "PATH_JOIN":
            return f"ForgePathJoin({e(0)}, Array({', '.join(self.ex(p)[0] for p in pos[1:])}))", "text"
        if name == "FOLDER_EXISTS":
            return f"ForgeFolderExists({e(0)})", "bool"
        if name == "EXISTS":
            return f"ForgeDictExists({e(0)}, {e(1)})", "bool"
        if name == "COUNT":
            return f"{e(0)}.Count", "num"
        if name == "SHEET":
            return f"{e(0)}.Worksheets({e(1)})", "obj"
        if name == "ARR":
            return f"{e(0)}({n(1)}, {n(2)})", "raw"
        if name == "ROWS":
            return f"(UBound({e(0)}, 1) - LBound({e(0)}, 1) + 1)", "num"
        if name == "COLS":
            return f"(UBound({e(0)}, 2) - LBound({e(0)}, 2) + 1)", "num"
        if name == "FILE_NAME":
            return f"ForgeFileName({e(0)})", "text"
        raise CompileError(f"{name} is not supported by the MVP compiler")

    # ---------- statements ----------
    def assign_code(self, var, t):
        code, k = self.ex(t)
        typ = self.vars.get(var)
        if typ is None:
            raise CompileError(f"{var} is not declared")
        if typ == "String" and k != "text":
            code = f"ForgeText({code})"
        elif typ in ("Long", "Double") and k in ("raw", "word"):
            code = self.as_num(t)[0]
        elif typ == "Date" and k == "raw":
            code = f"CDate({code})"
        prefix = "Set " if KIND_OF[typ] in OBJECT_KINDS or (typ == "Variant" and k in OBJECT_KINDS) else ""
        return f"{prefix}{vba_name(var)} = {code}"

    def action_lines(self, call):
        """VBA statements for one action, as a list of strings."""
        name = str(call.children[0])
        pos, opts = pos_opts(call)
        e = lambda x: self.ex(x)[0]
        if name == "LET":
            return [self.assign_code(str(pos[0]), pos[1])]
        if name == "SET_VALUE":
            if isinstance(pos[0], Tree) and pos[0].data == "call" and str(pos[0].children[0]) == "ARR":
                ap, _ = pos_opts(pos[0])
                return [f"{e(ap[0])}({self.as_num(ap[1])[0]}, {self.as_num(ap[2])[0]}) = {e(pos[1])}"]
            return [f"ForgeSetValue {self.rng(pos[0])}, {e(pos[1])}"]   # S17: text stays text
        if name == "WRITE_ROW":
            col = self.as_num(opts["col"])[0] if "col" in opts else "1"
            values = ", ".join(e(v) for v in pos[2:])
            return [f"ForgeWriteRow {e(pos[0])}, {self.as_num(pos[1])[0]}, {col}, Array({values})"]
        if name == "CLEAR":
            what = str(opts.get("what", "all"))
            method = {"all": "Clear", "contents": "ClearContents", "formats": "ClearFormats"}[what]
            return [f"{self.rng(pos[0])}.{method}"]
        if name == "FORMAT":
            target = self.rng(pos[0])
            out = []
            if "fill" in opts:
                f = opts["fill"]
                if isinstance(f, Token) and f.type == "COLOR":
                    out.append(f"{target}.Interior.Color = {color(f)}")
                elif str(f) == "none":
                    out.append(f"ForgeNoFill {target}")
                elif str(f) in S.FILL_COLORS:
                    out.append(f"{target}.Interior.Color = {color(S.FILL_COLORS[str(f)])}")
                else:
                    raise CompileError(f"unknown fill colour name {f}; use #RRGGBB")
            if "numfmt" in opts:
                out.append(f"{target}.NumberFormat = {e(opts['numfmt'])}")
            if "bold" in opts:
                out.append(f"{target}.Font.Bold = {'True' if str(opts['bold']) == 'yes' else 'False'}")
            if "width" in opts:
                w = opts["width"]
                out.append(f"{target}.EntireColumn.AutoFit" if str(w) == "auto" else f"{target}.ColumnWidth = {e(w)}")
            return out
        if name == "COPY_ROW":
            return [f"ForgeCopyRow {e(pos[0])}, {self.as_num(pos[1])[0]}, {e(pos[2])}, {self.as_num(pos[3])[0]}, {self.as_num(opts['cols'])[0]}"]
        if name == "DELETE_ROW":
            return [f"{e(pos[0])}.Rows({self.as_num(pos[1])[0]}).Delete"]
        if name == "INSERT_ROW":   # a blank row, whatever the row above holds (D36)
            return [f"ForgeInsertRow {e(pos[0])}, {self.as_num(pos[1])[0]}"]
        if name == "DICT_ADD":
            return [f'ForgeDictAdd {e(pos[0])}, {e(pos[1])}, {e(pos[2])}, "{opts["mode"]}"']
        if name == "DICT_GET":
            var = str(opts["into"])
            typ = self.vars[var]
            src = f"ForgeDictGet({e(pos[0])}, {e(pos[1])})"
            if typ == "String":
                src = f"ForgeText({src})"
            elif typ in ("Long", "Double"):
                src = f"ForgeNum({src})"
            return [f"{vba_name(var)} = {src}"]
        if name == "SORT":
            asc = "True" if str(opts.get("order", "asc")) == "asc" else "False"
            hdr = "True" if str(opts.get("header", "yes")) == "yes" else "False"
            return [f"ForgeSort {self.rng(pos[0])}, {self.as_num(opts['key'])[0]}, {asc}, {hdr}"]
        if name in ("ADD_SHEET", "COPY_SHEET"):   # VCL 0.4 (D40, D41): checks the name before it adds anything
            helper = "ForgeAddSheet" if name == "ADD_SHEET" else "ForgeCopySheet"
            args = f"{e(pos[0])}, {e(opts['name'])}, {e(opts['after']) if 'after' in opts else 'Nothing'}"
            if "into" in opts:
                return [f"Set {vba_name(str(opts['into']))} = {helper}({args})"]
            return [f"{helper} {args}"]
        if name == "SET_FORMULA":   # the English formula, as the validator checked it (D43)
            r1c1 = "True" if str(opts.get("style", "a1")) == "r1c1" else "False"
            return [f"ForgeSetFormula {self.rng(pos[0])}, {e(pos[1])}, {r1c1}"]
        if name == "FILTER":   # rows hidden by SAME, never AutoFilter (D42)
            return [f"ForgeFilter {self.rng(pos[0])}, {self.as_num(opts['col'])[0]}, {e(opts['crit'])}"]
        if name == "DEDUPE":   # S21 keys, whole rows deleted bottom up (D38)
            cols = opts["cols"]
            items = cols.children if isinstance(cols, Tree) and cols.data == "list_lit" else [cols]
            hdr = "True" if str(opts.get("header", "yes")) == "yes" else "False"
            return [f"ForgeDedupe {self.rng(pos[0])}, Array({', '.join(self.as_num(c)[0] for c in items)}), {hdr}"]
        if name == "READ_ARRAY":   # always a 2-D array, even for one cell (D36)
            return [f"{vba_name(str(opts['into']))} = ForgeReadArray({self.rng(pos[0])})"]
        if name == "WRITE_ARRAY":
            return [f"ForgeWriteArray {e(pos[0])}, {self.rng(opts['to'])}"]
        if name == "LIST_FILES":   # names in S24 order; never an error (D44)
            root = e(pos[0]) if isinstance(pos[0], Token) else self.path_root(pos[0])
            return [f"Set {vba_name(str(opts['into']))} = ForgeListFiles({root}, {e(pos[0])}, {e(pos[1])})"]
        if name in ("READ_TEXT", "WRITE_TEXT"):   # a status into a String VAR (E11, D45), never an error
            root, path = self.path_root(pos[0]), e(pos[0])
            if name == "READ_TEXT":
                return [f"{vba_name(str(opts['into']))} = ForgeReadText({root}, {path}, {vba_name(str(pos[1]))})"]
            text, k = self.ex(pos[1])
            text = text if k == "text" else f"ForgeText({text})"
            return [f"{vba_name(str(opts['into']))} = ForgeWriteText({root}, {path}, {text})"]
        if name in ("MAKE_FOLDER", "COPY_FILE"):   # a status into a String VAR (E11), never an error
            paths = ", ".join(f"{self.path_root(p)}, {e(p)}" for p in pos)
            helper = "ForgeMakeFolder" if name == "MAKE_FOLDER" else "ForgeCopyFile"
            return [f"{vba_name(str(opts['into']))} = {helper}({paths})"]
        if name == "RETURN":
            out = []
            if self.kind == "Function" and pos:
                out.append(f"{self.proc} = {self.ex(pos[0])[0] if self.ex(pos[0])[1] != 'raw' else self.as_num(pos[0])[0]}")
            out.append("GoTo CleanExit")
            return out
        if name == "LOG":
            return [f'ForgeInfo "{self.proc}", ForgeText({e(pos[0])})']
        if name == "MSG":
            return [f"MsgBox ForgeText({e(pos[0])})"]
        if name == "CALL":
            target = str(pos[0]).split(".")[-1]
            rest = ", ".join(e(p) for p in pos[1:])
            return [f"Call {target}({rest})" if rest else f"Call {target}"]
        raise CompileError(f"{name} is not supported by the MVP compiler")

    def path_root(self, t):
        """The $root of a file op's path: the validator allows only PATH_JOIN($root, ...) (E9)."""
        if isinstance(t, Tree) and t.data == "call" and str(t.children[0]) == "PATH_JOIN":
            return self.ex(pos_opts(t)[0][0])[0]
        raise CompileError(f"{canon(t)} is not PATH_JOIN($root, ...)")

    def update_facts_after(self, call):
        name = str(call.children[0])
        pos, opts = pos_opts(call)
        def drop(v):
            for k in [k for k in self.facts if mentions(k, v)]:
                del self.facts[k]
        if name == "LET":
            drop(str(pos[0]))
            _, k = self.ex(pos[1])
            if self.vars.get(str(pos[0])) == "Variant":
                self.facts[str(pos[0])] = k if k in ("num", "text", "date", "bool") else "raw"
        elif name in ("DICT_GET", "READ_ARRAY", "MAKE_FOLDER", "COPY_FILE", "LIST_FILES", "READ_TEXT",
                      "WRITE_TEXT") and "into" in opts:
            drop(str(opts["into"]))
            if name == "READ_TEXT":
                drop(str(pos[1]))
        elif name in ("SET_VALUE", "CLEAR", "SORT", "DELETE_ROW", "INSERT_ROW"):
            target = canon(pos[0])
            if target in self.facts:
                del self.facts[target]
            sheet = target[target.find("(") + 1:].split(",")[0].rstrip(")") if "(" in target else target
            if not target.startswith("CELL("):
                for k in [k for k in self.facts if k.startswith("CELL(" + sheet + ",")]:
                    del self.facts[k]
        elif name == "COPY_ROW" and len(pos) > 2:
            sheet = canon(pos[2])
            for k in [k for k in self.facts if k.startswith("CELL(" + sheet + ",")]:
                del self.facts[k]

    def with_facts(self, extra):
        saved = dict(self.facts)
        for k, v in extra.items():
            self.facts[k] = "date" if v == "Date" else "num"
        return saved

    def block(self, nodes, depth):
        for n in nodes:
            s, line, kind = n["stmt"], n["line"], n["kind"]
            if kind == "call":
                for text in self.action_lines(s):
                    self.emit(text, depth, line)
                self.update_facts_after(s)
            elif kind == "if_inline":
                cond, act = s.children
                if act.data == "note":
                    raise CompileError(f"line {line}: NOTE steps need the local model; the MVP compiler can't translate them")
                cc, _ = self.ex(cond)
                saved = self.with_facts(self.checker.facts_if(cond, True))
                body = self.action_lines(act)
                self.facts = saved
                if len(body) == 1:
                    self.emit(f"If {cc} Then {body[0]}", depth, line)
                else:
                    self.emit(f"If {cc} Then", depth, line)
                    for text in body:
                        self.emit(text, depth + 1, line)
                    self.emit("End If", depth)
                self.update_facts_after(act)
            elif kind == "if_block":
                cond = s.children[0]
                cc, _ = self.ex(cond)
                self.emit(f"If {cc} Then", depth, line)
                saved = self.with_facts(self.checker.facts_if(cond, True))
                self.block(n["children"], depth + 1)
                self.facts = saved
                if n["else"]:
                    self.emit("Else", depth)
                    saved = self.with_facts(self.checker.facts_if(cond, False))
                    self.block(n["else"], depth + 1)
                    self.facts = saved
                self.emit("End If", depth)
            elif kind in ("exit_if", "skip_if", "exit_loop_if"):
                cond = s.children[0]
                if kind == "exit_if":
                    jump = "GoTo CleanExit"
                elif kind == "exit_loop_if":
                    jump = "Exit For"
                else:   # VBA has no Continue: jump to a label just before the loop's Next
                    self.loops[-1][1] = True
                    jump = f"GoTo {self.loops[-1][0]}"
                self.emit(f"If {self.ex(cond)[0]} Then {jump}", depth, line)
                for k, v in self.checker.facts_if(cond, False).items():
                    self.facts[k] = "date" if v == "Date" else "num"
            elif kind in ("loop_rows", "loop_cols"):
                o = {str(x.children[0]): x.children[1] for x in s.children}
                var = str(o["var"])
                self.loopvars[var] = "num"
                step = " Step -1" if str(o.get("dir", "down")) in ("up", "left") else ""
                self.emit(f"For {vba_name(var)} = {self.as_num(o['from'])[0]} To {self.as_num(o['to'])[0]}{step}", depth, line)
                saved = {k: v for k, v in self.facts.items()}
                self.facts = {k: v for k, v in self.facts.items() if not mentions(k, var)}
                self.loop_body(n["children"], depth + 1, line)
                self.facts = {k: v for k, v in saved.items() if not mentions(k, var)}
                self.emit(f"Next {vba_name(var)}", depth)
            elif kind == "for_each":
                o = {str(x.children[0]): x.children[1] for x in s.children}
                var = str(o["var"])
                src = o["in"]
                fn = str(src.children[0])
                spos, _ = pos_opts(src)
                if fn == "KEYS":
                    coll = f"{self.ex(spos[0])[0]}.Keys"
                elif fn == "ITEMS":
                    coll = self.ex(spos[0])[0]
                elif fn == "SHEETS":
                    coll = f"{self.ex(spos[0])[0]}.Worksheets"
                elif fn == "CELLS":
                    coll = f"{self.rng(spos[0])}.Cells"
                else:
                    raise CompileError(f"line {line}: FOR_EACH over {fn} is not supported")
                self.loopvars[var] = {"KEYS": "text", "SHEETS": "obj", "ITEMS": "text"}.get(fn, "raw")
                # over CELLS the loop variable is the cell's value, as in the interpreter (D36)
                walker = f"forgeCell_{var}" if fn == "CELLS" else vba_name(var)
                if fn == "CELLS":
                    self.loopvars[walker] = "obj"
                self.emit(f"For Each {walker} In {coll}", depth, line)
                if fn == "CELLS":
                    self.emit(f"{vba_name(var)} = {walker}.Value", depth + 1, line)
                saved = dict(self.facts)
                self.facts = {k: v for k, v in self.facts.items() if not mentions(k, var)}
                self.loop_body(n["children"], depth + 1, line)
                self.facts = {k: v for k, v in saved.items() if not mentions(k, var)}
                self.emit(f"Next {walker}", depth)
            elif kind == "select_case":   # each CASE value is compared as = does (S7), not by VBA (D36)
                sel = f"forgeSel{len(self.temps) + 1}"   # the selector is evaluated once
                code, self.temps[sel] = self.ex(s.children[0])
                self.emit(f"{sel} = {code}", depth, line)
                self.emit("Select Case True", depth, line)
                self.selectors.append(Token("NAME", sel))
                self.block(n["children"], depth + 1)
                self.selectors.pop()
                self.emit("End Select", depth)
            elif kind == "case_line":
                tests = [self.ex_compare(Tree("comparison", [self.selectors[-1], Token("COMP_OP", "="), c]))[0]
                         for c in s.children]
                self.emit("Case " + ", ".join(tests), depth, line)
                self.block(n["children"], depth + 1)
            elif kind == "case_else":
                self.emit("Case Else", depth)
                self.block(n["children"], depth + 1)
            elif kind == "note":
                raise CompileError(f"line {line}: NOTE steps need the local model; the MVP compiler can't translate them")
            else:
                raise CompileError(f"line {line}: {kind} is not supported by the MVP compiler")

    def loop_body(self, nodes, depth, line):
        """A loop's body, with the label SKIP_IF jumps to (Next_<spec line>) when it is used."""
        self.loops.append([f"Next_{line}", False])
        self.block(nodes, depth)
        label, used = self.loops.pop()
        if used:
            self.lines.append(f"{label}:")

    def compile(self):
        body_start = len(self.lines)
        # body first (to discover loop variables), then assemble around it
        self.block(self.spec["ops"], 1)
        body = self.lines[body_start:]
        self.lines = []
        params = ", ".join(f"ByVal {vba_name(n)} As {DIM_TYPE[t]}" for n, t in self.args)
        head = f"Public {self.kind} {self.proc}({params})" + (f" As {DIM_TYPE[self.ret]}" if self.kind == "Function" and self.ret else "")
        out = [f"'@VCL PROC={self.proc}  source: {self.source}", head]
        for name, typ in self.vars.items():
            out.append(f"{'':<6}    Dim {vba_name(name)} As {DIM_TYPE[typ]}")
        for name, k in self.loopvars.items():
            if name not in self.vars:
                out.append(f"{'':<6}    Dim {vba_name(name)} As {'Long' if k == 'num' else 'Variant'}")
        for name in self.temps:
            out.append(f"{'':<6}    Dim {name} As Variant")
        out += [f"{'':<6}    Dim forgeScreen As Boolean", f"{'':<6}    Dim forgeCalc As Long",
                f"{'':<6}    Dim forgeEvents As Boolean",
                f"{'':<6}    forgeScreen = Application.ScreenUpdating",
                f"{'':<6}    forgeCalc = Application.Calculation",
                f"{'':<6}    forgeEvents = Application.EnableEvents",
                f"{'':<6}    On Error GoTo ErrHandler"]
        for name, typ in self.vars.items():
            if typ == "Dictionary":
                out.append(f"{'':<6}    Set {vba_name(name)} = ForgeNewDict()")
            elif typ == "Collection":
                out.append(f"{'':<6}    Set {vba_name(name)} = New Collection")
        self.lines = []
        self.label_k = {}
        for line, text, conds in self.pre:
            for i, cond in enumerate(conds):
                cc, _ = self.ex(cond)
                msg = f"PRE failed: {canon(cond)}".replace('"', '""')
                fail = f"{self.proc} = -1: " if self.kind == "Function" else ""
                self.emit(f'If Not {cc} Then ForgeLog "{self.proc}", 0, "{msg}", {line * 10}: {fail}GoTo CleanExit', 1, line)
        out += self.lines
        perf = {"screen_off": "Application.ScreenUpdating = False",
                "calc_manual": "Application.Calculation = xlCalculationManual",
                "events_off": "Application.EnableEvents = False"}
        out += [f"{'':<6}    {perf[p]}" for p in self.perf]
        out += body
        end = "Function" if self.kind == "Function" else "Sub"
        out += ["CleanExit:", f"{'':<6}    ForgeRestore forgeScreen, forgeCalc, forgeEvents",
                f"{'':<6}    Exit {end}", "ErrHandler:",
                f'{"":<6}    ForgeLog "{self.proc}", Err.Number, Err.Description, Erl']
        if self.kind == "Function":
            out.append(f"{'':<6}    {self.proc} = -2")
        out += [f"{'':<6}    Resume CleanExit", f"End {end}"]
        return "\n".join(line.rstrip() for line in out) + "\n"


def compile_text(text, source_name=""):
    spec, problems = parse(text)
    if problems:
        raise CompileError(f"spec doesn't parse: {problems[:3]}")
    if spec["kind"] != "TASK":
        raise CompileError("only TASK specs can be compiled by the MVP compiler")
    found = Checker(spec).run()
    if found:
        raise CompileError(f"spec doesn't type-check: {found[:3]}")
    return Gen(spec, source_name).compile()


FORGELIB = HERE.parent / "vba" / "ForgeLib.bas"


def module_title(n):
    return f"VBA Forge module: {n} procedures"


def assemble(procedures, title="VBA Forge module"):
    """One paste-ready module: header, Option Explicit, ForgeLib, then the procedures."""
    head = [f"' {title}", "' Generated by vclc. Paste into ONE new module (Insert > Module). Do not edit by hand.",
            "Option Explicit", ""]
    return "\n".join(head) + FORGELIB.read_text(encoding="utf-8") + "\n" + "\n".join(procedures)
