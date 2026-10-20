"""VCL 0.3 (docs/vep-0.3-file-ops.md). Router Phase 1a: the version line, SAFE_NAME, PATH_JOIN
and FOLDER_EXISTS, and the path rules of the type checker (T4). Phase 1b: MAKE_FOLDER and
COPY_FILE (E9), MODIFIES $root files (E10) and folder FIXTUREs (F6). Phase 2: the mock file
system, the file op statuses (E11), FILE_EXISTS and FILE_COUNT (F7) and UNCHANGED for files.
Phase 3a: the compiler, ForgeLib's file helpers and the linter's file safety rules (D29). Phase 3b:
the sandbox builder for Excel and the sandbox length check (D30)."""
import glob
import inspect
import os
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
import semantics  # noqa: E402
import validator  # noqa: E402
import vcl_interp  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_fixtures import (EXCEL_SANDBOX_MAX, SANDBOX, MockFS, build_folders, default_size,  # noqa: E402
                          fixtures_module)
from vcl_interp import (VclError, copy_file, file_unchanged_problems, make_folder, modifies_of,  # noqa: E402
                        sandbox_length_problems)
from vcl_mutate import mutate_report  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import check_text  # noqa: E402

# Job IDs typed by hand: two good, one padded, three unsafe (.., CON, A/B).
NAMES = r"""VCL 0.3
TASK
# ASSUME: Jobs headers (row 1) include "Job ID", "Safe Name" and "Path"
PROC     CheckJobNames
KIND     Function
MODULE   modJobs
ARGS     ws As Worksheet, root As String
RET      Long
VAR      cJob As Long
VAR      cName As Long
VAR      cPath As Long
VAR      n As Long
PRE      COL($ws, "Job ID") > 0; COL($ws, "Safe Name") > 0; COL($ws, "Path") > 0
MODIFIES $ws["Safe Name", "Path"]
OP       LET(cJob, COL($ws, "Job ID"))
OP       LET(cName, COL($ws, "Safe Name"))
OP       LET(cPath, COL($ws, "Path"))
OP       LET(n, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SET_VALUE(CELL($ws, r, cName), SAFE_NAME(CELL($ws, r, cJob)))
OP         SET_VALUE(CELL($ws, r, cPath), PATH_JOIN($root, SAFE_NAME(CELL($ws, r, cJob)), "Reports"))
OP         IF IS_EMPTY(CELL($ws, r, cName)) THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Safe Name,C:Path rows=6 data=A:values("J-101"x2," J-102 "x1,".."x1,"CON"x1,"A/B"x1)
TEST     fixture=jobs_01.xlsx call=CheckJobNames(Jobs, "C:\Jobs") runs=2 assert=RESULT = 3 AND COUNTIF(Jobs![Path], "C:\Jobs\J-101\Reports") = 2 AND COUNTIF(Jobs![Safe Name], "J-102") = 1 AND ERROR_COUNT = 0
TEST     fixture=jobs_01.xlsx call=CheckJobNames(Jobs, "C:\Jobs") assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
"""

# Router Phase 1b: a small Job-Pack Router. Each tracker row names a job and its report; the
# report is copied from the dump folder into the job's folder, and the status written back.
ROUTER = r"""VCL 0.3
TASK
# ASSUME: Tracker headers (row 1) include "Job ID", "Report File" and "Status"
PROC     RouteJobPacks
KIND     Function
MODULE   modRouter
ARGS     tracker As Worksheet, dump As String, jobs As String
RET      Long
VAR      cJob As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      n As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Report File") > 0; COL($tracker, "Status") > 0; FOLDER_EXISTS($dump); FOLDER_EXISTS($jobs)
MODIFIES $tracker["Status"]
MODIFIES $jobs files
OP       LET(cJob, COL($tracker, "Job ID"))
OP       LET(cFile, COL($tracker, "Report File"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LET(n, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cJob))
OP         IF SAFE_NAME(CELL($tracker, r, cJob)) = "" OR SAFE_NAME(CELL($tracker, r, cFile)) = ""
OP           SET_VALUE(CELL($tracker, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=status)
OP           IF status = "Created" OR status = "Exists"
OP             COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))), PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Copied" OR status = "Already there" THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  router_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=7 data=A:values("J-101"x1,"J-102"x1,"J-103"x1,"..\x"x1,"J-104"x1,""x1,"J-105"x1)+boundary("J-101","J-102","J-103","..\x","J-104","","J-105");B:values("r1.pdf"x2,"r2.pdf"x1,"r3.pdf"x1,"r4.pdf"x1,"r5.pdf"x1,"r9.pdf"x1)+boundary("r1.pdf","r2.pdf","r3.pdf","r4.pdf","r1.pdf","r5.pdf","r9.pdf")
FIXTURE  router_01.xlsx folder=Dump files=r1.pdf:120,r2.pdf:80,r3.pdf locked=r3.pdf
FIXTURE  router_01.xlsx folder=Jobs files=J-102\r2.pdf:99,J-104:5
TEST     fixture=router_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Copied") = 1 AND FILE_EXISTS("Jobs\J-101\r1.pdf") AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 1
TEST     fixture=router_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Already there") = 1 AND COUNTIF(Tracker![Status], "Conflict") = 1 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Bad name") = 1 AND COUNTIF(Tracker![Status], "Blocked by file") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND FILE_COUNT("Jobs") = 3
TEST     fixture=router_01.xlsx call=RouteJobPacks(Tracker, Dump, "S:\Missing") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID]) AND FILE_COUNT("Dump") = 3
"""


def router(*edits):
    """ROUTER with each (old, new) edit applied once."""
    text = ROUTER
    for old, new in edits:
        assert text.count(old) == 1, old
        text = text.replace(old, new)
    return text


def problems(text):
    return [(p.severity, p.message) for p in validator.validate_text(text)]


HEAD = ["VCL 0.3", "TASK", "PROC     P", "KIND     Function", "MODULE   modP",
        "ARGS     ws As Worksheet, root As String", "RET      Long", "VAR      n As Long",
        "VAR      t As String", "VAR      flag As Boolean"]


def spec(*ops):
    return "\n".join(HEAD + ["OP       " + o for o in ops] + ["OP       RETURN(n)", "ERR      standard"])


def errors(text):
    return [p.message for p in validator.validate_text(text) if p.severity == "error"]


def type_problems(text):
    return [m for _, m in check_text(text)]


class VersionLine(unittest.TestCase):
    def test_both_parsers_read_it(self):
        self.assertEqual(parse(NAMES)[0]["version"], "VCL 0.3")
        self.assertEqual(validator.parse_spec(NAMES).version, "VCL 0.3")

    def test_every_vcl_02_spec_stays_valid_under_0_3(self):
        texts = {os.path.basename(p): Path(p).read_text(encoding="utf-8")
                 for p in sorted(glob.glob(os.path.join(ROOT, "specs", "*.vcl")))}
        self.assertEqual(sorted(n for n, t in texts.items() if not t.startswith("VCL 0.2\n")), ["task_12.vcl"])
        for name, text in texts.items():
            if name == "task_12.vcl":   # the router, written for VCL 0.3 (router Phase 4)
                continue
            with self.subTest(spec=name):
                newer = text.replace("VCL 0.2", "VCL 0.3", 1)
                self.assertEqual([(p.line, p.severity, p.message) for p in validator.validate_text(text)],
                                 [(p.line, p.severity, p.message) for p in validator.validate_text(newer)])
                self.assertEqual([(r.passed, r.result, r.state) for r in run_tests(text)],
                                 [(r.passed, r.result, r.state) for r in run_tests(newer)])

    def test_the_new_functions_need_the_0_3_line(self):
        older = NAMES.replace("VCL 0.3", "VCL 0.2", 1)
        for name in ("SAFE_NAME", "PATH_JOIN"):
            self.assertIn(f"{name} is VCL 0.3: start the spec with 'VCL 0.3' to use it", errors(older))
        self.assertIn("FOLDER_EXISTS is VCL 0.3: start the spec with 'VCL 0.3' to use it",
                      errors(spec("IF FOLDER_EXISTS($root) THEN LET(n, 1)").replace("VCL 0.3", "VCL 0.2", 1)))


class Validator(unittest.TestCase):
    def test_the_example_is_valid(self):
        self.assertEqual([(p.line, p.message) for p in validator.validate_text(NAMES)], [])

    def test_argument_counts(self):
        self.assertIn("PATH_JOIN takes 2 or more arguments, got 1",
                      errors(spec("LET(t, PATH_JOIN($root))")))
        self.assertIn("SAFE_NAME takes 1 argument, got 2", errors(spec('LET(t, SAFE_NAME("a", "b"))')))
        # The helper has no TESTs (rule 5); only the call's own problems matter here.
        self.assertEqual(errors(spec('LET(t, PATH_JOIN($root, "Jobs", "J-101", "r1.pdf"))')),
                         ["a spec needs at least two TEST lines (rule 5)"])


class PathTypes(unittest.TestCase):
    """T4: a path is a String $argument or PATH_JOIN($root, part, ...) with every part
    SAFE_NAME(...) or a quoted safe name, written in place."""

    def test_good_paths(self):
        for op in ('LET(t, PATH_JOIN($root, SAFE_NAME(CELL($ws, 2, 1))))',
                   'LET(t, PATH_JOIN($root, "Reports", SAFE_NAME(CELL($ws, 2, 1)), "r1.pdf"))',
                   'IF FOLDER_EXISTS($root) THEN LET(n, 1)',
                   'IF FOLDER_EXISTS(PATH_JOIN($root, "Reports")) THEN LET(n, 1)',
                   'SET_VALUE(CELL($ws, 2, 2), SAFE_NAME(CELL($ws, 2, 1)))'):
            with self.subTest(op=op):
                self.assertEqual(type_problems(spec(op)), [])

    def test_a_path_must_start_from_a_string_argument(self):
        for root in ('"C:\\Jobs"', "t", "$ws", "TRIM(t)"):
            with self.subTest(root=root):
                found = type_problems(spec(f'LET(t, PATH_JOIN({root}, "Reports"))'))
                self.assertTrue(any("must start from a String $argument" in m for m in found), found)

    def test_every_part_is_checked_where_it_is_used(self):
        for part, why in (('".."', "not a safe name"), ('"A\\B"', "not a safe name"), ('"CON"', "not a safe name"),
                          ("t", "must be SAFE_NAME(...)"), ("CELL($ws, 2, 1)", "must be SAFE_NAME(...)"),
                          ('TRIM(CELL($ws, 2, 1))', "must be SAFE_NAME(...)")):
            with self.subTest(part=part):
                found = type_problems(spec(f"LET(t, PATH_JOIN($root, {part}))"))
                self.assertTrue(any(why in m for m in found), found)

    def test_folder_exists_takes_a_path(self):
        for arg in ("t", '"C:\\Jobs"', "SAFE_NAME(t)"):
            with self.subTest(arg=arg):
                found = type_problems(spec(f"IF FOLDER_EXISTS({arg}) THEN LET(n, 1)"))
                self.assertTrue(any("is not a path" in m for m in found), found)

    def test_names_and_paths_are_read_as_text(self):
        for op in ('IF SAFE_NAME(t) = "" THEN LET(n, 1)', 'IF SAFE_NAME(t) <> t THEN LET(n, 1)',
                   'IF PATH_JOIN($root, "a") > t THEN LET(n, 1)', 'LET(t, FILE_NAME(PATH_JOIN($root, "a")))',
                   'LET(t, SAFE_NAME(t) & "x")', 'LET(t, IIF(flag, SAFE_NAME(t), "x"))',
                   'LET(n, LEN(PATH_JOIN($root, "a")))', 'LET(flag, FOLDER_EXISTS($root))'):
            with self.subTest(op=op):
                self.assertEqual(type_problems(spec(op)), [])
        self.assertIn("can't compare Name with Num", type_problems(spec('IF SAFE_NAME(t) = 1 THEN LET(n, 1)')))

    def test_names_and_paths_are_text_but_not_numbers(self):
        self.assertTrue(any("expected a number" in m or "not a number" in m
                            for m in type_problems(spec("LET(n, MOD(SAFE_NAME(t), 2))"))))
        self.assertTrue(any("expected text" in m for m in type_problems(spec("LET(t, SAFE_NAME($ws))"))))


MAKE = "MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=status)"
COPY_FROM = "COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))),"
RUN = "call=RouteJobPacks(Tracker, Dump, Jobs) assert="


class FileOps(unittest.TestCase):
    """E9: MAKE_FOLDER(<path>, into=<status>) and COPY_FILE(<from>, <to>, into=<status>)."""

    def test_the_router_is_valid(self):
        self.assertEqual(problems(ROUTER), [])

    def test_everything_new_needs_the_0_3_line(self):
        found = errors(ROUTER.replace("VCL 0.3", "VCL 0.2", 1))
        for message in ("MAKE_FOLDER is VCL 0.3: start the spec with 'VCL 0.3' to use it",
                        "COPY_FILE is VCL 0.3: start the spec with 'VCL 0.3' to use it",
                        "MODIFIES ... files is VCL 0.3: start the spec with 'VCL 0.3' to use it",
                        "folder FIXTUREs are VCL 0.3: start the spec with 'VCL 0.3' to use them"):
            self.assertIn(message, found)

    def test_the_status_goes_into_a_string_var(self):
        self.assertIn("MAKE_FOLDER needs into=", errors(router((MAKE, MAKE.replace(", into=status", "")))))
        self.assertIn("type: MAKE_FOLDER: into=n is Num, but the status is text: declare it As String",
                      errors(router((MAKE, MAKE.replace("into=status", "into=n")))))

    def test_paths_are_paths(self):
        found = errors(router((COPY_FROM, "COPY_FILE(CELL($tracker, r, cFile),")))
        self.assertIn("type: COPY_FILE: CELL($tracker,r,cFile) is not a path: write a String $argument or "
                      "PATH_JOIN($root, ...)", found)

    def test_a_root_itself_is_never_made_or_copied(self):
        self.assertIn("MAKE_FOLDER: the folder to make must be PATH_JOIN($jobs, ...), not the root $jobs itself, "
                      "which must already exist", errors(router((MAKE, "MAKE_FOLDER($jobs, into=status)"))))
        self.assertIn("COPY_FILE: the file to copy from must be PATH_JOIN($dump, ...), not the root $dump itself, "
                      "which must already exist", errors(router((COPY_FROM, "COPY_FILE($dump,"))))

    def test_every_root_is_checked_in_pre(self):
        self.assertEqual(errors(router(("; FOLDER_EXISTS($dump)", ""))),
                         ["file ops use $dump, but PRE doesn't check FOLDER_EXISTS($dump) (rule 14)"])


class ModifiesFiles(unittest.TestCase):
    """E10: MODIFIES $root files, for a String argument, is the only way to add below a root."""

    def test_adding_below_an_undeclared_root(self):
        self.assertEqual(errors(router(("MODIFIES $jobs files\n", ""))),
                         ["MAKE_FOLDER adds to $jobs, but no line declares MODIFIES $jobs files",
                          "COPY_FILE adds to $jobs, but no line declares MODIFIES $jobs files"])

    def test_the_line_must_fit_the_argument(self):
        self.assertIn("MODIFIES $tracker files: $tracker must be a String argument (a root folder)",
                      errors(router(("MODIFIES $jobs files", "MODIFIES $tracker files"))))
        self.assertIn("MODIFIES $jobs: a String argument is a root folder, so write MODIFIES $jobs files",
                      errors(router(("MODIFIES $jobs files", "MODIFIES $jobs"))))
        self.assertIn("MODIFIES $jobs files appears twice: use one line",
                      errors(router(("MODIFIES $jobs files", "MODIFIES $jobs files\nMODIFIES $jobs files"))))

    def test_a_copy_source_needs_no_line_but_an_unused_one_warns(self):
        self.assertEqual(problems(router(("MODIFIES $jobs files", "MODIFIES $jobs files\nMODIFIES $dump files"))),
                         [("warning", "MODIFIES $dump files, but no op adds to $dump")])

    def test_the_interpreter_reads_the_line(self):
        from vcl_parse import parse as grammar_parse
        self.assertEqual(modifies_of(grammar_parse(ROUTER)[0]), {"tracker": {"cols": ["Status"]}, "jobs": {"files": True}})


class FolderFixtures(unittest.TestCase):
    """F6: FIXTURE <file>.xlsx folder=<Name> [files=<path>[:<bytes>],...] [locked=<path>,...]."""

    def test_the_parser_reads_them(self):
        dump, jobs = validator.parse_spec(ROUTER).folders
        self.assertEqual((dump.file, dump.folder, dump.files, dump.locked),
                         ("router_01.xlsx", "Dump", [("r1.pdf", 120, None), ("r2.pdf", 80, None), ("r3.pdf", None, None)],
                          ["r3.pdf"]))
        self.assertEqual((jobs.folder, jobs.files, jobs.locked),
                         ("Jobs", [("J-102\\r2.pdf", 99, None), ("J-104", 5, None)], []))
        self.assertEqual(problems(router(("folder=Jobs files=J-102\\r2.pdf:99,J-104:5", "folder=Jobs"))), [])

    def test_bad_lines(self):
        cases = [
            (("files=r1.pdf:120,", "files=CON.pdf,"), "FIXTURE files: CON.pdf is not a safe path: CON.pdf is not a "
                                                      "safe name (see SAFE_NAME)"),
            (("files=r1.pdf:120,", "files=a\\..\\b,"), "FIXTURE files: a\\..\\b is not a safe path: .. is not a "
                                                       "safe name (see SAFE_NAME)"),
            (("files=r1.pdf:120,", "files=r1.pdf:120,R1.PDF,"), "FIXTURE folder Dump lists R1.PDF twice"),
            (("files=J-102\\r2.pdf:99", "files=J-102:5,J-102\\r2.pdf:99"),
             "FIXTURE folder Jobs: J-102 is a file, so nothing can be below it"),
            ((" locked=r3.pdf", " locked=r9.pdf"), "FIXTURE locked= r9.pdf is not one of the folder's files="),
            ((" locked=r3.pdf", " locked=r3.pdf,r3.pdf"), "FIXTURE folder Dump lists a locked file twice"),
            (("folder=Jobs", "folder=Dump"), "router_01.xlsx has two FIXTURE lines for folder Dump"),
            (("folder=Jobs", "folder=Tracker"), "router_01.xlsx has a sheet and a folder named Tracker: a TEST "
                                                "could not tell them apart"),
            (("files=r1.pdf:120,", "files=r 1.pdf:120,"), "a folder FIXTURE must look like: <name>.xlsx "
                                                          "folder=<Name> [files=<path>[:<bytes>],...] "
                                                          "[locked=<path>,...]"),
            (("files=r1.pdf:120,", "files=r1.pdf:big,"), 'FIXTURE files: expected <path>, <path>:<bytes> or '
                                                         '<path>:"<text>", got \'r1.pdf:big\''),
        ]
        for edit, message in cases:
            with self.subTest(edit=edit):
                self.assertIn(message, errors(router(edit)))

    def test_tests_pass_folders_bare(self):
        self.assertIn("TEST passes Jobz for $jobs, but router_01.xlsx has no folder Jobz",
                      errors(router((RUN, RUN.replace("Jobs)", "Jobz)")))))
        found = errors(router((RUN, "call=RouteJobPacks(Dump, Tracker, Jobs) assert=")))
        self.assertIn("TEST passes Dump for $tracker, but router_01.xlsx has no sheet Dump", found)
        self.assertIn("TEST passes Tracker for $dump, but router_01.xlsx has no folder Tracker", found)

    def test_a_missing_root_fails_pre(self):
        self.assertEqual(errors(router(("RESULT = -1 AND ERROR_COUNT = 1 AND", "ERROR_COUNT = 0 AND"))),
                         ["TEST fails PRE ($jobs is \"S:\\Missing\", which is not a folder of router_01.xlsx) but "
                          "doesn't expect RESULT = -1"])

    def test_rule_19_wants_a_locked_file(self):
        self.assertEqual(problems(router((" locked=r3.pdf", ""))),
                         [("warning", "the spec copies files, but no FIXTURE folder has a locked= file, so no TEST "
                                      "shows what happens to a file that is open elsewhere (rule 19)")])



def fixture_trees(text):
    return [k["tree"] for k in parse(text)[0]["keys"] if k["key"] == "FIXTURE"]


class SandboxBuilder(unittest.TestCase):
    """Router Phase 3b: ForgeSandbox_<file>() builds the folder FIXTUREs in Excel (D30)."""

    def test_it_builds_what_the_interpreter_tests_on(self):
        m = fixtures_module([("task_x", fixture_trees(ROUTER))])
        self.assertEqual(lint(m), [])
        self.assertIn("Public Sub ForgeFixture_router_01(ByVal wb As Workbook)", m)   # the sheets, as before
        body = m[m.index("Public Function ForgeSandbox_router_01() As String"):]
        body = body[:body.index("End Function")]
        steps = [line.strip() for line in body.splitlines()[2:]]
        self.assertEqual(steps[0], "root = SandboxNewRoot()")
        self.assertEqual(steps[-1], "ForgeSandbox_router_01 = root")
        built = {}
        for step in steps[1:-1]:
            kind, path, *size = re.fullmatch(r'(Sandbox\w+) root & "\\(.+?)"(?:, (\d+))?', step).groups()
            built.setdefault(kind, []).append((path, int(size[0])) if size[0] else path)
        mock = build_folders(fixture_trees(ROUTER))["router_01.xlsx"]
        prefix = len(SANDBOX) + 1
        folders = [e["path"][prefix:] for e in mock.entries.values() if e["kind"] == "folder" and e["path"] != SANDBOX]
        self.assertEqual(sorted(built["SandboxFolder"]), sorted(folders))
        self.assertEqual(sorted(built["SandboxFile"]),
                         sorted((e["path"][prefix:], e["size"]) for e in mock.entries.values() if e["kind"] == "file"))
        self.assertEqual(built["SandboxLock"], [e["path"][prefix:] for e in mock.entries.values() if e["locked"]])
        self.assertLess(steps.index('SandboxFolder root & "\\Jobs"'), steps.index('SandboxFolder root & "\\Jobs\\J-102"'))
        self.assertGreater(min(i for i, x in enumerate(steps) if x.startswith("SandboxLock")),
                           max(i for i, x in enumerate(steps) if x.startswith("SandboxFile")))

    def test_the_root_limits_come_from_the_interpreter(self):
        m = fixtures_module([("task_x", fixture_trees(ROUTER))])
        self.assertIn(f"Private Const SANDBOX_MAX_ROOT As Long = {EXCEL_SANDBOX_MAX}", m)
        self.assertIn(f"Private Const SANDBOX_TEST_ROOT As Long = {len(SANDBOX)}", m)
        self.assertIn('Environ$("TEMP") & "\\KGC_KRONOS_SANDBOX_" & Format$(Now, "yyyymmdd_hhnnss")', m)
        self.assertIn('root = stem & "_" & n', m)   # several sandboxes a second each get a new folder

    def test_it_never_writes_over_anything(self):
        m = fixtures_module([("task_x", fixture_trees(ROUTER))])
        body = m[m.index("Private Sub SandboxFile("):]
        body = body[:body.index("End Sub")]
        self.assertLess(body.index("If SandboxFso().FileExists(path) Or SandboxFso().FolderExists(path) Then"),
                        body.index("Open path For Output"))
        self.assertIn("SandboxFso().CreateFolder path", m)   # error 58 when the folder is there
        self.assertNotRegex(m, r"(?i)\b(Kill|DeleteFile|DeleteFolder|MoveFile|RmDir)\b")

    def test_bad_fixtures_stop_the_build(self):
        long_name = "a" * 100
        for line, why in ((r"FIXTURE  router_01.xlsx folder=Jobs files=J-102\r2.pdf:99,J-104:10000001", "bytes"),
                          (rf"FIXTURE  router_01.xlsx folder=Jobs files={long_name}\{long_name}.pdf", "too long")):
            with self.subTest(why):
                text = router((r"FIXTURE  router_01.xlsx folder=Jobs files=J-102\r2.pdf:99,J-104:5", line))
                with self.assertRaisesRegex(ValueError, why):
                    fixtures_module([("task_x", fixture_trees(text))])

    def test_a_shared_fixture_file_must_build_the_same_folders(self):
        other = router(("files=r1.pdf:120,", "files=r1.pdf:121,"))
        with self.assertRaisesRegex(ValueError, "defined differently"):
            fixtures_module([("task_x", fixture_trees(ROUTER)), ("task_y", fixture_trees(other))])
        self.assertIn("ForgeSandbox_router_01", fixtures_module([("task_x", fixture_trees(ROUTER)),
                                                                 ("task_y", fixture_trees(ROUTER))]))

    def test_lint_keeps_file_writes_in_the_sandbox_helpers(self):
        m = fixtures_module([("task_x", fixture_trees(ROUTER))])
        faults = {"Open elsewhere": m.replace("    root = SandboxNewRoot()\n",
                                              '    root = SandboxNewRoot()\n    Open "x" For Append As #1\n', 1),
                  "SandboxFso elsewhere": m.replace("    root = SandboxNewRoot()\n",
                                                    "    root = SandboxNewRoot()\n    SandboxFso().CreateFolder root\n", 1),
                  "two procedures, one name": m.replace("Private Sub SandboxLock(", "Private Sub SandboxFile(", 1)}
        for what, bad in faults.items():
            with self.subTest(what):
                self.assertNotEqual(bad, m)
                self.assertTrue(lint(bad))


SHORT = "a" * 50
LONG = "a" * 100


def deep_folder(*parts):
    """A spec that makes one folder below $jobs from the given quoted parts."""
    return "\n".join(["VCL 0.3", "TASK", "PROC     P", "KIND     Function", "MODULE   modP", "ARGS     jobs As String",
                      "RET      Long", "VAR      t As String", "VAR      n As Long", "PRE      FOLDER_EXISTS($jobs)",
                      "MODIFIES $jobs files",
                      "OP       MAKE_FOLDER(PATH_JOIN($jobs, " + ", ".join(f'"{p}"' for p in parts) + "), into=t)",
                      "OP       RETURN(n)", "ERR      standard", "FIXTURE  deep_01.xlsx folder=Jobs",
                      "TEST     fixture=deep_01.xlsx call=P(Jobs) assert=RESULT = 0 AND ERROR_COUNT = 0"])


class SandboxLength(unittest.TestCase):
    """D28's gap, closed in 3b (D30): a TEST fails when the longer Excel root could change a status."""

    def test_the_rule(self):
        jobs = SANDBOX + "\\Jobs\\"
        near = jobs + "a" * (247 - len(jobs))   # exactly at the folder limit here
        self.assertEqual(sandbox_length_problems([(jobs + "J-101", 247)]), [])
        self.assertEqual(sandbox_length_problems([(near + "a" * EXCEL_SANDBOX_MAX, 247)]), [])   # Too long anywhere
        self.assertEqual(sandbox_length_problems([("S:\\Missing\\" + "a" * 240, 247)]), [])  # not in the sandbox
        self.assertEqual(len(sandbox_length_problems([(near, 247), (near, 247)])), 1)
        self.assertIn("has 247 characters here but up to 350", sandbox_length_problems([(near, 247)])[0])

    def test_it_fails_the_test(self):
        for parts, passes in (((SHORT,), True), ((LONG, LONG), False), ((LONG, LONG, LONG), True)):
            with self.subTest(len("\\".join(parts))):
                (result,) = run_tests(deep_folder(*parts))
                self.assertEqual(result.passed, passes, result.failed)
                if not passes:
                    self.assertEqual([p for p in result.failed if isinstance(p, str) and p.startswith("SANDBOX:")],
                                     [f"SANDBOX: Jobs\\{LONG}\\{LONG} has 224 characters here but up to 327 in "
                                      "Excel's sandbox, where it would be Too long (over 247): make it far longer or "
                                      "shorter"])


def mock(*files, folders=(), locked=()):
    """A MockFS with the sandbox's Jobs folder and the given folders and (path, size) files below it."""
    fs = MockFS()
    fs.add_folder(JOBS)
    for f in folders:
        fs.add_folder(JOBS + "\\" + f)
    for f, size in files:
        fs.add_file(JOBS + "\\" + f, size, f in locked)
    return fs


JOBS = SANDBOX + "\\Jobs"


class MakeFolder(unittest.TestCase):
    """E11: MAKE_FOLDER's statuses on the mock file system."""

    def test_statuses(self):
        cases = [
            (mock(), "", "Bad name"),
            (mock(), JOBS + "\\J*1", "Bad name"),
            (mock(), JOBS + "\\" + "x" * 300, "Too long"),
            (mock(folders=["J-101"]), JOBS + "\\J-101", "Exists"),
            (mock(("J-101", 5)), JOBS + "\\J-101", "Blocked by file"),
            (mock(("J-101", 5)), JOBS + "\\J-101\\Reports", "Blocked by file"),
            (mock(), JOBS + "\\J-101", "Created"),
        ]
        for fs, path, status in cases:
            with self.subTest(path=path, status=status):
                self.assertEqual(make_folder(fs, JOBS, path), status)

    def test_it_makes_every_missing_level_below_the_root_only(self):
        fs = mock()
        self.assertEqual(make_folder(fs, JOBS, JOBS + "\\J-101\\Reports\\NDE"), "Created")
        for level in ("J-101", "J-101\\Reports", "J-101\\Reports\\NDE"):
            self.assertTrue(fs.is_folder(JOBS + "\\" + level), level)
        fs = MockFS()   # the root itself is never made
        self.assertEqual(make_folder(fs, JOBS, JOBS + "\\J-101"), "No folder")
        self.assertFalse(fs.is_folder(JOBS))

    def test_a_blocked_level_makes_nothing(self):
        fs = mock(("J-101\\Reports", 5), folders=["J-101"])
        before = dict(fs.entries)
        self.assertEqual(make_folder(fs, JOBS, JOBS + "\\J-101\\Reports\\NDE"), "Blocked by file")
        self.assertEqual(fs.entries, before)

    def test_the_length_limits(self):
        fs = mock()
        name = "x" * (247 - len(JOBS) - 1)
        self.assertEqual(make_folder(fs, JOBS, JOBS + "\\" + name), "Created")
        self.assertEqual(make_folder(fs, JOBS, JOBS + "\\" + name + "y"), "Too long")


class CopyFile(unittest.TestCase):
    """E11: COPY_FILE's statuses. It never overwrites, deletes or makes folders."""

    def test_statuses(self):
        src = JOBS + "\\in\\r1.pdf"
        dst = JOBS + "\\J-101\\r1.pdf"
        cases = [
            (mock(), "", dst, "Bad name"),
            (mock(), JOBS + "\\in\\*.pdf", dst, "Bad name"),
            (mock(), src, JOBS + "\\" + "x" * 300, "Too long"),
            (mock(folders=["J-101"]), src, dst, "Missing"),
            (mock(("in\\r1.pdf", 120), folders=["in"]), src, dst, "No folder"),
            (mock(("in\\r1.pdf", 120), ("J-101\\r1.pdf", 120), folders=["in", "J-101"]), src, dst, "Already there"),
            (mock(("in\\r1.pdf", 120), ("J-101\\r1.pdf", 99), folders=["in", "J-101"]), src, dst, "Conflict"),
            (mock(("in\\r1.pdf", 120), folders=["in", "J-101", "J-101\\r1.pdf"]), src, dst, "Conflict"),
            (mock(("in\\r1.pdf", 120), folders=["in", "J-101"], locked=["in\\r1.pdf"]), src, dst, "Locked"),
            (mock(("in\\r1.pdf", 120), folders=["in", "J-101"]), src, dst, "Copied"),
        ]
        for fs, a, b, status in cases:
            with self.subTest(status=status):
                self.assertEqual(copy_file(fs, a, b), status)

    def test_a_copy_has_the_same_size_and_the_source_stays(self):
        fs = mock(("in\\r1.pdf", 120), folders=["in", "J-101"], locked=())
        self.assertEqual(copy_file(fs, JOBS + "\\in\\r1.pdf", JOBS + "\\J-101\\R1.pdf"), "Copied")
        self.assertEqual(fs.get(JOBS + "\\j-101\\r1.PDF")["size"], 120)   # paths ignore case, as on Windows
        self.assertTrue(fs.is_file(JOBS + "\\in\\r1.pdf"))
        self.assertEqual(copy_file(fs, JOBS + "\\in\\r1.pdf", JOBS + "\\J-101\\r1.pdf"), "Already there")

    def test_the_length_limit(self):
        fs = mock(("in\\r1.pdf", 120), folders=["in"])
        name = "x" * (258 - len(JOBS) - 1)
        self.assertEqual(copy_file(fs, JOBS + "\\in\\r1.pdf", JOBS + "\\" + name), "Copied")
        self.assertEqual(copy_file(fs, JOBS + "\\in\\r1.pdf", JOBS + "\\" + name + "y"), "Too long")


class MockFileSystem(unittest.TestCase):
    """F6 and F7: the folders the FIXTUREs build, and what the TESTs read from them."""

    def test_default_sizes_never_change(self):
        # The seed rule: 100 + below(900) from "<file>|<folder>|<path>|size". Pinned, as fixtures are.
        self.assertEqual(default_size("router_01.xlsx", "Dump", "r3.pdf"), 593)
        self.assertEqual(default_size("router_01.xlsx", "Dump", "r3.pdf"), 593)
        self.assertTrue(all(100 <= default_size("x_01.xlsx", "F", f"r{i}.pdf") <= 999 for i in range(200)))

    def test_folder_exists(self):
        spec = router(("OP       RETURN(n)", 'OP       IF FOLDER_EXISTS(PATH_JOIN($jobs, "J-102")) THEN LET(n, n + 10)\n'
                                            'OP       IF FOLDER_EXISTS(PATH_JOIN($jobs, "J-104")) THEN LET(n, n + 100)\n'
                                            "OP       RETURN(n)"))
        self.assertEqual([r.result for r in run_tests(spec)], [11, 11, -1])   # J-104 is a file

    def test_a_test_passes_a_name_that_is_neither_sheet_nor_folder(self):
        with self.assertRaises(VclError):
            run_tests(router((RUN, "call=RouteJobPacks(Tracker, Dump, Jobz) assert=")))

    def test_file_assertions_can_fail(self):
        spec = router(('FILE_EXISTS("Jobs\\J-101\\r1.pdf")', 'FILE_EXISTS("Jobs\\J-102\\r1.pdf")'))
        self.assertFalse(run_tests(spec)[0].passed)
        spec = router(('FILE_COUNT("Jobs") = 3', 'FILE_COUNT("Jobs") = 4'))
        self.assertFalse(run_tests(spec)[1].passed)

    def test_the_validator_checks_file_paths(self):
        self.assertIn('assertion: FILE_COUNT needs a "quoted" path that starts with a FIXTURE folder of '
                      'router_01.xlsx, like "Jobs\\J-101\\r1.pdf", got \'"Jobz"\'',
                      errors(router(('FILE_COUNT("Jobs") = 3', 'FILE_COUNT("Jobz") = 3'))))
        self.assertIn('assertion: FILE_EXISTS: "Jobs\\..\\Dump\\r1.pdf" is not a safe path (every part must be '
                      'a SAFE_NAME)', errors(router(('"Jobs\\J-101\\r1.pdf"', '"Jobs\\..\\Dump\\r1.pdf"'))))
        older = TEST_HEAD_02 + 'TEST     fixture=jobs_01.xlsx call=P(Jobs) assert=FILE_COUNT("Jobs") = 0\n'
        self.assertIn("assertion: FILE_COUNT is VCL 0.3: start the spec with 'VCL 0.3' to use it", errors(older))


TEST_HEAD_02 = """VCL 0.2
TASK
PROC     P
KIND     Sub
MODULE   modP
ARGS     ws As Worksheet
OP       LOG("x")
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID rows=2 data=A:values("J-1"x2)
"""


class UnchangedFiles(unittest.TestCase):
    """E10: nothing on disk changes except new folders and files below a declared root."""

    def test_a_copy_outside_the_declared_root_fails_the_test(self):
        # The validator rejects this spec; the interpreter must still catch what it does.
        dest = "PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile)))"
        results = run_tests(router((dest, 'PATH_JOIN($dump, "copy.pdf")')))
        self.assertIn("UNCHANGED: file Dump\\copy.pdf was added", results[0].failed)

    def test_the_problems(self):
        before = mock(("a.pdf", 5), ("b.pdf", 6), folders=["J-101"])
        after = before.snapshot()
        after.add_file(JOBS + "\\J-101\\new.pdf", 7)
        after.add_file(JOBS + "\\b.pdf", 9)
        del after.entries[(JOBS + "\\a.pdf").casefold()]
        after.add_file(SANDBOX + "\\Other\\x.pdf", 1)
        self.assertEqual(file_unchanged_problems([JOBS], before, after, False),
                         ["UNCHANGED: file Jobs\\a.pdf was removed",
                          "UNCHANGED: file Jobs\\b.pdf was file, 6 bytes, now file, 9 bytes",
                          "UNCHANGED: file Other\\x.pdf was added"])
        self.assertIn("UNCHANGED: file Jobs\\J-101\\new.pdf was added",
                      file_unchanged_problems([JOBS], before, after, True))   # every run stopped at PRE


class Router(unittest.TestCase):
    """The sample router end to end: every status, re-runs and a missing root."""

    def test_it_passes_its_tests(self):
        results = run_tests(ROUTER)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed) for r in results])
        self.assertEqual([r.result for r in results], [1, 1, -1])
        statuses = [v for (row, col), v in results[1].state["Tracker"][0] if col == 3 and row > 1]
        self.assertEqual(statuses, ["Already there", "Conflict", "Locked", "Bad name", "Blocked by file", "Missing"])
        self.assertIn(("Jobs\\J-101\\r1.pdf", "file", 120, False), results[0].files)

    def test_its_tests_catch_mistakes(self):
        # Both survivors change nothing: a Long starts at 0 (D14), and with AND an unsafe name still
        # reaches the file ops, which give "Bad name" for the empty path PATH_JOIN returns (S19, E11).
        killed, total, survivors = mutate_report(ROUTER, snapshot=True)
        self.assertEqual((killed, total), (29, 31))
        self.assertEqual(survivors, [(20, "step deleted"), (23, "'OR' -> 'AND'")])


class Task12(unittest.TestCase):
    """Router Phase 4: specs/task_12.vcl RouteJobPacks, the MVP spec (D32)."""

    TEXT = (ROOT / "specs" / "task_12.vcl").read_text(encoding="utf-8")

    def test_every_status_on_its_row(self):
        results = run_tests(self.TEXT)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed) for r in results])
        self.assertEqual([r.result for r in results], [3, 3, -1, -1])
        status = {row: v for (row, col), v in results[0].state["Tracker"][0] if col == 3 and row > 1}
        self.assertEqual(status, {2: "Copied", 3: "Conflict", 4: "Locked", 5: "Bad name", 6: "Blocked by file",
                                  7: "old", 8: "Missing", 9: "Copied", 10: "Bad name", 11: "Already there"})
        folders = {path for path, kind, _, _ in results[0].files if kind == "folder"}
        self.assertNotIn("Jobs\\J-106", folders)   # a bad Report File: nothing is made for the row
        self.assertTrue({"Jobs\\J-103", "Jobs\\J-105"} <= folders)   # made, though the copy failed

    def test_its_tests_catch_every_mutant(self):
        self.assertEqual(mutate_report(self.TEXT, snapshot=True), (29, 29, []))

    def test_it_compiles_with_a_sandbox_builder(self):
        vba = compile_text(self.TEXT, "task_12.vcl")
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn("ForgeSandbox_jobpacks_01", fixtures_module([("task_12.vcl", fixture_trees(self.TEXT))]))


class Interpreter(unittest.TestCase):
    def test_the_example_passes_its_tests(self):
        results = run_tests(NAMES)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed, r.result) for r in results])
        self.assertEqual(results[0].result, 3)



FORGELIB = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")


def vba_body(name):
    """The text of one ForgeLib procedure, from its first line to its End line."""
    start = re.search(r"^(Public|Private) (Function|Sub) " + name + r"\(", FORGELIB, re.M).start()
    return FORGELIB[start:FORGELIB.index("\nEnd ", start)]


class Compiler(unittest.TestCase):
    """Router Phase 3a: the file ops compile to calls of ForgeLib's file helpers (D29)."""

    def test_the_value_functions(self):
        vba = compile_text(NAMES)
        self.assertIn("ForgeSetValue ws.Cells(r, cName), ForgeSafeName(ws.Cells(r, cJob).Value)", vba)
        self.assertIn('ForgePathJoin(root, Array(ForgeSafeName(ws.Cells(r, cJob).Value), "Reports"))', vba)

    def test_the_file_ops_get_their_roots(self):
        vba = compile_text(ROUTER)
        self.assertIn('If Not ForgeFolderExists(dump) Then ForgeLog "RouteJobPacks", 0, '
                      '"PRE failed: FOLDER_EXISTS($dump)"', vba)
        job = "ForgeSafeName(tracker.Cells(r, cJob).Value)"
        report = "ForgeSafeName(tracker.Cells(r, cFile).Value)"
        self.assertIn(f"status = ForgeMakeFolder(jobs, ForgePathJoin(jobs, Array({job})))", vba)
        self.assertIn(f"status = ForgeCopyFile(dump, ForgePathJoin(dump, Array({report})), "
                      f"jobs, ForgePathJoin(jobs, Array({job}, {report})))", vba)

    def test_the_module_is_lint_clean(self):
        self.assertEqual(lint(assemble([compile_text(ROUTER), compile_text(NAMES)])), [])

    def test_a_bare_root_is_refused(self):
        with self.assertRaises(CompileError):
            compile_text(spec("MAKE_FOLDER($root, into=t)"))


class ForgeLibFiles(unittest.TestCase):
    """ForgeLib can't run here (Excel only, router Phase 5), so these pin it to the reference."""

    def const(self, name):
        return re.search(r"^Private Const " + name + r" As \w+ = (.+)$", FORGELIB, re.M).group(1)

    def test_the_limits_are_the_reference_limits(self):
        self.assertEqual(int(self.const("FORGE_MAX_NAME")), semantics.MAX_NAME)
        self.assertEqual(int(self.const("FORGE_MAX_FILE_PATH")), vcl_interp.MAX_FILE_PATH)
        self.assertEqual(int(self.const("FORGE_MAX_FOLDER_PATH")), vcl_interp.MAX_FOLDER_PATH)
        self.assertEqual(self.const("FORGE_NAME_FORBIDDEN")[1:-1].replace('""', '"'), semantics.NAME_FORBIDDEN)

    def test_the_reserved_names(self):
        self.assertEqual(semantics.RESERVED_NAMES, {"CON", "PRN", "AUX", "NUL"} | {f"{d}{i}" for d in ("COM", "LPT")
                                                                                     for i in range(1, 10)})
        body = vba_body("ForgeSafeName")
        self.assertIn('If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Then Exit Function', body)
        self.assertIn('(Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT")', body)
        self.assertIn("AscW(Right$(stem, 1)) >= 49 And AscW(Right$(stem, 1)) <= 57", body)   # 1 to 9

    def test_every_status_and_no_other(self):
        for fn, helper, excel_only in ((vcl_interp.make_folder, "ForgeMakeFolder", {"Failed"}),
                                       (vcl_interp.copy_file, "ForgeCopyFile", {"Check copy", "Failed"})):
            reference = {status for line in inspect.getsource(fn).splitlines() if "return " in line
                         for status in re.findall(r'"([A-Z][^"]*)"', line)}
            self.assertEqual(set(re.findall(helper + r' = "([^"]+)"', vba_body(helper))), reference | excel_only)

    def test_one_attempt_never_overwrite_never_delete(self):
        copy = vba_body("ForgeCopyFile")
        self.assertEqual(re.findall(r"\.CopyFile [^\n]*", copy), [".CopyFile src, dst, False"])
        for helper in ("ForgeFolderExists", "ForgeMakeFolder", "ForgeCopyFile"):
            body = vba_body(helper)
            self.assertIn("On Error GoTo FailedOp", body)
            self.assertNotRegex(body, r"(?i)\b(Delete|Kill|Move|Resume|Do )")


class LintFileSafety(unittest.TestCase):
    """The linter keeps every file write inside ForgeLib's helpers (D29)."""

    def test_catches_planted_faults(self):
        m = assemble([compile_text(ROUTER)])
        self.assertEqual(lint(m), [])
        faults = {"Kill": m.replace("    n = 0\n", '    n = 0\n          Kill "x"\n', 1),
                  "DeleteFile": m.replace("fso.CopyFile src, dst, False", "fso.DeleteFile dst", 1),
                  "MkDir": m.replace("    n = 0\n", '    n = 0\n          MkDir "C:\\x"\n', 1),
                  "Open": m.replace("    n = 0\n", '    n = 0\n          Open "x" For Output As #1\n', 1),
                  "Shell": m.replace('CreateObject("Scripting.Dictionary")', 'CreateObject("WScript.Shell")', 1),
                  "SaveAs": m.replace("    n = 0\n", '    n = 0\n          ThisWorkbook.SaveAs "x"\n', 1),
                  "registry": m.replace("    n = 0\n", '    n = 0\n          SaveSetting "a", "b", "c", "d"\n', 1),
                  "FSO made elsewhere": m.replace('CreateObject("Scripting.Dictionary")',
                                                  'CreateObject("Scripting.FileSystemObject")', 1),
                  "CopyFile outside": m.replace("    n = 0\n", '    n = 0\n          ForgeFso().CopyFile "a", "b"\n', 1),
                  "CreateFolder outside": m.replace("fso.CopyFile src, dst, False", "fso.CreateFolder dst", 1),
                  "late declaration": m + "Private gLate As Long\n"}
        for what, bad in faults.items():
            with self.subTest(what):
                self.assertNotEqual(bad, m)
                self.assertTrue(lint(bad))


if __name__ == "__main__":
    unittest.main()
