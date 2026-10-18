"""VCL 0.3 (docs/vep-0.3-file-ops.md). Router Phase 1a: the version line, SAFE_NAME, PATH_JOIN
and FOLDER_EXISTS, and the path rules of the type checker (T4). Phase 1b: MAKE_FOLDER and
COPY_FILE (E9), MODIFIES $root files (E10) and folder FIXTUREs (F6)."""
import glob
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
import validator  # noqa: E402
from vcl_fixtures import fixtures_module  # noqa: E402
from vcl_interp import VclError, modifies_of  # noqa: E402
from vcl_codegen import CompileError, compile_text  # noqa: E402
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
OP           COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))), PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Copied" THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  router_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=6 data=A:values("J-101"x2,"J-102"x1,"J-103"x1,"..\x"x1,""x1);B:values("r1.pdf"x1,"r2.pdf"x2,"r3.pdf"x1,"r4.pdf"x1,"r5.pdf"x1)
FIXTURE  router_01.xlsx folder=Dump files=r1.pdf:120,r2.pdf:80,r3.pdf locked=r3.pdf
FIXTURE  router_01.xlsx folder=Jobs files=J-102\r2.pdf:99
TEST     fixture=router_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) runs=2 assert=ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Bad name") = 1 AND FILE_EXISTS("Jobs\J-101\r1.pdf")
TEST     fixture=router_01.xlsx call=RouteJobPacks(Tracker, Dump, "S:\Missing") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID])
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
        paths = sorted(glob.glob(os.path.join(ROOT, "specs", "*.vcl")))
        self.assertGreaterEqual(len(paths), 11)
        for path in paths:
            text = Path(path).read_text(encoding="utf-8")
            with self.subTest(spec=os.path.basename(path)):
                self.assertTrue(text.startswith("VCL 0.2\n"))
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
RUN = "call=RouteJobPacks(Tracker, Dump, Jobs)"


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
                         ("router_01.xlsx", "Dump", [("r1.pdf", 120), ("r2.pdf", 80), ("r3.pdf", None)], ["r3.pdf"]))
        self.assertEqual((jobs.folder, jobs.files, jobs.locked), ("Jobs", [("J-102\\r2.pdf", 99)], []))
        self.assertEqual(problems(router(("folder=Jobs files=J-102\\r2.pdf:99", "folder=Jobs"))), [])

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
            (("files=r1.pdf:120,", "files=r1.pdf:big,"), "FIXTURE files: expected <path> or <path>:<bytes>, "
                                                         "got 'r1.pdf:big'"),
        ]
        for edit, message in cases:
            with self.subTest(edit=edit):
                self.assertIn(message, errors(router(edit)))

    def test_tests_pass_folders_bare(self):
        self.assertIn("TEST passes Jobz for $jobs, but router_01.xlsx has no folder Jobz",
                      errors(router((RUN, RUN.replace("Jobs)", "Jobz)")))))
        found = errors(router((RUN, "call=RouteJobPacks(Dump, Tracker, Jobs)")))
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

    def test_folders_wait_for_the_mock_file_system(self):
        with self.assertRaises(VclError):
            run_tests(ROUTER)
        from vcl_parse import parse as grammar_parse
        trees = [k["tree"] for k in grammar_parse(ROUTER)[0]["keys"] if k["key"] == "FIXTURE"]
        with self.assertRaises(ValueError):
            fixtures_module([("task_x", trees)])


class Interpreter(unittest.TestCase):
    def test_the_example_passes_its_tests(self):
        results = run_tests(NAMES)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed, r.result) for r in results])
        self.assertEqual(results[0].result, 3)

    def test_the_compiler_waits_for_phase_3(self):
        with self.assertRaises(CompileError):
            compile_text(NAMES)


if __name__ == "__main__":
    unittest.main()
