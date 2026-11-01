# PyInstaller spec for KRONOS.exe (item 5 Phase 3, D79). From the repo root:
#     python -m PyInstaller --noconfirm packaging/kronos.spec        ->  dist/KRONOS (KRONOS.exe on Windows)
# One file, no console window. The program's own modules are found through `pathex` (the code adds vcl/ and
# compiler/ to sys.path at run time, which the build cannot see), and the data files the program reads go into the
# unpacking folder under the same names they have in the repo, so compiler/resources.py finds them.
from pathlib import Path

root = Path(SPECPATH).resolve().parent   # noqa: F821  (SPECPATH is set by PyInstaller)

a = Analysis(   # noqa: F821
    [str(root / "app" / "kronos_app.py")],
    pathex=[str(root / "app"), str(root / "compiler"), str(root / "vcl")],
    datas=[(str(root / "vcl" / "grammar.lark"), "vcl"),
           (str(root / "vcl" / "orchestrator_prompt.txt"), "vcl"),
           (str(root / "vcl" / "examples" / "example1_rev3.vcl"), "vcl/examples"),   # the --self-test's recipe
           (str(root / "vba" / "ForgeLib.bas"), "vba")],
    hiddenimports=[],
    excludes=["tests", "unittest", "pydoc", "doctest"],
)
pyz = PYZ(a.pure)   # noqa: F821
exe = EXE(   # noqa: F821
    pyz, a.scripts, a.binaries, a.datas, [],
    name="KRONOS",
    console=False,   # a window, not a console; --self-test writes to its --log file instead
    upx=False,       # no packer: it makes antivirus tools more suspicious of the file
    disable_windowed_traceback=True,   # app/kronos_app.py reports errors itself (crash.log and a box)
)
