"""Where the repo's data files are (item 5 Phase 3, D79).

Run from a checkout, ROOT is the repo: vcl/ (the grammar, the Orchestrator prompt, semantics.py) and vba/
(ForgeLib.bas) sit beside compiler/. Inside KRONOS.exe (PyInstaller, one file), the program is unpacked into a
temporary folder named by sys._MEIPASS, and the build puts vcl/ and vba/ there, so ROOT points at it. Code that
needs one of those files asks here, instead of counting parent folders from its own file.
"""
import sys
from pathlib import Path

FROZEN = getattr(sys, "_MEIPASS", None) is not None
ROOT = Path(sys._MEIPASS) if FROZEN else Path(__file__).resolve().parent.parent   # noqa: SLF001
