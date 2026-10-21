"""The prompt's examples (vcl/examples/) as VCL 0.6 files, and as the VCL 0.5 specs they were before the GOAL
line (D63), which the tests of the older versions still use."""
from pathlib import Path

EXAMPLES = Path(__file__).resolve().parent.parent / "vcl" / "examples"


def example(name):
    """The example file as it is: VCL 0.6, with its GOAL lines."""
    return (EXAMPLES / name).read_text(encoding="utf-8")


def without_goals(text, version="VCL 0.5"):
    """The spec with its GOAL lines taken out and the version line `version` (every section of a BUILD)."""
    lines = [l for l in text.split("\n") if not l.startswith("GOAL ")]
    return "\n".join(version if l == "VCL 0.6" else l for l in lines)


def old_example(name):
    """The example as the VCL 0.5 spec it was: no GOAL lines."""
    return without_goals(example(name))
