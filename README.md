# VBA Forge

VBA Forge is an offline, VBA-only code generator for Excel automation in pressure-vessel
fabrication registers (BOM, weld log, NDE, NCR, calibration and similar). Claude writes each task
as a spec in VCL (VBA Command Language), and VBA Forge turns the spec into VBA and proves the
macro works with tests run in Excel. The verified spec-to-VBA pairs become the dataset for a
fine-tuned local model that takes over generation from March 2027.

## Current phase

**Phase 0, Foundation (Oct 14 to 31, 2026).** Everything runs in the cloud (Claude Code on the
web, GitHub Codespaces, GitHub Actions) until the new PC arrives in March 2027. Excel testing is
manual on the work PC.

See `CLAUDE.md` for the working rules and `docs/VBA_Forge_Project_Roadmap.docx` for the full plan.

## Check the specs

From the repo root, with Python 3.11:

```
python compiler/vcl_check.py specs/*.vcl
```

It prints `OK` or `FAIL` per spec, one line per problem, and ends with
`N files checked, N problems`. It exits 0 when everything passes and 1 otherwise. GitHub Actions
runs the same check on every push and pull request.

## Dataset count

Dataset count: 0 verified pairs

## Progress log

- **2026-10-14:** Repo set up. Added the 10 task specs (VCL 0.1 rev 2), `vcl/spec.md`, the
  Orchestrator prompt, the starter checker and the roadmap; all 10 specs pass the checker. Added
  the project folders, `.gitignore` and the spec-check CI workflow.
