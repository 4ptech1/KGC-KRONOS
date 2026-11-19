# The AI writes the VBA, with its context in VCL (D90)

Status: **decided by Boss on 17 Nov 2026.** This is the design; each phase below gets its own PR. No training starts
before the new PC arrives in March 2027 (CLAUDE.md), so Phases 1 and 2 come first and need no AI at all.

## Boss's decision

Asked whether KRONOS should keep the compiler as the only writer of VBA (D31), Boss answered: "Let proof work around
my idea. I want kronos to use ai written in vcl to do vba tasks." His reasons: VCL is far shorter and more exact than
English, so an AI whose context and memory are VCL is smaller, faster and easier to train, and his own PC has room for
it (300 GB).

So **our own model (D25) writes the VBA for a task, and its context and memory are written in VCL.** This changes D31
("the model writes VCL, never VBA"): the model may now write VBA too, from a VCL recipe. Everything else in D25 stays:
the model learns from in-house data only. (Since D102, 19 Nov 2026, it may start from a small open-weight model tuned
on that data instead of from zero; see `docs/ai-plan.md`.)

## How the proof works around it

The rule that nothing unproven reaches a workbook unlabelled stays. The proof moves to where the AI works:

1. **The recipe comes first, and is proven as today.** Every job starts from a VCL recipe (written by a person, an AI
   chat, or later our model), which KRONOS checks and tests on practice data (`kronos_engine.prove`). The recipe is
   the AI's context and its contract: it says exactly what the VBA must do, and its TESTs say how to tell.
2. **The AI writes the VBA from the recipe.** Its context: the recipe. Its memory: the library of proven recipes, each
   with the VBA it became (`data/verified/`, today 84 pairs in 1.6 MB, so even 10,000 recipes fit easily on the PC).
   Similar recipes are found in that library, in-house only (D25), and shown to the model as examples.
3. **KRONOS proves the AI's VBA against the recipe's own TESTs, without Excel.** A new VBA simulator
   (`compiler/vba_sim.py`, Phase 1) runs the kind of VBA KRONOS allows, with ForgeLib, on the same practice workbooks the
   TESTs use. The AI's macro must end with exactly the workbook, RESULT and ERROR_COUNT the recipe's TESTs expect (the
   interpreter's answers, `vclc build`'s `expected/`), and pass `vba_lint` (no Kill, Shell, internet, Select, ...).
4. **The labels Boss sees.**
   - **PROVEN, written by AI:** the AI's VBA passed every TEST in the simulator and the lint.
   - **Not proven:** KRONOS says what failed and gives the compiler's VBA for the same recipe instead, which is always
     proven. So Boss always gets a working macro; the AI's version is used only when it is proven.
5. **The last check, in Excel (optional, as now):** the practice run (D81) with the AI's module, on a new empty workbook,
   compared with the TESTs.

## Phases

| Phase | What | When |
| --- | --- | --- |
| 0 | This design, D90, CLAUDE.md and `docs/ai-plan.md` changed to match | now |
| 1 | **`compiler/vba_sim.py`, the VBA simulator.** Part 1, the language core, is done (D98): the ForgeLib self-test passes in it, and planted bugs are caught. Part 2 is done (D99): every compiled TASK without folders ends where the interpreter does on every TEST, and two compiler bugs it found are fixed. Part 3 is done (D100): files and BUILDs too, 165 specs, examples and pairs, `python compiler/vclc.py sim` in CI. **Phase 1 is complete.** It runs the subset of VBA that KRONOS writes and allows, with ForgeLib, on the TESTs' practice workbooks. First proof: the compiler's own VBA for every spec, example and corpus pair ends exactly where the interpreter says (this already makes today's macros better proven, with no AI) | before March 2027 |
| 2 | **The memory.** The training pairs grow from request → VCL to request → VCL → VBA (`vclc export` already compiles the VBA); a small index finds similar proven recipes for a new one | before March 2027 |
| 3 | **The model** (new PC): trained on the pairs (from zero, or a small open model tuned on them, D102); KRONOS gets "the AI writes the VBA": prove the recipe, ask the model for the VBA, prove it in the simulator, else fall back to the compiler | from March 2027 |
| 4 | KRONOS.exe and KRONOS Offline show the labels; the practice run checks the AI's module in Excel | with Phase 3 |

## Where it lives

From March 2027 the memory, the simulator's results, the training data, the training runs and the released models live on the new PC's 300 GB, in `memory\`, `sim\`, `train\` and `models\` (`docs/kronos-storage-plan.md`, D91). Models trained on real data never leave that PC.

## What doesn't change

- D25 and D102: everything the model learns from is ours and proven; no outside data or outside retrieval. Since D102 its
  starting weights may be a small open model's (Apache 2.0 or MIT, at most about 7 billion parameters).
- Synthetic data only in the repo; no training before March 2027.
- The compiler and ForgeLib stay: they are the reference for what VBA a recipe means, and the fallback whenever the AI's
  VBA is not proven.
- Running anything in Excel stays optional and never blocks work.
