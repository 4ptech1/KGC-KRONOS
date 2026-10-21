# VEP 0.5: BUILD specs and CALL (the VEP 0.4 group D)

Status: **design, 21 Oct 2026**. Boss said "continue building the next part according to your
recommendations" after item 2; the recommendation was CALL and BUILD specs before the AI plan's
Stage A, because the Orchestrator's rule 9 already asks for a BUILD when a task needs several
procedures or more than about 60 lines, and today no BUILD can be proven. Choices that needed
Boss's decision were made the conservative way and recorded in `DECISIONS.md` D49; Boss can
overrule any of them.

## The problem

A BUILD is several TASK procedures that make one tool: `modMain.RunHeadless` runs the whole tool
and is what TESTs call, `modMain.Run` is what the user starts, and CALL joins the procedures. The
validator has checked BUILDs since revision 3 (keys, SIGs, one TASK per SIG, CALL targets and
argument counts, the TESTs), but nothing runs or compiles them: `vclc test` and `vclc build` stop
with "only TASK specs". So a valid BUILD is not a provable BUILD, which is the gap item 2 closed
for the single ops. It matters for the corpus too: the larger register tasks (task_04 has 48 OP
lines) need several procedures.

What BUILD and CALL lack today:

| Part | Today |
| --- | --- |
| CALL | compiles to `Call Proc(...)`; the interpreter can't run it; a Function's result is lost |
| BUILD TESTs | validated, never run |
| MODIFIES in a BUILD | not required in its TASKs, so no UNCHANGED check is possible |
| BUILD output | none: `vclc build` refuses BUILDs |
| FORM, form events | in the language, but no op can show a form or read a control |

## Principles (unchanged)

- The interpreter is the reference meaning; every meaning goes through ForgeLib or plain
  generated VBA that the compiler proves; a TEST proves it on synthetic data.
- MODIFIES covers every change, and UNCHANGED checks it after every TEST.
- One language version per spec; the new parts need `VCL 0.5`; spec.md and the prompt move to
  0.5 in one step after the last phase, so Boss updates the Orchestrator once.

## Design

**A BUILD is one tool in one workbook.** The compiler writes one module per BUILD,
`dist/builds/<NAME>.bas`: ForgeLib plus every procedure of the BUILD, so Boss pastes one module,
as today. MODULE lines stay as the spec's own grouping; procedure names must be unique within a
BUILD (VBA can't call two same-named procedures from different modules without the module name,
and the compiler calls them by name). Two BUILDs never share a module, so every BUILD can have its
own Run and RunHeadless.

**CALL runs another procedure of the same BUILD.**
- `CALL(<module>.<Proc>, <args>)` runs that TASK; `CALL(..., into=<var>)` (VCL 0.5) keeps a
  Function's result, so RunHeadless can add up what its steps did and return it (rule 17).
- Each procedure keeps its own ERR standard and PRE, as the compiled VBA does: a failed PRE in the
  callee logs once and gives -1, a run-time error in it logs once and gives -2, and the caller goes
  on. A caller checks `into=` (IF n < 0 ...) when a step must not be skipped silently.
- A Worksheet, Workbook or Range parameter is passed one of the caller's own `$arguments` of that
  type, written in place, so MODIFIES can follow the sheet from the TEST down to the step that
  changes it. Values (Long, Double, String, Boolean, Date) are any expression of their type.
- No recursion: the CALL graph must have no cycle (every loop has a bound, CLAUDE.md).
- CALL is only for procedures of the same BUILD; in a VCL 0.5 TASK spec it is an error.
- FLOW lists the calls in order for the reader; every `a -> b` in FLOW must be a CALL from a to
  b, and every CALL edge must be in a FLOW line.

**MODIFIES follows the arguments.** Every TASK in a 0.5 BUILD declares MODIFIES for its own
writes, as a TASK spec does. A CALL that passes `$x` to a parameter the callee modifies needs a
MODIFIES line on `$x` in the caller that allows the same change (the whole sheet allows any
change; `["H"]` lines must name the same headers; `fills`, `files` and `sheets` must match). So
RunHeadless's MODIFIES lines cover the whole tool, and every BUILD TEST gets the same UNCHANGED
check a TASK TEST gets.

**modMain.Run starts the tool.** With no form, Run is a Sub with no ARGS that calls RunHeadless
with the sheets of the workbook the tool lives in: `CALL(modMain.RunHeadless, SHEET(THIS_BOOK,
"Orders"), ..., into=n)` and `MSG(...)`. It may only CALL RunHeadless, LET, IF and MSG, so it has
nothing left to test; TESTs call RunHeadless, as today. Run is the one caller that passes sheets
by their real names, with `SHEET(THIS_BOOK, "<name>")` written in place, so it needs no MODIFIES
lines; the names are the real workbook's, as its ASSUME lines describe it.

**Deferred (D49):**
- **FORM and form events:** a UserForm can't be added by pasting text (its controls live in a
  binary .frx file, and building one from code needs the Trust Center change CLAUDE.md rules out),
  and no op shows a form or reads a control. FORM lines stay valid, the compiler refuses them with
  a clear message, and the reference text stops teaching them until a design for after March 2027.
- **TARGET xlam:** an add-in's THIS_BOOK is the add-in, not the data, and installing one may need
  settings Boss's PC doesn't allow. `xlsm` only for now; `xlam` is refused with a clear message.
- **BUILDs before 0.5** stay check-only: their TASKs don't declare MODIFIES, so no TEST could check
  that nothing else changed.

## Also in 0.5: the gaps the item 2 review found (D47)

Small language gaps the Orchestrator dry runs hit; each gets a decision in its phase:
- No row number of a found cell, so FIND can't lead to another column of the same row (D39).
- No count of the rows DEDUPE removed (rule 17 wants a count).
- Rule 14 has no form for a header the macro itself adds to an existing sheet.
- An assertion can't name a sheet whose name has spaces.

## Phases (one or two PRs each, in this order)

1. **The language surface and the validator:** the `VCL 0.5` line, CALL `into=`, unique procedure
   names, the CALL graph without cycles, CALL only in a BUILD, FLOW as the CALL edges, MODIFIES
   in every TASK of a BUILD and its propagation, the rules for Run, FORM and xlam.
2. **The interpreter:** BUILD TESTs run RunHeadless with a procedure table; CALL with per-procedure
   PRE and ERR; UNCHANGED from RunHeadless's MODIFIES; `vclc test` on BUILDs.
3. **The compiler:** `dist/builds/<NAME>.bas`, CALL `into=` as a Function call, the BUILD's
   fixture builders and expected results, lint, and the mutation scorer over every TASK of a BUILD.
4. **The review gaps** above, as the evidence decides.
5. **A proven BUILD and the switch:** a sample BUILD in the tests that catches every mutant, a new
   Example 4 in the prompt, rule 9 and a CALL rule in the hard rules, spec.md and the prompt move
   to 0.5, and Boss updates the Orchestrator once.

## Verification (every phase)

The four standard checks and `vclc mutate --snapshot` (the score stays 95.4% or better), plus for
each new rule a validator test, for each new meaning an interpreter test, and for the compiler a
lint-clean module. Every 0.4 spec stays valid under 0.5 with the same results (as E8 and E12).
