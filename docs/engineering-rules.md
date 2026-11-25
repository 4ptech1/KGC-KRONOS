# KGC's engineering rules: the sheet for Boss to confirm (D133, E2)

These are the only engineering rules KRONOS may use (`docs/engineering-knowledge-plan.md`). Each comes from a formula in
KGC's own samples (`samples/kgc/`, published by Boss, D106), with the worked example the sample itself gives: its
inputs and the result Excel saved. Nothing here comes from an AI chat or from Claude's own knowledge (D132).

**What Boss does:** read the table and reply **"confirmed"**, or name a row to correct. Rows marked **"needs your word"**
are written in the samples only as words, not as a formula: KRONOS uses them only if Boss confirms them too.

After confirmation, each rule becomes a proven recipe (E3, E4) whose TESTs use invented look-alike values (tests never
use the samples' own numbers, D105). Every worked example below was recomputed from its formula when this sheet was
written, and each matches the result the sample saved. Every result KRONOS gives with these rules is
labelled preliminary, to be checked by a qualified engineer.

## Welding (WPS and PQR)

| # | Rule | Formula | Units | Worked example from the sample | Source |
| --- | --- | --- | --- | --- | --- |
| W1 | Heat input, lowest of a run | Q min = k x U min x I min x 60 / (v max x 1000) | kJ/mm (U in V, I in A, v in mm/min) | k 0.6, U 11, I 90, v 110: **0.324** | WPS P105 R33 |
| W2 | Heat input, highest of a run | Q max = k x U max x I max x 60 / (v min x 1000) | kJ/mm | k 0.6, U 15, I 150, v 70: **1.157** | WPS P105 S33 |
| W3 | Heat input of a PQR run | Q = k x U x I x 60 / (v x 1000), on the PQR's recorded values | kJ/mm | k 0.6, U 15, I 140, v 93: **0.813** | WPS P105 AB33 |
| W4 | Run within the WPS heat-input limit | "OK" when Q max <= the limit, else "OVER LIMIT"; with a minimum too: "OK" when Q max <= max and Q min >= min, else "OUT OF RANGE" | | P105: 1.157 <= 1.85: **OK**; P104: 2.3625 <= 2.37 and 1.0047 >= 0.99: **OK** | WPS P105 T33, P104 T34 |
| W5 | PQR run inside the WPS range | "Yes" when PQR I min >= WPS I min, PQR I max <= WPS I max, PQR U min >= WPS U min, PQR U max <= WPS U max and WPS v min <= PQR v <= WPS v max, else "No" | | I 100-140 in 90-150, U 12-15 in 11-15, v 93 in 70-110: **Yes** | WPS P105 AC33 |
| W6 | k-factor of the process | k = 0.6 for process 141 (GTAW) | | **0.6** | WPS P105 N26 (a typed value with its label) |
| W7 | Qualified thickness range (multi-run butt welds) | from 0.5 x t to 2 x t | mm | t 14: **7 to 28**; t 22: **11 to 44** | WPS P105 F28, P104 F29 (**needs your word**: written as text) |
| W8 | WPS heat-input limits from the PQR | max = 1.25 x the PQR's highest when impact tested; min = 0.75 x the PQR's lowest when hardness tested; otherwise the PQR's own highest | kJ/mm | 1.25 x 1.895 = 2.369, shown **2.37**; 0.75 x 1.325 = 0.994, shown **0.99** (rounded to 2 places; a minimum rounded down widens the range a little: say which way you want it) | WPS P104 N27, N28 (**needs your word**: typed values with the rule in their labels) |

## Vessel data

| # | Rule | Formula | Units | Worked example from the sample | Source |
| --- | --- | --- | --- | --- | --- |
| V1 | Shell outside diameter | OD = ID + 2 x thickness | mm | ID 400, thickness 10: **420** | register Q9 |
| V2 | Pressure from bar to kg/cm2 | ROUND(bar x 1.01972, 2) | kg/cm2 (g) | 18 barg: **18.35** | register W9 |

## Estimation

| # | Rule | Formula | Units | Worked example from the sample | Source |
| --- | --- | --- | --- | --- | --- |
| E1 | Raw weight of a forged ring | PI / 4 x ((OD + 2a)^2 - (ID - 2a)^2) x (H + 2a) x density, a = machining allowance per surface | kg (mm, kg/mm3) | OD 1417, ID 1190, H 185, a 5, density 7.85E-06: **774.16** | estimation BOM-RM M5 |
| E2 | Raw weight of a plate | (L + 2m) x (W + 2m) x T x density x (1 + scrap), m = cutting margin per edge | kg | L 3801, W 569, T 20, m 25, scrap 0.1: **411.68** | estimation BOM-RM M7 |
| E3 | Weight of a pipe | length / 1000 x kg per metre | kg | 1500 mm at 133 kg/m: **199.5** | estimation BOM-RM M12 |
| E4 | Welding consumable to buy | weld metal / deposition efficiency | kg | 35 kg at 0.6: **58.33** | estimation Conversion E7 |
| E5 | Overheads and contingency | factory cost x the rate | INR | 4,905,904.65 x 0.12: **588,708.56**; x 0.03: **147,177.14** | estimation Summary B9, B10 |
| E6 | Selling price per unit | ROUNDUP(total cost x (1 + margin) / number of units, -3): up to the next 1,000 | INR | 5,873,828.74 x 1.15 / 2: **3,378,000** | estimation Summary B21 |
| E7 | Price in US dollars | INR price / exchange rate | USD | 3,378,000 / 88: **38,386.36** | estimation Summary B23 |
| E8 | Price per kg finished | total price / total finished weight | INR/kg | 11,608,000 / (3,836.35 + 1,733.359): **2,084.13** | estimation Summary B31 |

## Not on this sheet (and why)

- The ASME VIII-1 design checks the estimation sheets show (shell and head thickness, MAWP, hydrotest pressure,
  external pressure, flange ratings): the samples give their results, not their formulas. They come in E6, only when
  Boss gives each formula, the code edition and a worked example.
- Values from code books (allowable stresses, flange ratings): only from Boss's own workbook (E5), never here.
