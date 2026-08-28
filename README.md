# Formulation Reports

Standalone, one-page formulation records generated from the lab's master workbook,
`FORMULATIONS_DATABASE_V2.xlsm`. Each record is a self-contained recipe + guarantee
sheet a lab or a partner can read on its own, without the master file.

## Output

**`IS_all_formulation_reports.xlsx`** — one workbook, 199 sheets:

- `Index` — all 197 formulations with date, site, process, dry/wet and a data-completeness
  tag, hyperlinked to each sheet.
- `Fo1` … `Fo185`, plus the IFDC sub-variants `Fo81A`-`Fo81D`, `Fo82A`-`Fo82D`, `Fo83A`-`Fo83D`
  (197 report sheets total).
- `Needs Input` — every formulation with an unresolved gap, and exactly what is missing.

Each report has five blocks: header/identity, raw materials (recipe table), final-product
active content (guarantees), protocol + protocol check, and a footer stating the basis
(dry = dried granule, water excluded; wet = includes water).

## Data completeness

Numbers are never invented. Every report carries a `Data completeness` line:

| Tag | Meaning | Count |
|---|---|---|
| `Complete` | value(s) taken directly from the master | 136 |
| `Deduced: <fields>` | value present, but the source Flag/interpretation note shows it was computed/interpreted (a target-driven dose, a split ratio, a renormalisation, etc.) rather than a raw recorded figure | 46 |
| `Needs input: <fields>` | genuinely missing — no dose recorded, a guarantee not yet filled in the Guarantees sheet, or the source note explicitly asks for confirmation | 15 |

The 15 `Needs input` formulations (see the `Needs Input` sheet for the exact field per Fo):

- **Fo40, Fo41** — coating dose not recorded (Humic W Coat, PhosCoat).
- **Fo87–Fo91** — Phosphoric acid's P2O5 guarantee is not filled in the Guarantees sheet;
  the P2O5 (%) field is marked `guarantee pending` rather than showing the partial,
  Rock-phosphate-only number as if it were final (that partial number is still shown,
  transparently, in the row's Note).
- **Fo146, Fo147** — the source note explicitly asks to confirm a 12.5% Zn dilution
  assumption.
- **Fo166–Fo171** — partial-acidulation series; the RP:H2SO4 ratio (and so both ingredient
  doses) is not recorded.

Classification is data-driven, not a hardcoded Fo list: a blank `Amount` cell anywhere in
`Formulations` always triggers `needs_input` for that ingredient, and the source
`Flag`/`Flag / interpretation` text is scanned for "to be filled" / "please confirm"
(needs input) or target/deduction/dilution language (deduced). Re-running the script
against a future, more complete version of the master will re-classify automatically.

## Formulations that use another formulation as an ingredient

72 recipe rows across 65 formulations use another formulation as an ingredient
(`Role = Sub-formulation`, e.g. Fo51's recipe is 96.15% TSP + 3.85% Fo50), including
two-level chains (e.g. Fo124 uses Fo123, which itself uses Fo121).

The master's own `Summary` mineral/biostimulant guarantee percentages (the 17 tracked
actives: N, P2O5, K2O, S, Ca, Mg, Zn, B, Mn, Cu, Mo, Se, Fe, Si, Humic acid, Fulvic acid,
Seaweed extract) already fully and correctly resolve these chains — verified
independently by recomputing every (formulation, active) pair bottom-up from the raw
recipe rows only, with zero mismatches across all 197 x 17 = 3,349 pairs. Those figures
are used as-is, exactly as before.

What the master does *not* propagate through a sub-formulation reference is the
free-text `Raw / self-active` field (non-tracked raw materials: coatings, biopolymers,
acids used as pH adjusters, etc.). A parent formulation's own `Summary` row simply omits
a raw/self-active item that only exists inside a referenced sub-formulation — e.g. Fo51's
own row lists no raw/self-active items at all, even though it contains Eco Coat OW 560
and Bountigel via its 3.85%-of-blend share of Fo50 (which is itself 50% Eco Coat OW 560 /
50% Bountigel).

53 of the 197 reports carry at least one raw/self-active line affected by this. For those,
the report now adds the sub-formulation's own raw/self-active composition, scaled by that
ingredient row's `Pct_of_blend` (recursively, so a two-level chain folds in correctly) and
merged by name with the formulation's own direct items. Every added line carries a note
naming the source sub-formulation; nothing is invented — each added percentage is the
product of two numbers already in the master (the sub-formulation's own recorded
composition, and its recorded share of this blend).

## Regenerating

```
pip install openpyxl pandas
python3 generate_reports.py
```

Requires `FORMULATIONS_DATABASE_V2.xlsm` in the same folder (not committed here — it is
the live master workbook, kept outside this repo). The script only ever reads it
(`data_only=True`); it writes a brand-new workbook and never touches the master.

## Source sheets used

`Summary` (one row per formulation — dates, site, process, final-blend active %,
protocol, protocol check, flags) for everything except the recipe table; `Formulations`
(one row per ingredient) for the recipe lines, sorted by `Step`. `Guarantees` and
`Ingredients` were not needed directly — `Summary`/`Formulations` already carry the
computed final-product percentages on the correct dry/wet basis, used as-is per the brief
(never recomputed or re-rounded).