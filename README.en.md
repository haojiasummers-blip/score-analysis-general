# Score Report Analysis (score-analysis-general)

Turn one or two school exam score reports into a full statistical workbook:
**13 sheets** for a single exam, **8 sheets** for two-exam value-added analysis.
Pure `pandas` + `openpyxl`, no data dependencies.

## What it solves

School score reports usually come with a **two-row merged header**
(row 1 = subject names, row 2 = `raw score / scaled score / class rank / school rank`).
Reading them naively shifts every column. This skill detects the header shape,
flattens it, runs 12 analysis layers, and prints validation anchors you must check.

The second problem is subtler: **a rise in raw scores is often just an easier paper.**
For two exams the skill reports both raw-score and standard-score deltas,
plus a quick discriminator: if the raw-score "rise rate" is far from 50% while the
standard-score rise rate sits near 50%, the change came from exam difficulty.

## Layout

```
SKILL.md                     skill description (read by the agent)
references/SOP.md            method handbook: 12 analysis layers, checklists, 20 pitfalls
scripts/cross_sectional.py   cross-sectional analysis (single exam)
scripts/longitudinal.py      value-added analysis (two exams)
examples/                    600-student synthetic dataset + expected output
```

## Install

Drop the folder into your skills directory (e.g. `~/.dsh/skills/`),
or just let the agent read `SKILL.md`.

## Usage

```bash
# single exam → 13 sheets
python scripts/cross_sectional.py report.xlsx [output.xlsx]

# two exams → 8 sheets (raw-score and standard-score views side by side)
python scripts/longitudinal.py examA.xlsx examB.xlsx [output.xlsx]

# try it on the synthetic example
python examples/make_example_data.py
python scripts/cross_sectional.py examples/样例_A期.xlsx
```

Requires `pandas`, `numpy`, `openpyxl` (`xlrd` for legacy `.xls`).

## Supported report shapes

- two-row merged header (most common)
- single-row header
- three-row header (extra title row)
- region column variants: 县 / 县区 / 区域 / 区县

## Validation anchors (printed by the scripts — check them by hand)

| Check | Expected |
|---|---|
| band counts vs valid N | equal for every subject |
| contribution shares | sum to 100.0% |
| gap decomposition | four subject gaps sum to the total gap |
| longitudinal identity | error < 0.5 for > 99% of rows |
| match rate | name overlap / larger cohort > 85% |

If any anchor fails, fix the data before writing a report.

## Data

**This repository contains no real score data.** See [DATA-POLICY.md](DATA-POLICY.md).
The `examples/` folder ships a generator that produces purely synthetic records
(`同学001`…`同学600`, fictional schools).

## License

MIT — see [LICENSE](LICENSE).
