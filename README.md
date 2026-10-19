# Weekly Multi-Channel Order Pipeline (Excel VBA)

An Excel VBA module (~1,500 lines) that processes weekly radiopharmaceutical orders in five steps. Each step is a separate one-click macro: it asks for the week number, reads the two order sheets, and writes that week's results into its own target sheet.

> **Note:** This is a sanitized/anonymized version of a tool used in a real workflow. All hospital/client codes and the distributor's name have been replaced with generic placeholders (`HOSP01`-`HOSP25`, `PARTNER_CO`), and shipment-route labels have been generalized (`GroupA` / `GroupB` / `GroupC`). Business logic and structure are preserved as-is. Some inline comments remain in Korean, the original working language.

## Overview

```
 주문데이터 (1st round) ─┐
                         ├─► Step 1  Company form (Sheet1)
 2차주문데이터 (2nd round)┤   Step 2  Matrix (new week columns)
                         ├─► Step 3  Per hospital log  ─► pivot refresh
                         ├─► Step 4  Container log
                         └─► Step 5  Summary
```

The order sheets use columns **A** = client code, **B** = weekday, **C** = mCi, **D** = quantity. Steps 2-5 read the orders through one shared loader, so every step sees exactly the same data.

## Step-by-step

### Step 1 - Company form  (`캡슐데이터_회사폼_자동생성`)
Rebuilds the weekly company form (`Sheet1`) from scratch every run.

- Clears the form area, then sorts every order into a bucket by **group and weekday**:
  - GroupA / GroupB high-dose capsules (100, 50, 30, 25 mCi)
  - GroupA / GroupB low-dose capsules (5, 2 mCi), kept in their own sections (`GroupA_Low`, `GroupB_Low`)
  - GroupC from the second-round sheet (100, 50, 30, 25 mCi)
  - Solution (200 / 100 mCi) in its own section
- Writes each section with its weekday blocks (Sun → Mon), the client codes, and `SUMIFS` formulas that pull the quantity per mCi tier straight from the order sheets.
- Highlights every cell that has a quantity > 0 in yellow, and adds a **Grand Total**.
- **Checks:** invalid weekdays, 10 mCi rows (no column in the template), unknown mCi values, clients missing from the group list, and a total check (orders vs. form). Anything found is listed in a single message at the end.

### Step 2 - Matrix  (`Matrix_신규Week_추가`)
Adds the new week to the Matrix sheet as three columns: **Diag / Ther / Solution** (`W38 Diag`, `W38 Ther`, `W38 Solution`).

- Inserts the three columns just before `Total`, copies the formatting of the previous week, and re-merges the title.
- Fills each hospital row (names matched ignoring spaces and upper/lower case): Diag = 2/5/10 mCi, Ther = 25/30/50/100/200 mCi, Solution = the solution product.
- **Color flags:** a hospital that appears this week but not last week is marked **green**, one that dropped out is marked **red**. Colors copied from the previous week are cleared first.
- Hospitals that are not in the Matrix yet are appended automatically (green). `HOSP24` is inserted at its fixed position if it is missing.
- Rebuilds the row totals, `Sub total` and `Total`. Optionally includes second-round orders (`MATRIX_INCLUDE_2ND`).

### Step 3 - Per hospital log  (`Step3_PerHospital_추가`)
Appends one row per **shipment × hospital × mCi** to the running "Per hospital" log.

- Shipments: `GroupB- 1st`, `GroupA- 1st`, `GroupB- 2nd`. Within a hospital, mCi values are sorted from high to low and written as `100mCi`, `50mCi`, ...
- Columns written: month, `Week NN`, shipment, hospital, mCi, quantity, plus a `SUM` formula at the start of each shipment block.
- After writing, the pivot table's source range is extended and refreshed (`Step3b_PerHospital_Pivot_갱신`, which can also be run by itself).
- **Month label:** follows the style already used in the sheet (English `September` or Korean `10월`), because the pivot's month filter matches on the text. Mixing styles would make a new week drop out of the pivot.

### Step 4 - Container log  (`Step4_ContainerLog_추가`)
Adds the week's shipment quantities to the container log (`2026 Containers+면장`).

- Orders are split into six shipment rows:
  `GroupB-01 high dose` / `GroupB-01 diagnosis` / `GroupB-01 sol` / `GroupA-01 high dose` / `GroupA-01 diagnosis` / `GroupB-02 caps`
- Rows are written above the `TOTAL` row, and rows are inserted if there is not enough room.
- The **container count (column E) and customs declaration (column G) are never touched**. They are typed in by hand, and the weekly container sum picks them up automatically.
- Month label is written once per week (merged), in the sheet's existing style.

### Step 5 - Summary  (`Step5_Summary_추가`)
Appends a block for the week at the bottom of the Summary sheet (`Week 40- OCTOBER`):

| Row | Content |
|---|---|
| 25mCi | all 25 mCi capsules |
| 5mCi / 2mCi | low-dose capsules |
| Solution | solution quantity |
| Capsules- All shipments | all remaining high-dose capsules (30 / 50 / 100 / 200 mCi) |

- Uses real cell-reference formulas (not hand-typed sums) and copies the formatting of the previous block.
- Unclassified quantities (for example 10 mCi) are reported as a warning.

## Rules shared by all steps

- **Run order:** 1 → 5. Each step can be re-run on its own.
- **Re-running a week is safe:** if the week already exists, the macro asks whether to overwrite it. Answering "No" leaves the sheet untouched. Existing weeks are never changed silently.
- **Solution detection:** a row is treated as solution if the client is `PARTNER_CO`, or the strength is 200 mCi. If one of the first two rows of the first-round sheet is 200 mCi, the 100 mCi in those two rows is also treated as solution. Rows classified this way are listed in the final message so they can be double-checked.
- **Week → month:** week 1 starts on the Monday set in `ANCHOR_MON`; the month of a week is the month of its **Sunday**.
- **Input checks:** week number must be 1-53; errors are caught and shown in a message instead of leaving the sheet half-written.

## How to use

1. Open the Excel file containing the sheets named in the constants at the top of the module (`SH_*`).
2. Press `Alt+F11` → File → Import File → select `OrderPipeline_MultiChannel.bas`.
3. Press `Alt+F8`, run the step you need, and enter the week number when prompted.

If Korean text is garbled after importing, save the `.bas` file as ANSI (CP949) and import again.

## Configuration (top of the module)

| Constant | Meaning |
|---|---|
| `GROUP_A_LIST` / `GROUP_B_LIST` | client codes per group |
| `SOLUTION_KEY` | client code treated as the solution product |
| `SH_*` | sheet names |
| `ANCHOR_MON` | Monday of week 1 |
| `MATRIX_INCLUDE_2ND` | include second-round orders in the Matrix step |

## Notes

- No real business data is included, and Excel files are not part of this repository.
- The code was written for one specific workbook layout. Row/column positions and sheet names will need adjusting for other files.
