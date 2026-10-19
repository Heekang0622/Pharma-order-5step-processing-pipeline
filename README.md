# Weekly Multi-Channel Order Pipeline (Excel VBA)

An Excel VBA module (~1,500 lines) that processes weekly radiopharmaceutical orders in five steps. Each step is a separate macro that asks for the week number and writes that week's results into the matching sheet.

> **Note:** This is a sanitized/anonymized version of a tool used in a real workflow. All hospital/client codes and the distributor's name have been replaced with generic placeholders (`HOSP01`-`HOSP25`, `PARTNER_CO`), and shipment-route labels have been generalized (`GroupA` / `GroupB` / `GroupC`). Business logic and structure are preserved as-is. Some inline comments remain in Korean, the original working language.

## The 5 steps

| Step | Macro | What it does |
|---|---|---|
| 1 | `캡슐데이터_회사폼_자동생성` | Reads the order sheets and fills the company form (`Sheet1`) by group and weekday, inserting rows when a block runs out of slots |
| 2 | `Matrix_신규Week_추가` | Adds a new week (Diag / Ther / Solution columns) to the Matrix sheet, flags new and dropped hospitals by color, updates subtotals |
| 3 | `Step3_PerHospital_추가` | Appends the week's rows to the "Per hospital" log (by shipment, hospital, mCi) and refreshes the pivot |
| 4 | `Step4_ContainerLog_추가` | Appends the week's shipment quantities to the container log |
| 5 | `Step5_Summary_추가` | Appends the week's block (25 mCi, 5 mCi, 2 mCi, Solution, all other shipments) to the Summary sheet |

The pivot refresh is also available separately as `Step3b_PerHospital_Pivot_갱신`. Existing weeks are never overwritten silently: if the week already exists, the macro asks before replacing it.

## Input data

Two sheets hold the raw orders:

- `주문데이터` (first-round orders)
- `2차주문데이터` (second-round orders)

Columns used: A = client code, B = weekday, C = mCi, D = quantity.

Client grouping is defined once at the top of the module (`GROUP_A_LIST`, `GROUP_B_LIST`). Orders from `PARTNER_CO` are treated as the solution product.

## How to use

1. Open the Excel file that contains the sheets referenced in the constants at the top of the module (`SH_*`).
2. Press `Alt+F11` → File → Import File → select `OrderPipeline_MultiChannel.bas`.
3. Press `Alt+F8`, choose a macro, and enter the week number when prompted.
4. Run the steps in order (1 → 5).

If Korean text is garbled after importing, save the `.bas` file as ANSI (CP949) and import again.

## Configuration

At the top of the module:

- `GROUP_A_LIST` / `GROUP_B_LIST`: client codes per group
- `SOLUTION_KEY`: client code treated as the solution product
- `SH_*`: sheet names
- `ANCHOR_MON`: Monday of week 1, used to calculate the month of each week
- `MATRIX_INCLUDE_2ND`: whether the Matrix step includes second-round orders

## Notes

- No real business data is included. Excel files are not part of this repository.
- The code was written for one specific workbook layout. Row/column positions and sheet names will need adjusting for other files.
