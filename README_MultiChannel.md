# Weekly Multi-Channel Order Pipeline (Excel VBA)

A single Excel VBA module (~1,000 lines) that replaced roughly a month of
recurring manual data-entry work with a set of one-click macros. Built after
being handed a fully manual weekly workflow and immediately recognizing that
every step of it was mechanical enough to automate.

> **Note:** This is a sanitized/anonymized version of a production tool used
> in my actual job. All hospital/client codes and the distributor's name have
> been replaced with generic placeholders (`HOSP01`-`HOSP24`, `PARTNER_CO`),
> and shipment-route labels have been generalized (`GroupA`/`GroupB`/`GroupC`).
> Business logic and structure are preserved as-is. Some inline comments
> remain in Korean, the original working language.

## Background

A colleague had built (and maintained entirely by hand) a set of weekly
Excel reports tracking radiopharmaceutical capsule orders across two order
channels and roughly two dozen hospital clients grouped into shipment
batches. By her own account, getting the manual process to a point where it
ran without errors took about a month. I inherited this workflow, and
within a day of looking at it, rebuilt the entire thing as VBA macros that
reproduce the same output.

## What the pipeline does

| Function | What it does |
|---|---|
| `CapsuleData_FirstRound_Generate` | Reads two raw order-data sheets, buckets every order line by client group + delivery day, and fills a pre-formatted weekly capsule sheet — including the low-dose (2/5mCi) split that has to route to a different cell block than standard doses |
| `Matrix_신규Week_추가` | Adds a new week's column block to a running order-frequency matrix; grows merged header/data ranges safely (3-stage unmerge → insert → re-merge) instead of breaking existing merged cells |
| `InsertHospitalAtPosition` | Inserts a newly-onboarded client at a fixed position in the matrix (not just appended at the end) without disturbing the row numbering of existing clients |
| `Step3_PerHospital_추가` / `Step3b_PerHospital_Pivot_갱신` | Appends this week's per-client mCi/quantity totals to a running log, then refreshes the pivot table that summarizes it — the pivot's source range is reset dynamically each run instead of staying hardcoded |
| `Step4_ContainerLog_추가` | Logs this week's total shipment quantities by route (3 fixed rows: two "Group B" shipments and one "Group A" shipment), skipping any route with zero quantity |
| `Step5_Summary_추가` | Aggregates the week into a single summary block by dose tier (100/50/30/25/5/2 mCi + solution), computing high-dose/low-dose/all-shipment rollups |

## Why this was worth automating

The original manual process wasn't just slow — its error rate came from the
sheer repetition (tens of raw order lines every week, expanded and
re-checked by hand) rather than any real complexity in the underlying
logic. Once written down as rules, the whole thing collapsed into deterministic
steps:

- **Bucket-then-fill instead of read-then-type**: orders are first grouped
  into an in-memory dictionary keyed by (client group, day), then written to
  the sheet in one pass — so a client with orders on three different days
  ends up in the right three cells without anyone manually cross-referencing
  a calendar.
- **Position-aware insertion**: `InsertHospitalAtPosition` and `GrowBlock`
  handle the case that a client needs to be spliced into row 19 of an
  existing 24-client matrix — not appended — while keeping every merged
  cell and every other client's row references intact.
- **Self-resetting pivot source**: rather than a pivot table whose source
  range silently goes stale as new weeks are appended, the macro
  recalculates the correct range and re-points the pivot every run.
- **Idempotent by construction**: re-running any step clears/un-merges its
  own target range first, so nothing has to be manually undone if a step is
  run twice.

## Stack

Excel VBA only — no external dependencies, designed to run in a
locked-down corporate environment where installing a script interpreter
isn't an option.

## Related project

See [pharma-order-processing-pipeline pattern in `order-automation-portfolio`](https://github.com/Heekang0622/order-automation-portfolio)
for a from-scratch, fully public-data demo of the PDF-extraction +
cross-validation half of this same overall workflow (Python + VBA), which
feeds into a pipeline like this one.
