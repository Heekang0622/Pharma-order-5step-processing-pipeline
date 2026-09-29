# Pharma Order Processing Pipeline (Excel VBA)

A multi-step Excel VBA macro pipeline that automates the weekly order-processing
workflow for a pharmaceutical import/export operation. Originally built to
replace a fully manual process of transcribing purchase orders into a
production-request format.

> **Note:** This is a sanitized/anonymized version of a production tool used in
> my actual job. All client names, agency codes, and partner company names have
> been replaced with generic placeholders (`CLIENT01`–`CLIENT31`, `PARTNER_CO`,
> etc.). Sheet names have been translated to English. Business logic and
> structure are preserved as-is.

## Problem

Each week, orders arrive from multiple sales channels (international
distributor orders, domestic direct orders, and a third reconciliation-based
channel) in inconsistent formats — one row per order line, with a quantity
column, rather than one row per physical unit needed for the internal
production-request template. Preparing the production request previously
required manually expanding, sorting, and grouping hundreds of rows by hand
every week, including error-prone manual cell merging based on package size
rules.

## What the pipeline does

| Step | Function | What it does |
|---|---|---|
| 0 | `Step0_Clear_Columns` | Resets the staging area before a new run |
| 1 | `Step1_AutoFill_Smart_Optimized` | Expands international-channel order rows (qty column) into one row per unit |
| 2 | `Step2_Final_Merge_Resolved` | Sorts by dose size → delivery day → client priority, then merges rows into physical "package" groups based on dose-denomination rules, and flags priority clients |
| 3 | `Step3_Domestic_Only_Final_Fixed` | Same expand/sort logic for the domestic order channel, appended below |
| 4 | `Step4_Solution_Mapping` | Expands a third order channel, standardizes free-text client names via a lookup table, then sorts |

## Design notes

- **Dynamic row expansion**: quantity-based rows are expanded into individual
  unit rows via array iteration rather than manual copy/paste, eliminating a
  common source of quantity errors.
- **Rule-based merging**: package grouping size (1/2/4 physical units per
  package) is driven by dose denomination, not manually judged per row.
- **Custom multi-key sort**: combines a numeric key (dose size), a
  domain-specific custom day-of-week order, and a client-priority custom order
  in a single sort operation.
- **Name normalization**: a separate lookup sheet maps inconsistent
  free-text client names to standardized display names via fuzzy `MATCH`.
- **Idempotent staging**: each step explicitly clears/un-merges its target
  range before writing, so macros can be safely re-run.

## Known limitations / next steps

- Business logic is currently split across VBA and manual PDF-reading steps
  upstream; a Python-based PDF parsing layer (pdfplumber) was later added to
  reduce upstream transcription errors, with plans to feed structured
  output directly into this pipeline.
- The client-priority sort list and package-size mapping are hardcoded;
  in a next iteration these would move to a config sheet or external file.
- No automated tests; validation is currently done by comparing generated
  totals against printed PO totals.

## Stack

Excel VBA (no external dependencies). Designed to run in a locked-down
corporate environment without local script/interpreter installation
permissions.
