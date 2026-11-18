# input/: place your validation files here

The `sttm-validator` agent reads files from this folder **read-only**. It never modifies, renames or
deletes anything here.

## Required files

| File | Formats | Notes |
|---|---|---|
| **STTM mapping document** | `.xlsx`, `.xlsm`, `.xls` (also `.csv`) | All worksheets are read: mapping sheets, lookups/code lists, layouts, notes and cell comments. |
| **Source extract** (database extract or proxy) | `.txt`, `.csv`, `.dat`, `.tsv`, `.psv`, `.xlsx`, `.xls` | One file per source table. Delimited (`,` `\|` TAB `;` `~` `^`), fixed-width (layout taken from the STTM) or Excel. |
| **Target extract** | same as source | One file per target table. |

## Optional files

- **Reference / lookup data** (for example a country-code table) needed by STTM lookup rules.
  Without it, those rules are reported as REVIEW and are not guessed.
- Additional source tables used in STTM joins.

## Naming tips (optional, helps auto-detection)

Include `STTM` or `mapping` in the STTM file name, `src`/`source` in source files and
`tgt`/`target` in target files, for example:

```
input/
  STTM_Customer_v1.2.xlsx
  src_customer.txt
  tgt_dim_customer.csv
  ref_country_codes.csv
```

If the names are not clear, the agent will ask you which file is which.

## Data handling notes

- Values are read as text, so leading zeros, decimals and formatting are preserved.
- Excel cells stored as *numbers* have already lost leading zeros inside the workbook. Prefer
  text-formatted cells or delimited extracts for key and code columns.
- Header and trailer records are supported. The agent detects them, or you can mention them.
- These files may be confidential. `.gitignore` excludes everything in this folder except this README.
