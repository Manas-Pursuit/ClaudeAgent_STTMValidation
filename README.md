# STTM Source-to-Target Data Validation Agent

A Claude Code **custom agent** (`sttm-validator`) and **skill** (`sttm-validation`) that validate a
target data extract against a source extract, using the rules in an STTM (Source-to-Target Mapping)
Excel document. The result is a detailed Excel report.

Claude reads and interprets the STTM. A small internal Python engine performs the deterministic
comparisons and builds the workbook. This is not a web app or a standalone tool.

## Project structure

```
<project root>/
├── .claude/
│   ├── agents/
│   │   └── sttm-validator.md        # the agent (orchestrator)
│   └── skills/
│       └── sttm-validation/
│           └── SKILL.md             # validation playbook + config/expression reference
├── input/
│   └── README.md                    # put STTM, source and target files here (read-only)
├── output/
│   └── README.md                    # STTM_Validation_Report.xlsx and work/ files are generated here
├── scripts/
│   └── validate_data.py             # internal engine used by the agent
├── requirements.txt
├── CLAUDE.md
├── README.md
└── .gitignore
```

## Setup (one time)

Requires Python 3.9+ (tested with 3.15). Run these commands from the project root.

**Windows (PowerShell):**
```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

**macOS / Linux:**
```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

Dependencies: `openpyxl` (read .xlsx), `XlsxWriter` (write the report), `xlrd` (read legacy .xls).
All are pure Python.

## Usage

1. Copy your files into `input/` (see [input/README.md](input/README.md)):
   - the STTM workbook (for example `STTM_Customer.xlsx`)
   - the source extract (for example `src_customer.txt`)
   - the target extract (for example `tgt_dim_customer.csv`)
   - optional lookup or reference files
2. Open Claude Code in this folder (VS Code extension or `claude` in a terminal started here).
3. Invoke the agent in any of these ways:
   - `@agent-sttm-validator validate the files in input/`
   - In plain language: `Use the sttm-validator agent to validate the source and target files in input against the STTM`
   - As the whole session: `claude --agent sttm-validator`
   - Using the skill directly in the main conversation: `/sttm-validation`
4. The agent then:
   1. discovers the files and confirms their roles
   2. dumps and reads the whole STTM
   3. profiles the files
   4. encodes each STTM rule into `output/work/validation_config.json`
   5. checks the config
   6. runs the validation
   7. verifies reconciliation
   8. reports its findings
5. Open `output/STTM_Validation_Report.xlsx`.

### Example prompts

- `@agent-sttm-validator Validate input/src_customer.txt against input/tgt_customer.csv using input/STTM_Customer.xlsx. The key is CUSTOMER_ID.`
- `@agent-sttm-validator Re-run the validation. STTM row 27 means the target stores DOB as DD/MM/YYYY.`
- `@agent-sttm-validator Explain the REVIEW items in the last report and what information you need to resolve them.`

## Report

`output/STTM_Validation_Report.xlsx` contains these sheets:

| Sheet | Contents |
|---|---|
| **Summary** | STTM, source and target file names; run info and assumptions; record counts; matched, missing and extra records; duplicate and null keys; field PASS/FAIL/REVIEW counts (live formulas over the detail sheets); overall result; reconciliation checks |
| **Record Validation** | Per key: source and target status, line numbers, field counts, result, failure reason. Also lists null-key, malformed and filter-excluded records |
| **Field Validation** | Per matched record and mapped field: key, source and target table/column, source value, expected value, actual value, transformation rule, status, reason, STTM reference |
| **Exceptions** | Malformed records, null keys, duplicate keys, missing and extra records, failed records, field mismatches, data-quality issues, review items, unvalidated STTM rules |
| **Rule Log** | Every key, join, filter and mapping rule: STTM rule text, transformation applied, validation method, counts, result, notes; plus unmapped target columns |

The workbook has frozen headers, autofilters, sized columns and colour-coded statuses:
**PASS** (green), **FAIL** (red), **REVIEW** (amber).

### Status definitions

- **PASS**: the expected value from the STTM rule equals the actual target value.
- **FAIL**: a definite mismatch or rule violation, such as a wrong value, a missing or extra record, a duplicate or null key, or a malformed record.
- **REVIEW**: the mapping is unclear, reference data is missing, a value is not covered by the STTM, or the rule cannot be evaluated confidently. REVIEW is never counted as PASS.

## Supported transformations

The engine supports these transformations:

- direct mapping, constants, NULL
- trim, upper and lower case, concatenation (Oracle- and SQL Server-style null semantics), substring, left/right, padding, replace and regex
- date conversion (any strptime/strftime format), numeric conversion with scale, rounding, implied decimals and thousands separators
- arithmetic and derived values
- default, COALESCE and NULLIF null handling
- CASE/IF conditional logic, DECODE-style code maps, lookups against reference files
- left and inner joins, source filters and target scope filters
- string, numeric (with tolerance) and date comparison

The full reference is in [SKILL.md](.claude/skills/sttm-validation/SKILL.md).

## Safety and accuracy

- Input files are opened read-only, and the engine refuses to write inside `input/`.
- All values are read as text, so leading zeros, precision and formatting are preserved.
- Mapping rules, code values and expected values are never guessed. Unclear rules are marked REVIEW.
- Records are never skipped silently. Every record ends up counted as one of: matched, missing, extra, duplicate, null key, malformed, or excluded by filter. The Summary shows reconciliation checks that prove these counts add up.
- Logs and `validation_summary.json` contain counts only, never record values.
- The tool makes no database or network connections.

## Limitations

- Excel inputs: numbers stored as numeric cells have already lost their leading zeros, except cells with zero-padded number formats. Formulas are read as their cached values.
- Fixed-width files need a layout, usually taken from the STTM.
- Rules beyond the expression language (aggregations across records, SCD history, sequences or surrogate keys, run-time timestamps) are reported as REVIEW.
- Data is processed in memory. Files of a few million rows are practical.
- Writing the Excel file is the slowest step: about 2,500 field validations per second (about 7 minutes for 1.1 million). Sheets beyond Excel's row limit continue on `(2)`, `(3)`, … sheets.

## Engine CLI (used by the agent)

```
.venv\Scripts\python.exe scripts\validate_data.py discover
.venv\Scripts\python.exe scripts\validate_data.py inspect-sttm input\STTM.xlsx
.venv\Scripts\python.exe scripts\validate_data.py profile input\src.txt --key CUST_ID
.venv\Scripts\python.exe scripts\validate_data.py check-config
.venv\Scripts\python.exe scripts\validate_data.py validate
```
