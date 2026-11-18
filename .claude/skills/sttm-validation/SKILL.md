---
name: sttm-validation
description: End-to-end source-to-target data validation driven by an STTM (Source-to-Target Mapping) Excel document. Use when asked to validate, reconcile or compare a source extract against a target extract using an STTM/mapping spec, to interpret STTM transformation rules, or to produce the STTM_Validation_Report.xlsx. Covers STTM analysis, file profiling (TXT/CSV/Excel), key-based record matching, field-level transformation validation with PASS/FAIL/REVIEW statuses, and the Excel report.
---

# STTM Source-to-Target Validation

You (Claude) are responsible for **understanding the STTM and encoding its rules**. The Python
utility `scripts/validate_data.py` is only the deterministic execution engine: it reads files,
evaluates the rules you encoded, compares values and writes the Excel report. It never interprets
the STTM by itself.

Workflow at a glance:

1. Locate inputs -> 2. Analyse the STTM -> 3. Profile source/target files -> 4. Write the
validation config -> 5. `check-config` -> 6. `validate` -> 7. Verify reconciliation and report findings.

## 0. Non-negotiable rules

- **Never modify, rename, move or re-save anything in `input/`.** All generated files go to `output/`.
- **Never guess a transformation rule.** If the STTM wording is ambiguous, incomplete, contradicts
  the data, or needs data that was not provided, encode the mapping with `"review": "<reason>"`.
  It then produces REVIEW rows, never PASS.
- **Never fabricate expected values, mappings, keys, lookup tables or code lists.** A code/value
  mapping can only contain pairs that are written in the STTM (or in a reference file the user provided).
- **Never silently skip.** Every STTM mapping row must end up as a `mappings` entry, a
  `review` mapping, or an `unvalidated_rules` entry with a reason. Every record is accounted for by
  the engine (matched / missing / extra / duplicate / null key / malformed / excluded by filter).
- **A skipped or unevaluated validation is never PASS.**
- **Confidentiality:** do not print raw record values to the console unless needed to resolve a
  specific question. Profile output uses masked patterns (`9999-99-99`, `Aaaa`); use
  `--show-values N` sparingly and only for columns whose format you must confirm. Do not copy data
  values into logs, the config `notes`, or your chat summary beyond a few illustrative examples.
- **No database or network access** unless the user explicitly configures and authorises it in this
  conversation. Work only with files in `input/`.
- Ask the user only when genuinely blocked (e.g. cannot tell which file is source vs target, or the
  STTM key is not identifiable). Otherwise proceed and record assumptions.

## 1. Environment

Run all commands from the project root. Python interpreter:

- Windows: `.venv\Scripts\python.exe` (Git Bash: `.venv/Scripts/python.exe`)
- macOS/Linux: `.venv/bin/python`
- If `.venv` does not exist: `python -m venv .venv` then `<venv python> -m pip install -r requirements.txt`
  (ask the user before installing if installation is not already permitted).

Below, `PY` means the venv interpreter and `VD` means `scripts/validate_data.py`.

| Command | Purpose |
|---|---|
| `PY VD discover` | List files in `input/` with size, sheets and a filename-based role guess |
| `PY VD inspect-sttm input/<sttm>.xlsx` | Write every STTM sheet, with row/column references, merged ranges and cell comments, to `output/work/sttm_dump.md` |
| `PY VD profile input/<file> [--sheet S] [--delimiter "\|"] [--header true\|false] [--key COL1,COL2] [--show-values N]` | Encoding, delimiter, header detection, columns, record count, malformed rows, null counts, leading zeros, masked value patterns, key null/duplicate check |
| `PY VD check-config` | Validate `output/work/validation_config.json` (files, columns, expressions) without running |
| `PY VD validate` | Run validation; writes `output/STTM_Validation_Report.xlsx` and `output/work/validation_summary.json` |

All commands accept `--work-dir` (default `output/work`) and `--verbose`. Relative paths in the
config resolve against the **project root**.

## 2. Locate the inputs

1. Run `discover`. Expected: one STTM workbook plus one or more source and target extracts. There
   may also be reference/lookup files.
2. Confirm roles from file names, sheet names and (after step 3) column names. If two plausible
   assignments remain, ask the user.
3. Unsupported formats (PDF, Word, zipped files): tell the user; do not convert inside `input/`.

## 3. STTM analysis (A)

1. Run `inspect-sttm` and read `output/work/sttm_dump.md` **completely** (use Read with offset/limit
   for large dumps). Every cell is shown as `R<row>: <Col>="value"` so you can cite references.
2. Classify every worksheet: mapping sheet(s), table/layout definitions, lookup/code lists, change
   log, cover/instructions. Read all of them; rules are often spread across sheets (e.g. a "Lookups"
   sheet, footnotes, cell comments, merged header cells).
3. Find the header row of each mapping sheet (it is often not row 1). Typical columns: source
   system/table/column/datatype/length, target table/column/datatype/length/nullable, PK/key flag,
   transformation/business rule, default value, comments.
4. For **every** mapping row record:
   - source table(s) & column(s), target table & column, data types/lengths
   - the exact rule text (keep it verbatim in `rule_text`)
   - STTM reference as `<Sheet>!R<row>` (e.g. `Mapping!R27`) in `sttm_ref`
5. Identify:
   - **Primary / business keys**: PK flags, "key", "unique", "natural key" wording, or target PK
     constraints. Composite keys are supported. If no key is identifiable, ask the user; do not invent one.
   - **Default values** ("default 'N'", "if null then 0", "set to 'UNKNOWN'").
   - **Filters** ("only active records", "where STATUS <> 'D'", "exclude test accounts").
   - **Joins** (source enrichment from another source table) and **lookups** (code -> description,
     reference tables). If the lookup data is not in `input/`, those mappings become REVIEW.
   - **Null handling** ("blank to null", "NVL", "COALESCE", "if spaces then null").
   - Run-time / non-deterministic values (load timestamp, batch id, sequence/surrogate keys,
     `SYSDATE`, user id): these are REVIEW (`review` reason), unless the STTM gives a verifiable rule.
6. Write a short mapping inventory for yourself (rows, keys, rule types) before encoding. Count the
   mapping rows so you can prove every one was encoded.

## 4. File analysis (B)

Run `profile` on every source, target and reference file. Check:

- **JSON extracts** (`.json`, `.jsonl`, or a `.txt` whose content starts with `[`/`{`) are read with
  `"format": "json"` (auto-detected). Supported: an array of objects, an object wrapping one array (or
  set `"json_path": "data.items"`), and JSON Lines. Nested objects are flattened to dotted column names
  (`payment.amount`); reference them as `{"col": "payment.amount"}` (a dotted name is only treated as
  `alias.column` when the prefix is a declared alias). Numbers keep their exact source text (`1500` vs
  `1500.00`), so use a numeric compare for amounts. JSON `null` and absent keys become empty (null);
  nested arrays are kept as compact JSON text. "Line/Row" in the report is the JSON record number.
- **Format & encoding**: detected automatically (UTF-8, UTF-8 BOM, UTF-16, cp1252, latin-1). Set
  `encoding` explicitly if the profile shows a fallback warning or garbled characters.
- **Delimiter**: auto-detected from `, | TAB ; ~ ^ \x01`. Multi-character delimiters (e.g. `||`) must
  be configured explicitly (no quote handling for those). Fixed-width files need a `fixed_width` layout,
  usually from the STTM layout sheet.
- **Header**: auto-detection is heuristic. Always confirm it against the STTM column names and set
  `header` explicitly in the config. Headerless files need `columns` (ordered names from the STTM layout).
- **Leading zeros / precision**: all values are read as text, so `00123` and `1.50` are preserved.
  Excel inputs: numbers stored as numbers in the workbook have already lost leading zeros (except
  cells with an all-zero number format like `00000`, which are re-padded). Integer-valued floats
  are rendered without `.0`; dates as `YYYY-MM-DD` (or `YYYY-MM-DD HH:MM:SS`). Note this in assumptions.
- **Malformed records** (wrong field count), **blank lines**, trailing delimiters, unbalanced quotes
  (warning about records spanning many lines -> consider `quotechar: null`), header/trailer records
  (`skip_rows`, `skip_footer`).
- **Keys**: use `--key` to see null and duplicate keys before validating.
- **Nulls**: look at `literal_NULL_tokens` and `null_or_blank`. Decide `null_handling` from the STTM
  and the data. Default: empty/whitespace-only = NULL. Add `"NULL"` to `null_tokens` only if the files
  really use that literal for nulls.
- **Dates/decimals**: use the masked `top_patterns` to infer formats (e.g. `99/99/9999`). If a pattern
  is ambiguous (dd/mm vs mm/dd), check values with `--show-values 5` or ask the user. Never assume.

## 5. Write the validation config

Write `output/work/validation_config.json` (create the folder if needed). Structure:

```json
{
  "sttm_file": "input/<sttm>.xlsx",
  "report_path": "output/STTM_Validation_Report.xlsx",
  "null_handling": {"null_tokens": [""], "blank_is_null": true, "tokens_case_insensitive": false},
  "assumptions": ["free text - every assumption you made, for the final summary"],
  "datasets": {
    "SRC_CUSTOMER": {"path": "input/<source file>", "format": "delimited", "delimiter": "|",
                     "encoding": "utf-8", "header": true, "quotechar": "\"", "table": "<STTM source table name>"},
    "TGT_CUSTOMER": {"path": "input/<target file>.xlsx", "format": "excel", "sheet": "Sheet1",
                     "header": true, "table": "<STTM target table name>"},
    "REF_COUNTRY":  {"path": "input/<lookup file>.csv", "header": true}
  },
  "validations": [
    {
      "name": "CUSTOMER",
      "source": {
        "dataset": "SRC_CUSTOMER", "alias": "S",
        "joins": [{"dataset": "SRC_ADDRESS", "alias": "A", "type": "left",
                   "on": [{"left": "S.CUST_ID", "right": "CUST_ID"}],
                   "rule_id": "J1", "sttm_ref": "Mapping!R5", "rule_text": "verbatim STTM join text"}],
        "filter": {"op": "ne", "left": {"col": "S.STATUS"}, "right": "D"},
        "filter_text": "verbatim STTM filter text", "filter_sttm_ref": "Mapping!R4"
      },
      "target": {"dataset": "TGT_CUSTOMER", "alias": "T", "filter": null},
      "keys": [{"name": "CUSTOMER_ID", "source": "CUST_ID", "target": "CUSTOMER_ID",
                "sttm_ref": "Mapping!R8", "notes": "PK per STTM"}],
      "key_match": {"trim": false, "case_insensitive": false},
      "mappings": [
        {"rule_id": "R009", "sttm_ref": "Mapping!R9", "rule_text": "verbatim STTM rule",
         "target_column": "FULL_NAME",
         "expression": {"op": "upper", "arg": {"op": "concat", "sep": " ",
                        "args": [{"op": "trim", "arg": {"col": "FNAME"}}, {"op": "trim", "arg": {"col": "LNAME"}}]}},
         "compare": {"type": "string"}, "notes": "assumption if any"},
        {"rule_id": "R010", "sttm_ref": "Mapping!R10", "rule_text": "Populate with load timestamp",
         "target_column": "LOAD_TS", "review": "Run-time value - cannot be deterministically validated"}
      ],
      "unvalidated_rules": [
        {"rule_id": "R031", "sttm_ref": "Mapping!R31", "rule_text": "...", "source": "ADDR.ZIP",
         "target": "DIM_ADDRESS.ZIP_CD", "reason": "Target table DIM_ADDRESS extract not provided"}
      ]
    }
  ]
}
```

Notes:
- One `validations` entry per target table (or per source->target pair). Several are allowed; the
  report shows each as a "Validation Set".
- `rule_id`: use the STTM's own rule/row id if it has one, else `R<row>`. Must be unique per set.
- `source_columns` (optional) lists the source columns to display in the report; by default it is
  derived from the `col` references in the expression.
- `transformation` (optional) overrides the auto-generated rule description in the report.
- Keys: `source`/`target` accept a **column name string** or an **expression** (e.g. `lpad` when
  the STTM says the key is zero-padded in the target).
- In expressions, **bare JSON strings/numbers are literals**. Columns are always `{"col": "NAME"}`
  or `{"col": "ALIAS.NAME"}`. Unqualified names resolve to the primary source, then to a uniquely
  matching joined dataset. Column matching is exact, then case-insensitive.

### Expression reference (E. transformation validation)

Null in -> null out for all string functions. `null` means "is NULL under `null_handling`".

| STTM rule | Expression |
|---|---|
| Direct move | `{"col": "X"}` |
| Constant / hard-coded value | `{"lit": "ABC"}` or `"ABC"`; NULL: `{"op": "null"}` |
| Trim / LTRIM / RTRIM | `{"op": "trim", "arg": E}` (`ltrim`, `rtrim`; optional `"chars"`) |
| Upper / lower | `{"op": "upper", "arg": E}`, `{"op": "lower", "arg": E}` |
| Concatenation | `{"op": "concat", "args": [E1, E2], "sep": " ", "nulls": "empty"\|"skip"\|"propagate"}` (Oracle `\|\|` = `empty`; SQL Server `+` = `propagate`; `CONCAT_WS` = `skip`) |
| Substring | `{"op": "substr", "arg": E, "start": 1, "length": 3}` (1-based like SQL; negative start counts from the end) |
| Left / right | `{"op": "left", "arg": E, "length": n}`, `{"op": "right", ...}` |
| Padding | `{"op": "lpad", "arg": E, "length": 10, "char": "0", "truncate": false}` (`rpad`) |
| Replace / regex | `{"op": "replace", "arg": E, "old": "-", "new": ""}`, `{"op": "regex_replace", "arg": E, "pattern": "...", "replacement": "..."}`, `{"op": "regex_extract", "arg": E, "pattern": "...", "group": 1}` |
| Length | `{"op": "length", "arg": E}` |
| Date conversion | `{"op": "date", "arg": E, "in_format": "%Y%m%d" or [list], "out_format": "%Y-%m-%d"}` (Python strptime codes: `YYYY`->`%Y`, `MM`->`%m`, `DD`->`%d`, `HH24`->`%H`, `MI`->`%M`, `SS`->`%S`, `MON`->`%b`). Unparseable source value -> REVIEW + data-quality exception |
| Numeric conversion / rounding | `{"op": "number", "arg": E, "scale": 2, "rounding": "HALF_UP", "thousands_sep": ",", "decimal_sep": ".", "strip_chars": "$ ", "implied_decimals": 2, "strip_trailing_zeros": false}`. Roundings: HALF_UP, HALF_EVEN, HALF_DOWN, DOWN/TRUNCATE, UP, FLOOR, CEILING. Supports `(123)` and trailing-minus negatives |
| Arithmetic / derived values | `{"op": "arith", "operator": "+", "args": [E1, E2], "scale": 2}` (`+ - * /`; null propagates; divide-by-zero -> REVIEW) |
| Default if null (NVL) | `{"op": "default", "arg": E, "value": "N"}` |
| COALESCE | `{"op": "coalesce", "args": [E1, E2, "UNKNOWN"]}` |
| NULLIF | `{"op": "nullif", "arg": E, "value": "N/A"}` |
| Conditional (CASE/IF) | `{"op": "case", "cases": [{"when": C, "then": E}], "else": E}` (no `else` -> NULL) |
| Code/value mapping (DECODE) | `{"op": "map", "arg": E, "values": {"M": "Male", "F": "Female", "__NULL__": "Unknown"}, "default": E, "on_missing": "review"\|"null"\|"keep", "trim": false, "case_insensitive": false}`. Use `default` only when the STTM states the else-value; otherwise uncovered codes are REVIEW |
| Lookup to reference data | `{"op": "lookup", "dataset": "REF_X", "match": [{"column": "CODE", "value": E}], "return_column": "DESC", "on_missing": "review"\|"null", "on_multiple": "review"\|"first", "trim_keys": false, "case_insensitive": false}`. Use `"null"` only if the STTM says unmatched -> NULL |
| Join (enrichment) | `source.joins` with `type` `left` or `inner`; then reference `{"col": "A.COLUMN"}`. Left join with no match -> joined columns NULL; multiple matches -> REVIEW. Inner join with no match -> record excluded from scope (and must be absent from target) |
| Filter | `source.filter` = condition C. Records failing it are excluded; if their key appears in the target the record is reported as an extra record (FAIL). `target.filter` limits which target rows are in scope (e.g. one source system in a shared table) |

Conditions C: `eq ne gt ge lt le` (`left`, `right`, optional `numeric`, `trim`, `case_insensitive`),
`in` / `not_in` (`arg`, `values`), `is_null` / `not_null` (`arg`), `like` (`arg`, SQL `pattern` with
`%`/`_`), `regex` (`arg`, `pattern`, `fullmatch`), `between` (`arg`, `low`, `high`, `numeric`),
`and` / `or` (`args`), `not` (`arg`), `true` / `false`. SQL semantics: any comparison with NULL is
false; use `is_null` explicitly.

### Compare reference (D. field-level validation)

| `compare` | Use when |
|---|---|
| `{"type": "string"}` (default) | Exact text comparison. Options: `"trim": true\|"right"\|"left"` (only if the target is a CHAR-padded column or the STTM says so), `"case_insensitive": true` (only if the STTM allows it) |
| `{"type": "numeric", "tolerance": "0.01", "thousands_sep": ","}` | Numeric target columns (`10.5` = `10.50`). Tolerance only when the STTM states a rounding tolerance |
| `{"type": "date", "expected_formats": [...], "actual_formats": ["%d/%m/%Y"], "date_only": false}` | Target dates in a different text format than the expected value. Default formats are ISO-like only |

Both sides are null-normalised with `null_handling` first: NULL vs NULL = PASS, value vs NULL = FAIL.
Mismatch reasons are diagnosed automatically (whitespace-only, case-only, leading zeros, numeric
format, truncation).

### Statuses (F)

- **PASS**: expected value computed from the STTM rule equals the actual target value.
- **FAIL**: definite mismatch or rule violation (wrong value, value vs NULL, missing/extra record,
  duplicate or null key, malformed record, target column missing, excluded record present in target).
- **REVIEW**: mapping unclear (`review`), reference data missing, code not covered by the STTM
  mapping, source value cannot be converted, ambiguous join/lookup, rule not evaluable. Never PASS.

When to use `review` vs `unvalidated_rules`:
- Target column exists in the target extract but the rule can't be encoded -> mapping with `review`
  (REVIEW for every matched record).
- Target table/column is not present in any provided extract -> `unvalidated_rules` with a reason.
  But if the STTM says a column should exist in a provided target file and it does not, encode the
  mapping normally: the engine reports FAIL "target column not present".

## 6. Check, then run

1. `PY VD check-config`. Fix every **error**. Read every **warning**: a missing target column or
   source column may be a typo on your side (fix it) or a real defect (keep it). Confirm the
   "Target columns not covered" list contains only columns the STTM does not map (audit columns,
   surrogate keys) - otherwise you missed an STTM row.
2. `PY VD validate`. Runtime is roughly 2,000-3,000 field validations per second (dominated by Excel
   writing); for more than ~500k field validations, tell the user it will take several minutes and
   run it in the background. Sheets exceeding Excel's 1,048,575 data rows continue in
   `Field Validation (2)`, etc.
3. If the report file is locked (open in Excel), the engine writes a timestamped copy instead and
   says so - tell the user.

## 7. Verify and report

1. Read `output/work/validation_summary.json`. `reconciliation_ok` must be `true`. If not, investigate
   before reporting; never present unreconciled numbers as final.
2. Sanity-check the results against the STTM: e.g. a rule with 100% FAIL usually means a mis-encoded
   rule or a wrong compare type. Re-read the STTM row. Change the config only if your encoding was
   wrong, not to make results pass. Record any change as an assumption.
3. Report to the user:
   - files used (STTM, source, target, reference) and the report path
   - overall result and per-set PASS/FAIL/REVIEW counts; source/target counts, matched, missing, extra,
     duplicate and null keys, malformed records
   - top failing rules (rule id, target column, fail count, issue categories)
   - REVIEW items and why (unclear rules, missing reference data, unconvertible values)
   - every assumption and limitation (header decisions, null tokens, date formats, Excel numeric
     rendering, out-of-scope rules)
   - recommended next actions (clarify STTM rows X/Y, provide lookup file Z, fix extract issues)

## Report contents (generated by the engine)

`output/STTM_Validation_Report.xlsx`:
1. **Summary**: files, run info, overall result, per-set statistics (source/target counts, matched,
   missing, extra, duplicates, null keys, PASS/FAIL/REVIEW counts as live `COUNTIFS` formulas over the
   detail sheets), reconciliation checks, status definitions.
2. **Record Validation**: one row per key (matched/missing/extra/duplicate), per null-key record, per
   malformed record and per filter-excluded record, with source/target status, line numbers, field
   counts, result and failure reason.
3. **Field Validation**: one row per matched record x mapping: key, source table/column, target
   table/column, source value, expected value, actual value, transformation rule, status, reason, STTM reference.
4. **Exceptions**: malformed records, null keys, duplicate keys, missing records, extra records, failed
   records, field mismatches, data-quality issues, review items, unvalidated STTM rules.
5. **Rule Log**: every key, join, filter, field mapping and unvalidated rule with the STTM rule text,
   transformation applied, validation method, PASS/FAIL/REVIEW counts, result and notes; plus the list of
   unmapped target columns.

`<NULL>` = null value, `<EMPTY>` = empty string, `(not evaluated)` = expected value could not be computed.

## Known limitations

- Excel inputs: numeric cells lose leading zeros at source (unless zero-padded number format);
  formulas are read as their cached values; only the configured sheet is read.
- Fixed-width layouts must be supplied from the STTM; no automatic layout detection.
- Expression language covers the common ETL functions above; anything else (aggregations across
  records, windowing, SCD history logic, sequences) must be marked REVIEW with an explanation.
- Records are held in memory; files up to a few million rows are practical.
