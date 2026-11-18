#!/usr/bin/env python3
"""
validate_data.py - internal execution utility for the `sttm-validator` Claude Code agent.

This script is NOT the user interface. The sttm-validator agent reads and interprets the
STTM, writes a machine-readable validation config (output/work/validation_config.json) that
encodes every STTM rule with its STTM reference, and then calls this utility to execute the
deterministic comparison and build the Excel report.

Sub-commands
------------
  discover       List candidate input files in input/ and guess their role (STTM/source/target).
  inspect-sttm   Dump every worksheet of the STTM (with cell references) to output/work/sttm_dump.md
                 so the agent can read and interpret it.
  profile        Profile a source/target/reference extract (encoding, delimiter, header, columns,
                 record counts, malformed records, masked value patterns, key duplicates).
  check-config   Validate a validation config (structure, files, columns, expressions) without running.
  validate       Execute the validation and write output/STTM_Validation_Report.xlsx plus
                 output/work/validation_summary.json (aggregate statistics only, no data values).

Safety guarantees
-----------------
  * Input files are only ever opened read-only; the script refuses to write inside input/.
  * Every value is read as text (no type inference) so leading zeros, precision and formatting
    are preserved exactly.
  * Nothing is silently skipped: malformed, blank, filtered, null-key and duplicate-key records
    are all counted and reported.
  * The log file contains counts and structural information only, never record values.
"""
from __future__ import annotations

import argparse
import csv
import datetime as _dt
import json
import logging
import math
import re
import sys
import time
from collections import Counter
from decimal import (ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_DOWN, ROUND_HALF_EVEN,
                     ROUND_HALF_UP, ROUND_UP, Decimal, InvalidOperation)
from pathlib import Path

__version__ = "1.0.0"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
REPORT_NAME = "STTM_Validation_Report.xlsx"
DEFAULT_REPORT = OUTPUT_DIR / REPORT_NAME

# Mutable so --work-dir can override it.
WORK = {"dir": OUTPUT_DIR / "work"}

EXCEL_MAX_DATA_ROWS = 1_048_575          # 1,048,576 rows minus the header row
EXCEL_MAX_CELL_CHARS = 32_767
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
TEXT_EXT = {".csv", ".txt", ".dat", ".tsv", ".psv", ".del", ".out"}

PASS, FAIL, REVIEW, INFO = "PASS", "FAIL", "REVIEW", "INFO"
NULL_DISPLAY, EMPTY_DISPLAY = "<NULL>", "<EMPTY>"
AMBIGUOUS = object()   # marker: a join returned more than one row

ROUNDING = {
    "HALF_UP": ROUND_HALF_UP, "HALF_EVEN": ROUND_HALF_EVEN, "HALF_DOWN": ROUND_HALF_DOWN,
    "DOWN": ROUND_DOWN, "UP": ROUND_UP, "FLOOR": ROUND_FLOOR, "CEILING": ROUND_CEILING,
    "TRUNCATE": ROUND_DOWN,
}
DEFAULT_DATE_FORMATS = [
    "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%S.%f", "%Y%m%d", "%Y%m%d%H%M%S",
]
DELIMITER_ALIASES = {"tab": "\t", "\\t": "\t", "pipe": "|", "comma": ",", "semicolon": ";",
                     "tilde": "~", "caret": "^", "\\x01": "\x01", "ctrl-a": "\x01"}
CANDIDATE_DELIMITERS = [",", "|", "\t", ";", "~", "^", "\x01", "\x1f"]

log = logging.getLogger("sttm_validation")
csv.field_size_limit(min(sys.maxsize, 2**31 - 1))


# --------------------------------------------------------------------------------------------
# Errors
# --------------------------------------------------------------------------------------------
class ConfigError(Exception):
    """The validation config is invalid."""


class InputFileError(Exception):
    """An input file cannot be read as specified."""


class ReviewNeeded(Exception):
    """A rule cannot be evaluated confidently for this record -> REVIEW (never PASS)."""

    def __init__(self, reason: str, category: str = "CANNOT_EVALUATE", dq: bool = False):
        super().__init__(reason)
        self.reason = reason
        self.category = category
        self.dq = dq


# --------------------------------------------------------------------------------------------
# General helpers
# --------------------------------------------------------------------------------------------
def work_dir() -> Path:
    WORK["dir"].mkdir(parents=True, exist_ok=True)
    return WORK["dir"]


def setup_logging(verbose: bool = False) -> None:
    if log.handlers:
        return
    log.setLevel(logging.DEBUG)
    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    log.addHandler(ch)
    try:
        fh = logging.FileHandler(work_dir() / "validation.log", encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(fh)
    except OSError as exc:  # logging must never stop a run
        log.warning("Could not open log file: %s", exc)


def resolve_path(p, base: Path | None = None) -> Path:
    path = Path(p)
    if not path.is_absolute():
        path = (base or PROJECT_ROOT) / path
    return path.resolve()


def ensure_safe_output(path: Path, protected: list[Path] | None = None) -> Path:
    """Refuse to write inside input/ or over any input file."""
    path = path.resolve()
    inp = INPUT_DIR.resolve()
    if path == inp or inp in path.parents:
        raise ConfigError(f"Refusing to write output inside the input folder: {path}")
    for p in protected or []:
        if path == p.resolve():
            raise ConfigError(f"Refusing to overwrite an input file: {path}")
    return path


def lit_str(v):
    if v is None:
        return None
    if isinstance(v, bool):
        raise ConfigError("Boolean literals are not supported; use the strings 'TRUE'/'FALSE'.")
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return fmt_decimal(Decimal(repr(v)))
    return str(v)


def fmt_decimal(d: Decimal) -> str:
    s = format(d, "f")
    if d == 0 and s.startswith("-"):
        s = s[1:]
    return s


def parse_decimal(v: str, thousands: str | None = None, decimal_sep: str = ".",
                  strip_chars: str = "") -> Decimal | None:
    """Strict decimal parser. Returns None if the text is not a number."""
    s = v.strip()
    if strip_chars:
        s = s.strip(strip_chars).strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    if s.endswith("-") and not s.startswith("-"):
        neg, s = True, s[:-1].strip()
    if thousands:
        s = s.replace(thousands, "")
    if decimal_sep != ".":
        s = s.replace(decimal_sep, ".")
    if not re.fullmatch(r"[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?", s):
        return None
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    if not d.is_finite():
        return None
    return -d if neg else d


def shape(v: str, limit: int = 30) -> str:
    """Privacy-preserving value pattern: digits->9, upper->A, lower->a."""
    out = []
    for ch in v[:limit]:
        if ch.isdigit():
            out.append("9")
        elif ch.isalpha():
            out.append("A" if ch.isupper() else "a")
        elif ch == " ":
            out.append("_")
        else:
            out.append(ch)
    return "".join(out) + ("..." if len(v) > limit else "")


def col_letter(idx: int) -> str:
    s, n = "", idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def q(v, limit: int = 120) -> str:
    if v is None:
        return "NULL"
    s = v if len(v) <= limit else v[:limit] + "..."
    return "'" + s + "'"


def display(v):
    if v is None:
        return NULL_DISPLAY
    if v == "":
        return EMPTY_DISPLAY
    return v


# --------------------------------------------------------------------------------------------
# Null policy
# --------------------------------------------------------------------------------------------
class NullPolicy:
    def __init__(self, cfg: dict | None):
        cfg = cfg or {}
        self.tokens = set(cfg.get("null_tokens", [""]))
        self.blank_is_null = bool(cfg.get("blank_is_null", True))
        self.case_insensitive = bool(cfg.get("tokens_case_insensitive", False))
        self._ci = {t.strip().casefold() for t in self.tokens}

    def is_null(self, v) -> bool:
        if v is None:
            return True
        if v in self.tokens:
            return True
        s = v.strip()
        if self.blank_is_null and s == "":
            return True
        if s != v and s in self.tokens:
            return True
        if self.case_insensitive and s.casefold() in self._ci:
            return True
        return False

    def norm(self, v):
        return None if self.is_null(v) else v

    def describe(self) -> str:
        toks = ", ".join(repr(t) for t in sorted(self.tokens))
        return (f"null tokens [{toks}]; blank/whitespace-only is null: {self.blank_is_null}; "
                f"token match case-insensitive: {self.case_insensitive}")


# --------------------------------------------------------------------------------------------
# File reading
# --------------------------------------------------------------------------------------------
def detect_encoding(path: Path) -> str:
    with open(path, "rb") as f:
        raw = f.read(1 << 20)
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError as exc:
        if len(raw) == (1 << 20) and exc.start >= len(raw) - 3:
            return "utf-8"   # multi-byte character cut by the sample boundary
    try:
        raw.decode("cp1252")
        return "cp1252"
    except UnicodeDecodeError:
        return "latin-1"


def read_text_sample(path: Path, encoding: str, max_lines: int = 200) -> list[str]:
    lines = []
    with open(path, encoding=encoding, errors="replace", newline="") as f:
        for i, line in enumerate(f):
            if i >= max_lines:
                break
            lines.append(line.rstrip("\r\n"))
    return lines


def detect_delimiter(lines: list[str]) -> str | None:
    lines = [ln for ln in lines if ln.strip()][:100]
    if not lines:
        return None
    best = None
    for d in CANDIDATE_DELIMITERS:
        if not any(d in ln for ln in lines):
            continue
        counts = []
        for ln in lines:
            try:
                counts.append(len(next(csv.reader([ln], delimiter=d))))
            except (csv.Error, StopIteration):
                counts.append(0)
        mode, freq = Counter(counts).most_common(1)[0]
        if mode < 2:
            continue
        score = (freq / len(counts), mode)
        if best is None or score > best[0]:
            best = (score, d)
    return best[1] if best else None


_NUMERIC_RE = re.compile(r"^[-+]?(\d+([.,]\d*)?|[.,]\d+)([eE][-+]?\d+)?$")
_DATEISH_RE = re.compile(r"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}")
_IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_ .#$%()/\-&]*$")


def guess_header(first: list[str]) -> tuple[bool, str]:
    cells = [c.strip() for c in first]
    while cells and cells[-1] == "":
        cells.pop()            # trailing delimiter on the header line
    if not cells or any(not c for c in cells):
        return False, "first record has empty cells"
    if len({c.casefold() for c in cells}) != len(cells):
        return False, "first record has duplicate values"
    if any(_NUMERIC_RE.match(c) or _DATEISH_RE.match(c) for c in cells):
        return False, "first record contains numeric/date-like values"
    ident = sum(1 for c in cells if _IDENT_RE.match(c)) / len(cells)
    if ident < 0.8:
        return False, "first record values do not look like column names"
    return True, "first record looks like column names (heuristic - confirm against STTM)"


def normalize_delimiter(d):
    if d is None:
        return None
    return DELIMITER_ALIASES.get(str(d).lower(), d) if len(str(d)) > 1 else d


def excel_value_to_str(v, number_format: str | None = None) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, _dt.datetime):
        if v.time() == _dt.time(0):
            return v.strftime("%Y-%m-%d")
        s = v.strftime("%Y-%m-%d %H:%M:%S")
        return s + (f".{v.microsecond:06d}" if v.microsecond else "")
    if isinstance(v, _dt.date):
        return v.isoformat()
    if isinstance(v, (_dt.time, _dt.timedelta)):
        return str(v)
    if isinstance(v, int):
        s = str(v)
    elif isinstance(v, float):
        if math.isfinite(v) and v.is_integer() and abs(v) < 1e15:
            s = str(int(v))
        else:
            s = fmt_decimal(Decimal(repr(v)).normalize()) if math.isfinite(v) else str(v)
    elif isinstance(v, Decimal):
        s = fmt_decimal(v)
    else:
        return str(v)
    # Excel cells stored as numbers but displayed with zero padding (number format "00000").
    if number_format and re.fullmatch(r"0+", number_format) and s.isdigit() and len(s) < len(number_format):
        s = s.zfill(len(number_format))
    return s


def excel_sheet_names(path: Path) -> list[str]:
    if path.suffix.lower() == ".xls":
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            return book.sheet_names()
        finally:
            book.release_resources()
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def iter_excel_rows(path: Path, sheet):
    """Yield (row_number, [text values]) for one worksheet, values rendered as text."""
    if path.suffix.lower() == ".xls":
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            sh = book.sheet_by_name(sheet) if isinstance(sheet, str) else book.sheet_by_index(sheet or 0)
            for r in range(sh.nrows):
                vals = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        vals.append(excel_value_to_str(xlrd.xldate_as_datetime(cell.value, book.datemode)))
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        vals.append("TRUE" if cell.value else "FALSE")
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        vals.append("")
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        vals.append(xlrd.error_text_from_code.get(cell.value, "#ERR"))
                    else:
                        vals.append(excel_value_to_str(cell.value))
                yield r + 1, vals
        finally:
            book.release_resources()
        return
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet is None:
            ws = wb.worksheets[0]
        elif isinstance(sheet, int):
            ws = wb.worksheets[sheet]
        else:
            if sheet not in wb.sheetnames:
                raise InputFileError(f"Worksheet '{sheet}' not found in {path.name}. Available: {wb.sheetnames}")
            ws = wb[sheet]
        ws.reset_dimensions()
        for rn, row in enumerate(ws.iter_rows(min_row=1), start=1):
            yield rn, [excel_value_to_str(c.value, getattr(c, "number_format", None)) for c in row]
    finally:
        wb.close()


class Dataset:
    def __init__(self, name: str, spec: dict, path: Path, columns: list[str], rows: list[list[str]],
                 line_nos: list[int], malformed: list[dict], meta: dict):
        self.name = name
        self.spec = spec
        self.path = path
        self.table = spec.get("table") or name
        self.columns = columns
        self.rows = rows
        self.line_nos = line_nos
        self.malformed = malformed
        self.meta = meta
        self._exact = {}
        self._ci: dict[str, list[int]] = {}
        for i, c in enumerate(columns):
            self._exact.setdefault(c, i)
            self._ci.setdefault(c.strip().casefold(), []).append(i)

    def col_index(self, name: str):
        if name in self._exact:
            return self._exact[name]
        hits = self._ci.get(str(name).strip().casefold(), [])
        return hits[0] if len(hits) == 1 else None

    @property
    def records_read(self) -> int:
        return self.meta.get("records_read", len(self.rows) + len(self.malformed))


def infer_format(path: Path, spec: dict) -> str:
    fmt = (spec.get("format") or "").lower()
    if fmt:
        if fmt not in ("delimited", "excel", "fixed", "json"):
            raise ConfigError(f"Unsupported format '{fmt}' (use delimited, excel, fixed or json)")
        return fmt
    if path.suffix.lower() in EXCEL_EXT:
        return "excel"
    if path.suffix.lower() in (".json", ".jsonl", ".ndjson"):
        return "json"
    # sniff text files whose content is JSON (e.g. a .txt holding a JSON array)
    try:
        with open(path, "rb") as f:
            head = f.read(4096).lstrip(b"\xef\xbb\xbf\xff\xfe \t\r\n\x00")
        if head[:1] in (b"[", b"{"):
            return "json"
    except OSError:
        pass
    return "delimited"


def _json_scalar(v) -> str:
    if v is None:
        return ""                      # JSON null -> empty (null under the default null policy)
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False, separators=(",", ":"))
    return str(v)                      # numbers arrive as their exact source text (parse_float/int=str)


def _flatten_json(obj, prefix, out, meta):
    for k, v in obj.items():
        key = f"{prefix}.{k}" if prefix else str(k)
        if isinstance(v, dict) and v:
            _flatten_json(v, key, out, meta)
        else:
            if isinstance(v, list):
                meta["json_arrays_serialized"] = meta.get("json_arrays_serialized", 0) + 1
            if v is None:
                meta["json_null_values"] = meta.get("json_null_values", 0) + 1
            out[key] = _json_scalar(v)


def _read_json_records(path: Path, spec: dict, meta: dict):
    """Read a JSON array / object-wrapped array / JSON Lines file into flat records.

    Nested objects are flattened to dotted column names (payment.amount). Numbers keep their exact
    source text (1500 vs 1500.00) because they are parsed as strings. Returns (columns, records)."""
    enc = spec.get("encoding") or detect_encoding(path)
    meta["encoding"] = enc
    text = path.read_text(encoding=enc)
    items, malformed = [], []
    try:
        data = json.loads(text, parse_float=str, parse_int=str)
        if isinstance(data, dict):
            jp = spec.get("json_path")
            if jp:
                for part in str(jp).split("."):
                    if not isinstance(data, dict) or part not in data:
                        raise InputFileError(f"json_path '{jp}' not found in {path.name}")
                    data = data[part]
            else:
                arrays = [k for k, v in data.items() if isinstance(v, list)]
                if len(arrays) == 1:
                    data = data[arrays[0]]
                    meta.setdefault("warnings", []).append(f"Records read from top-level key '{arrays[0]}'.")
                else:
                    data = [data]
        if not isinstance(data, list):
            raise InputFileError(f"{path.name}: JSON root must be an array of records (or set json_path)")
        meta["json_layout"] = "array"
        for i, item in enumerate(data, start=1):
            items.append((i, item))
    except json.JSONDecodeError as exc:
        # JSON Lines fallback: one object per line
        lines = text.splitlines()
        parsed_any = False
        for ln, line in enumerate(lines, start=1):
            if not line.strip():
                continue
            try:
                items.append((ln, json.loads(line, parse_float=str, parse_int=str)))
                parsed_any = True
            except json.JSONDecodeError:
                malformed.append({"line": ln, "expected": "JSON object", "found": "invalid JSON"})
        if not parsed_any:
            raise InputFileError(f"{path.name}: not valid JSON or JSON Lines ({exc})") from exc
        meta["json_layout"] = "json-lines"
    columns, records = [], []
    seen = set()
    flat_items = []
    for ref, item in items:
        if not isinstance(item, dict):
            malformed.append({"line": ref, "expected": "JSON object", "found": type(item).__name__})
            continue
        flat = {}
        _flatten_json(item, "", flat, meta)
        for k in flat:
            if k not in seen:
                seen.add(k)
                columns.append(k)
        flat_items.append((ref, flat))
    missing = 0
    for ref, flat in flat_items:
        missing += sum(1 for c in columns if c not in flat)
        records.append((ref, [flat.get(c, "") for c in columns]))
    if missing:
        meta.setdefault("warnings", []).append(
            f"{missing} field value(s) absent from some JSON records were treated as null.")
    if meta.get("json_arrays_serialized"):
        meta.setdefault("warnings", []).append(
            f"{meta['json_arrays_serialized']} nested JSON array value(s) kept as compact JSON text.")
    meta["line_reference"] = "JSON record number" if meta["json_layout"] == "array" else "line number"
    return columns, records, malformed


def _dedupe_columns(cols: list[str], meta: dict) -> list[str]:
    seen, out = Counter(), []
    for c in cols:
        seen[c] += 1
        if seen[c] > 1:
            new = f"{c}__{seen[c]}"
            meta.setdefault("duplicate_columns", []).append(f"{c} -> {new}")
            out.append(new)
        else:
            out.append(c)
    return out


def _finalize_records(name, spec, records, meta, header_only):
    """records: list of (line_no, fields). Applies header/columns, malformed detection."""
    header_cfg = spec.get("header", "auto")
    cfg_cols = spec.get("columns")
    data = records
    if header_cfg == "auto":
        if not records:
            has_header, why = False, "no records"
        else:
            has_header, why = guess_header(records[0][1])
        meta["header_detection"] = why
    else:
        has_header = bool(header_cfg)
        meta["header_detection"] = "explicitly configured"
    meta["header"] = has_header

    if has_header and records:
        header_fields = [h.strip() for h in records[0][1]]
        while header_fields and header_fields[-1] == "" and not cfg_cols:
            header_fields.pop()   # trailing delimiter / empty Excel cells after the last column
        data = records[1:]
        columns = list(cfg_cols) if cfg_cols else header_fields
    elif cfg_cols:
        columns = list(cfg_cols)
    else:
        width = Counter(len(f) for _, f in records[:1000]).most_common(1)[0][0] if records else 0
        columns = [f"COL_{i + 1}" for i in range(width)]
        meta.setdefault("warnings", []).append(
            "No header row and no column list configured - generic names COL_1..COL_n were assigned.")
    if any(c == "" for c in columns):
        columns = [c if c else f"UNNAMED_{i + 1}" for i, c in enumerate(columns)]
        meta.setdefault("warnings", []).append("Empty column names were replaced by UNNAMED_n.")
    columns = _dedupe_columns(columns, meta)
    n = len(columns)
    rows, line_nos, malformed = [], [], []
    trailing_trimmed = 0
    for ln, fields in data:
        if len(fields) > n and all(x == "" for x in fields[n:]):
            fields = fields[:n]
            trailing_trimmed += 1
        elif len(fields) < n and spec.get("_ragged_ok"):
            fields = fields + [""] * (n - len(fields))
        if len(fields) != n:
            malformed.append({"line": ln, "expected": n, "found": len(fields)})
            continue
        rows.append(fields)
        line_nos.append(ln)
    meta["trailing_empty_fields_trimmed"] = trailing_trimmed
    meta["records_read"] = len(rows) + len(malformed)
    meta["malformed_records"] = len(malformed)
    if header_only:
        meta["header_only"] = True
    return columns, rows, line_nos, malformed


def load_dataset(name: str, spec: dict, header_only: bool = False) -> Dataset:
    if "path" not in spec:
        raise ConfigError(f"Dataset '{name}' has no 'path'")
    path = resolve_path(spec["path"])
    if not path.is_file():
        raise InputFileError(f"Dataset '{name}': file not found: {path}")
    fmt = infer_format(path, spec)
    meta: dict = {"format": fmt, "file": path.name}
    limit = 50 if header_only else None
    skip = int(spec.get("skip_rows", 0) or 0)
    footer = int(spec.get("skip_footer", 0) or 0)
    records: list[tuple[int, list[str]]] = []
    blank = 0

    if fmt == "json":
        cols, recs, bad = _read_json_records(path, spec, meta)
        jspec = dict(spec, columns=cols, header=False)
        columns, rows, line_nos, malformed = _finalize_records(name, jspec, recs, meta, False)
        malformed = bad + malformed
        meta["records_read"] = len(rows) + len(malformed)
        meta["malformed_records"] = len(malformed)
        meta["blank_lines_skipped"] = 0
        meta["header"] = "n/a (JSON keys)"
        meta["header_detection"] = "JSON keys flattened to dotted column names"
        log.info("Loaded dataset %s (%s, JSON): %d columns, %d records read, %d malformed",
                 name, path.name, len(columns), meta["records_read"], len(malformed))
        return Dataset(name, spec, path, columns, rows, line_nos, malformed, meta)

    if fmt == "excel":
        sheet = spec.get("sheet")
        meta["sheet"] = sheet if sheet is not None else "(first sheet)"
        spec = dict(spec, _ragged_ok=True)
        for rn, vals in iter_excel_rows(path, sheet):
            if rn <= skip:
                continue
            while vals and vals[-1] == "":
                vals.pop()
            if not vals:
                blank += 1
                continue
            records.append((rn, vals))
            if limit and len(records) > limit:
                break
    else:
        auto_enc = not spec.get("encoding")
        encodings = [spec.get("encoding") or detect_encoding(path)]
        if auto_enc:
            encodings += [e for e in ("cp1252", "latin-1") if e not in encodings]
        last_exc = None
        for enc in encodings:
            try:
                records, blank, extra = _read_text_records(path, enc, spec, fmt, skip, limit, meta)
                meta["encoding"] = enc
                if enc != encodings[0]:
                    meta.setdefault("warnings", []).append(
                        f"Encoding '{encodings[0]}' failed to decode the file; fell back to '{enc}'.")
                break
            except UnicodeDecodeError as exc:
                last_exc = exc
                records = []
        else:
            raise InputFileError(f"Dataset '{name}': cannot decode {path.name} ({last_exc}). "
                                 f"Specify the correct 'encoding' in the config.")
    if footer and not header_only:
        meta["footer_records_skipped"] = footer
        records = records[:-footer] if footer < len(records) else []
    meta["blank_lines_skipped"] = blank
    meta["header_rows_skipped_by_config"] = skip
    columns, rows, line_nos, malformed = _finalize_records(name, spec, records, meta, header_only)
    log.info("Loaded dataset %s (%s): %d columns, %d records read, %d malformed%s",
             name, path.name, len(columns), meta["records_read"], len(malformed),
             " [header only]" if header_only else "")
    return Dataset(name, spec, path, columns, rows, line_nos, malformed, meta)


def _read_text_records(path, enc, spec, fmt, skip, limit, meta):
    records, blank = [], 0
    if fmt == "fixed":
        layout = spec.get("fixed_width")
        if not layout:
            raise ConfigError(f"Fixed-width dataset {path.name} requires a 'fixed_width' layout")
        cuts = []
        for fld in layout:
            start = int(fld["start"]) - 1
            end = start + int(fld["length"]) if "length" in fld else int(fld["end"])
            cuts.append((fld["name"], start, end))
        max_end = max(e for _, _, e in cuts)
        spec.setdefault("columns", [c for c, _, _ in cuts])
        if spec.get("header") in (None, "auto"):
            spec["header"] = False
        with open(path, encoding=enc, newline="") as f:
            for ln, line in enumerate(f, start=1):
                if ln <= skip:
                    continue
                line = line.rstrip("\r\n")
                if not line.strip():
                    blank += 1
                    continue
                fields = [line[s:e] for _, s, e in cuts]
                if len(line) > max_end and line[max_end:].strip():
                    fields.append(line[max_end:])   # -> flagged as malformed (extra data)
                records.append((ln, fields))
                if limit and len(records) > limit:
                    break
        return records, blank, None

    delimiter = normalize_delimiter(spec.get("delimiter"))
    if not delimiter:
        delimiter = detect_delimiter(read_text_sample(path, enc, 200 + skip)[skip:])
        if delimiter is None:
            delimiter = "\x1e"
            meta.setdefault("warnings", []).append(
                "No delimiter detected - file treated as a single column. Configure 'delimiter' "
                "or a 'fixed_width' layout if this is wrong.")
        meta["delimiter_detection"] = "auto"
    else:
        meta["delimiter_detection"] = "configured"
    meta["delimiter"] = delimiter
    quotechar = spec.get("quotechar", '"')
    with open(path, encoding=enc, newline="") as f:
        if len(delimiter) > 1:
            for ln, line in enumerate(f, start=1):
                if ln <= skip:
                    continue
                line = line.rstrip("\r\n")
                if not line.strip():
                    blank += 1
                    continue
                records.append((ln, line.split(delimiter)))
                if limit and len(records) > limit:
                    break
            return records, blank, None
        for _ in range(skip):
            if not f.readline():
                break
        if quotechar:
            reader = csv.reader(f, delimiter=delimiter, quotechar=quotechar, doublequote=True, strict=False)
        else:
            reader = csv.reader(f, delimiter=delimiter, quoting=csv.QUOTE_NONE)
        prev = 0
        long_spans = 0
        for fields in reader:
            start_line = prev + 1 + skip
            end_line = reader.line_num + skip
            prev = reader.line_num
            if not fields or (len(fields) == 1 and fields[0].strip() == ""):
                blank += 1
                continue
            if end_line - start_line > 50:
                long_spans += 1
                if long_spans <= 5:
                    meta.setdefault("warnings", []).append(
                        f"Record starting at line {start_line} spans {end_line - start_line + 1} lines - "
                        f"possible unbalanced quote character. Consider quotechar=null.")
            records.append((start_line, fields))
            if limit and len(records) > limit:
                break
    return records, blank, None


# --------------------------------------------------------------------------------------------
# Expression engine (STTM transformation rules encoded by the agent)
# --------------------------------------------------------------------------------------------
class Scope:
    """Column resolution for one evaluation context (source side or target side)."""

    def __init__(self, aliases: dict, primary: str, nulls: NullPolicy, engine: "Engine", side: str = "source"):
        self.aliases = aliases          # alias -> Dataset
        self.side = side
        self.primary = primary
        self.nulls = nulls
        self.engine = engine
        self.warnings: list[str] = []

    def resolve(self, ref: str):
        ref = str(ref)
        if "." in ref:
            a, c = ref.split(".", 1)
            if a in self.aliases:
                idx = self.aliases[a].col_index(c)
                return (a, idx) if idx is not None else None
        idx = self.aliases[self.primary].col_index(ref)
        if idx is not None:
            return self.primary, idx
        hits = [(a, ds.col_index(ref)) for a, ds in self.aliases.items()
                if a != self.primary and ds.col_index(ref) is not None]
        return hits[0] if len(hits) == 1 else None

    def label(self, ref: str):
        res = self.resolve(ref)
        if res is None:
            return None, str(ref)
        a, idx = res
        return self.aliases[a].table, self.aliases[a].columns[idx]


def collect_cols(node) -> list[str]:
    out = []
    if isinstance(node, dict):
        if "col" in node:
            out.append(str(node["col"]))
        for k, v in node.items():
            if k != "col":
                out.extend(collect_cols(v))
    elif isinstance(node, list):
        for v in node:
            out.extend(collect_cols(v))
    return list(dict.fromkeys(out))


def describe(node) -> str:
    if node is None:
        return "NULL"
    if not isinstance(node, dict):
        return json.dumps(node, ensure_ascii=False) if isinstance(node, str) else str(node)
    if "col" in node:
        return str(node["col"])
    if "lit" in node:
        return "NULL" if node["lit"] is None else json.dumps(lit_str(node["lit"]), ensure_ascii=False)
    if "op" not in node:   # plain parameter object (e.g. map values, lookup match entry)
        return "{" + ", ".join(f"{k}: {describe(v)}" for k, v in node.items()) + "}"
    op = str(node.get("op", "?"))
    parts = []
    for k, v in node.items():
        if k == "op":
            continue
        if k == "cases" and isinstance(v, list):
            parts.extend(f"WHEN {describe(c.get('when'))} THEN {describe(c.get('then'))}" for c in v)
        elif isinstance(v, dict):
            parts.append(f"{k}={describe(v)}" if k not in ("arg", "left") else describe(v))
        elif isinstance(v, list) and v and all(isinstance(x, dict) for x in v):
            parts.append(", ".join(describe(x) for x in v))
        else:
            parts.append(f"{k}={json.dumps(v, ensure_ascii=False)}")
    return f"{op.upper()}({'; '.join(parts)})"


def _arg(node, key, scope, required=True):
    if key not in node:
        if required:
            raise ConfigError(f"Operation '{node.get('op')}' requires '{key}': {json.dumps(node)[:200]}")
        return None
    return compile_expr(node[key], scope)


def _compile_col(ref, scope: Scope):
    res = scope.resolve(ref)
    if res is None:
        msg = f"Column '{ref}' was not found in the {scope.side} data"
        scope.warnings.append(msg)

        def missing(ctx):
            raise ReviewNeeded(msg, "SOURCE_COLUMN_MISSING")
        return missing
    alias, idx = res
    is_null = scope.nulls.is_null

    def f(ctx):
        row = ctx.get(alias)
        if row is None:
            return None
        if row is AMBIGUOUS:
            raise ReviewNeeded(f"Join '{alias}' returned more than one row for this record", "AMBIGUOUS_JOIN")
        v = row[idx]
        return None if is_null(v) else v
    return f


def _unary(func):
    def build(node, scope):
        a = _arg(node, "arg", scope)

        def f(ctx):
            v = a(ctx)
            return None if v is None else func(v, node)
        return f
    return build


def _b_substr(node, scope):
    a = _arg(node, "arg", scope)
    start = int(node.get("start", 1))
    length = node.get("length")
    length = None if length is None else int(length)

    def f(ctx):
        v = a(ctx)
        if v is None:
            return None
        i = start - 1 if start > 0 else (0 if start == 0 else max(len(v) + start, 0))
        return v[i:] if length is None else v[i:i + length]
    return f


def _b_pad(left):
    def build(node, scope):
        a = _arg(node, "arg", scope)
        n = int(node["length"])
        ch = str(node.get("char", " "))
        if len(ch) != 1:
            raise ConfigError("lpad/rpad 'char' must be a single character")
        trunc = bool(node.get("truncate", False))

        def f(ctx):
            v = a(ctx)
            if v is None:
                return None
            if len(v) > n:
                return v[:n] if trunc else v
            return v.rjust(n, ch) if left else v.ljust(n, ch)
        return f
    return build


def _b_concat(node, scope):
    args = [compile_expr(x, scope) for x in node.get("args", [])]
    if not args:
        raise ConfigError("concat requires 'args'")
    sep = str(node.get("sep", ""))
    mode = node.get("nulls", "empty")
    if mode not in ("empty", "skip", "propagate"):
        raise ConfigError("concat 'nulls' must be empty, skip or propagate")
    is_null = scope.nulls.is_null

    def f(ctx):
        vals = [a(ctx) for a in args]
        if mode == "propagate" and any(is_null(v) for v in vals):
            return None
        if mode == "skip":
            vals = [v for v in vals if not is_null(v)]
        return sep.join("" if v is None else v for v in vals)
    return f


def _b_date(node, scope):
    a = _arg(node, "arg", scope)
    fmts = node.get("in_format") or node.get("in_formats")
    if not fmts:
        raise ConfigError("date requires 'in_format' (Python strptime format or list of formats)")
    fmts = [fmts] if isinstance(fmts, str) else list(fmts)
    out = node.get("out_format")
    if not out:
        raise ConfigError("date requires 'out_format' (Python strftime format)")

    def f(ctx):
        v = a(ctx)
        if v is None:
            return None
        s = v.strip()
        for fm in fmts:
            try:
                return _dt.datetime.strptime(s, fm).strftime(out)
            except ValueError:
                continue
        raise ReviewNeeded(f"Source value {q(v)} cannot be parsed as a date using {fmts}",
                           "SOURCE_UNCONVERTIBLE", dq=True)
    return f


def _number_opts(node):
    rnd = str(node.get("rounding", "HALF_UP")).upper()
    if rnd not in ROUNDING:
        raise ConfigError(f"Unknown rounding '{rnd}'. Use one of {sorted(ROUNDING)}")
    scale = node.get("scale")
    return (None if scale is None else int(scale)), ROUNDING[rnd]


def _b_number(node, scope):
    a = _arg(node, "arg", scope)
    scale, rounding = _number_opts(node)
    thousands = node.get("thousands_sep")
    dec = node.get("decimal_sep", ".")
    strip = node.get("strip_chars", "")
    implied = node.get("implied_decimals")
    normalize = bool(node.get("strip_trailing_zeros", False))

    def f(ctx):
        v = a(ctx)
        if v is None:
            return None
        d = parse_decimal(v, thousands, dec, strip)
        if d is None:
            raise ReviewNeeded(f"Source value {q(v)} is not numeric", "SOURCE_UNCONVERTIBLE", dq=True)
        if implied:
            d = d.scaleb(-int(implied))
        if scale is not None:
            d = d.quantize(Decimal(1).scaleb(-scale), rounding=rounding)
        elif normalize:
            d = d.normalize()
        return fmt_decimal(d)
    return f


def _b_arith(node, scope):
    args = [compile_expr(x, scope) for x in node.get("args", [])]
    oper = node.get("operator")
    if oper not in ("+", "-", "*", "/") or len(args) < 2:
        raise ConfigError("arith requires operator (+,-,*,/) and at least two args")
    scale, rounding = _number_opts(node)

    def f(ctx):
        vals = [a(ctx) for a in args]
        if any(v is None for v in vals):
            return None
        nums = []
        for v in vals:
            d = parse_decimal(v)
            if d is None:
                raise ReviewNeeded(f"Value {q(v)} used in arithmetic is not numeric", "SOURCE_UNCONVERTIBLE", dq=True)
            nums.append(d)
        acc = nums[0]
        for d in nums[1:]:
            if oper == "+":
                acc += d
            elif oper == "-":
                acc -= d
            elif oper == "*":
                acc *= d
            else:
                if d == 0:
                    raise ReviewNeeded("Division by zero in derived value", "CANNOT_EVALUATE", dq=True)
                acc /= d
        if scale is not None:
            acc = acc.quantize(Decimal(1).scaleb(-scale), rounding=rounding)
        return fmt_decimal(acc)
    return f


def _b_coalesce(node, scope):
    args = [compile_expr(x, scope) for x in node.get("args", [])]
    is_null = scope.nulls.is_null

    def f(ctx):
        for a in args:
            v = a(ctx)
            if not is_null(v):
                return v
        return None
    return f


def _b_default(node, scope):
    a = _arg(node, "arg", scope)
    if "value" not in node:
        raise ConfigError("default requires 'value'")
    d = compile_expr(node["value"], scope)
    is_null = scope.nulls.is_null

    def f(ctx):
        v = a(ctx)
        return d(ctx) if is_null(v) else v
    return f


def _b_nullif(node, scope):
    a = _arg(node, "arg", scope)
    b = _arg(node, "value", scope)

    def f(ctx):
        v = a(ctx)
        return None if v is not None and v == b(ctx) else v
    return f


def _b_case(node, scope):
    cases = []
    for c in node.get("cases", []):
        if "when" not in c or "then" not in c:
            raise ConfigError("Each case entry needs 'when' and 'then'")
        cases.append((compile_cond(c["when"], scope), compile_expr(c["then"], scope)))
    if not cases:
        raise ConfigError("case requires a non-empty 'cases' list")
    other = compile_expr(node["else"], scope) if "else" in node else (lambda ctx: None)

    def f(ctx):
        for cond, then in cases:
            if cond(ctx):
                return then(ctx)
        return other(ctx)
    return f


def _b_map(node, scope):
    a = _arg(node, "arg", scope)
    values = node.get("values")
    if not isinstance(values, dict):
        raise ConfigError("map requires a 'values' object (source code -> target value)")
    ci = bool(node.get("case_insensitive", False))
    trim = bool(node.get("trim", False))

    def norm(k):
        k = k.strip() if trim else k
        return k.casefold() if ci else k
    table = {("__NULL__" if k == "__NULL__" else norm(k)): lit_str(v) for k, v in values.items()}
    has_default = "default" in node
    default = compile_expr(node["default"], scope) if has_default else None
    on_missing = node.get("on_missing", "review")
    if on_missing not in ("review", "null", "keep"):
        raise ConfigError("map 'on_missing' must be review, null or keep")

    def f(ctx):
        v = a(ctx)
        k = "__NULL__" if v is None else norm(v)
        if k in table:
            return table[k]
        if has_default:
            return default(ctx)
        if on_missing == "null":
            return None
        if on_missing == "keep":
            return v
        raise ReviewNeeded(f"Source value {q(v)} is not covered by the STTM code/value mapping",
                           "MAPPING_VALUE_NOT_COVERED")
    return f


def _b_lookup(node, scope: Scope):
    ds_name = node.get("dataset")
    if not ds_name:
        raise ConfigError("lookup requires 'dataset' (a reference dataset declared in 'datasets')")
    ref = scope.engine.get_dataset(ds_name)
    match = node.get("match") or []
    if not match:
        raise ConfigError("lookup requires 'match': [{'column': <ref column>, 'value': <expr>}]")
    key_idx, key_exprs = [], []
    problems = []
    for m in match:
        idx = ref.col_index(m["column"])
        if idx is None:
            problems.append(f"lookup column '{m['column']}' not in dataset '{ds_name}'")
        key_idx.append(idx)
        key_exprs.append(compile_expr(m["value"], scope))
    ret_idx = ref.col_index(node.get("return_column", ""))
    if ret_idx is None:
        problems.append(f"return_column '{node.get('return_column')}' not in dataset '{ds_name}'")
    trim = bool(node.get("trim_keys", False))
    ci = bool(node.get("case_insensitive", False))
    on_missing = node.get("on_missing", "review")
    on_multiple = node.get("on_multiple", "review")
    nulls = scope.nulls
    cache = {}

    def norm(v):
        v = v.strip() if trim else v
        return v.casefold() if ci else v

    if problems:
        msg = "; ".join(problems)
        scope.warnings.append(msg)

        def broken(ctx):
            raise ReviewNeeded(msg, "REFERENCE_DATA_MISSING")
        return broken

    def index():
        if "idx" not in cache:
            ix = {}
            for i, row in enumerate(ref.rows):
                vals = [row[j] for j in key_idx]
                if any(nulls.is_null(v) for v in vals):
                    continue
                ix.setdefault(tuple(norm(v) for v in vals), []).append(i)
            cache["idx"] = ix
        return cache["idx"]

    def f(ctx):
        vals = [e(ctx) for e in key_exprs]
        if any(v is None for v in vals):
            return None
        hits = index().get(tuple(norm(v) for v in vals), [])
        if not hits:
            if on_missing == "null":
                return None
            raise ReviewNeeded(f"No match in reference data '{ds_name}' for lookup key {vals}",
                               "LOOKUP_NOT_FOUND")
        if len(hits) > 1 and on_multiple != "first":
            raise ReviewNeeded(f"Lookup in '{ds_name}' returned {len(hits)} rows", "LOOKUP_AMBIGUOUS")
        return nulls.norm(ref.rows[hits[0]][ret_idx])
    return f


EXPR_OPS = {
    "trim": _unary(lambda v, n: v.strip(n.get("chars"))),
    "ltrim": _unary(lambda v, n: v.lstrip(n.get("chars"))),
    "rtrim": _unary(lambda v, n: v.rstrip(n.get("chars"))),
    "upper": _unary(lambda v, n: v.upper()),
    "lower": _unary(lambda v, n: v.lower()),
    "length": _unary(lambda v, n: str(len(v))),
    "left": _unary(lambda v, n: v[:int(n["length"])]),
    "right": _unary(lambda v, n: v[-int(n["length"]):] if int(n["length"]) > 0 else ""),
    "replace": _unary(lambda v, n: v.replace(str(n["old"]), str(n.get("new", "")))),
    "regex_replace": _unary(lambda v, n: re.sub(n["pattern"], n.get("replacement", ""), v)),
    "regex_extract": _unary(lambda v, n: (lambda m: m.group(int(n.get("group", 0))) if m else None)(
        re.search(n["pattern"], v))),
    "substr": _b_substr,
    "lpad": _b_pad(True),
    "rpad": _b_pad(False),
    "concat": _b_concat,
    "date": _b_date,
    "number": _b_number,
    "arith": _b_arith,
    "coalesce": _b_coalesce,
    "default": _b_default,
    "nullif": _b_nullif,
    "case": _b_case,
    "map": _b_map,
    "lookup": _b_lookup,
}


def compile_expr(node, scope: Scope):
    if not isinstance(node, dict):
        val = lit_str(node)        # bare JSON scalars are LITERALS, never column names
        return lambda ctx: val
    if "col" in node:
        return _compile_col(node["col"], scope)
    if "lit" in node:
        val = lit_str(node["lit"])
        return lambda ctx: val
    op = node.get("op")
    if op == "null":
        return lambda ctx: None
    builder = EXPR_OPS.get(op)
    if builder is None:
        raise ConfigError(f"Unknown expression op '{op}'. Supported: col, lit, null, {', '.join(sorted(EXPR_OPS))}")
    return builder(node, scope)


def _like_to_regex(p: str) -> str:
    out = []
    for ch in p:
        out.append(".*" if ch == "%" else "." if ch == "_" else re.escape(ch))
    return "".join(out)


def compile_cond(node, scope: Scope):
    if not isinstance(node, dict) or "op" not in node:
        raise ConfigError(f"Invalid condition: {json.dumps(node)[:200]}")
    op = node["op"]
    is_null = scope.nulls.is_null
    if op in ("true", "false"):
        val = op == "true"
        return lambda ctx: val
    if op in ("and", "or"):
        parts = [compile_cond(x, scope) for x in node.get("args", [])]
        if not parts:
            raise ConfigError(f"'{op}' requires args")
        return (lambda ctx: all(p(ctx) for p in parts)) if op == "and" else (lambda ctx: any(p(ctx) for p in parts))
    if op == "not":
        inner = compile_cond(node["arg"], scope)
        return lambda ctx: not inner(ctx)
    if op in ("is_null", "not_null"):
        a = _arg(node, "arg", scope)
        return (lambda ctx: is_null(a(ctx))) if op == "is_null" else (lambda ctx: not is_null(a(ctx)))
    numeric = bool(node.get("numeric", False))
    ci = bool(node.get("case_insensitive", False))
    trim = bool(node.get("trim", False))

    def prep(v):
        if numeric:
            d = parse_decimal(v)
            if d is None:
                raise ReviewNeeded(f"Value {q(v)} in a numeric condition is not numeric", "SOURCE_UNCONVERTIBLE", dq=True)
            return d
        v = v.strip() if trim else v
        return v.casefold() if ci else v

    if op in ("eq", "ne", "gt", "ge", "lt", "le"):
        left, right = _arg(node, "left", scope), _arg(node, "right", scope)
        fn = {"eq": lambda x, y: x == y, "ne": lambda x, y: x != y, "gt": lambda x, y: x > y,
              "ge": lambda x, y: x >= y, "lt": lambda x, y: x < y, "le": lambda x, y: x <= y}[op]

        def f(ctx):
            lv, rv = left(ctx), right(ctx)
            if is_null(lv) or is_null(rv):
                return False          # SQL three-valued logic: comparisons with NULL are not true
            return fn(prep(lv), prep(rv))
        return f
    if op in ("in", "not_in"):
        a = _arg(node, "arg", scope)
        raw = [lit_str(x) for x in node.get("values", [])]

        def f(ctx):
            v = a(ctx)
            if is_null(v):
                return False
            pv = prep(v)
            hit = any(pv == prep(x) for x in raw)
            return hit if op == "in" else not hit
        return f
    if op in ("like", "regex"):
        a = _arg(node, "arg", scope)
        pat = node["pattern"]
        rx = re.compile(_like_to_regex(pat) if op == "like" else pat, re.IGNORECASE if ci else 0)
        full = op == "like" or bool(node.get("fullmatch", False))

        def f(ctx):
            v = a(ctx)
            if is_null(v):
                return False
            return bool(rx.fullmatch(v) if full else rx.search(v))
        return f
    if op == "between":
        a, lo, hi = _arg(node, "arg", scope), _arg(node, "low", scope), _arg(node, "high", scope)

        def f(ctx):
            v, l, h = a(ctx), lo(ctx), hi(ctx)
            if is_null(v) or is_null(l) or is_null(h):
                return False
            return prep(l) <= prep(v) <= prep(h)
        return f
    raise ConfigError(f"Unknown condition op '{op}'")


# --------------------------------------------------------------------------------------------
# Comparison of expected vs actual target values
# --------------------------------------------------------------------------------------------
def diagnose(e: str, a: str) -> tuple[str, str]:
    if e.strip() == a.strip():
        return "values differ only by leading/trailing whitespace", "WHITESPACE"
    if e.casefold() == a.casefold():
        return "values differ only by letter case", "CASE"
    if e.isdigit() and a.isdigit() and e.lstrip("0") == a.lstrip("0"):
        return "values differ only by leading zeros", "LEADING_ZEROS"
    de, da = parse_decimal(e), parse_decimal(a)
    if de is not None and da is not None and de == da:
        return "values are numerically equal but formatted differently", "NUMERIC_FORMAT"
    if a and e.startswith(a) and len(a) < len(e):
        return f"actual value appears truncated (expected length {len(e)}, actual length {len(a)})", "TRUNCATION"
    return "", "VALUE_MISMATCH"


def compile_compare(spec: dict | None, nulls: NullPolicy):
    spec = spec or {}
    ctype = spec.get("type", "string")
    is_null = nulls.is_null

    def null_check(e, a):
        if e is None and a is None:
            return PASS, "", "MATCH"
        if e is None:
            return FAIL, f"Expected NULL but target has value {q(a)}", "NULL_EXPECTED"
        if a is None:
            return FAIL, f"Expected {q(e)} but target value is NULL/empty", "TARGET_NULL"
        return None

    if ctype == "string":
        trim = spec.get("trim", False)
        ci = bool(spec.get("case_insensitive", False))
        desc = "Exact string comparison" + (f", trim={trim}" if trim else "") + (", case-insensitive" if ci else "")

        def prep(v):
            if trim in (True, "both"):
                v = v.strip()
            elif trim == "right":
                v = v.rstrip()
            elif trim == "left":
                v = v.lstrip()
            return v.casefold() if ci else v

        def cmp(e, a):
            e = None if is_null(e) else e
            a = None if is_null(a) else a
            r = null_check(e, a)
            if r:
                return r
            if prep(e) == prep(a):
                return PASS, "", "MATCH"
            extra, cat = diagnose(e, a)
            return FAIL, f"Expected {q(e)} but found {q(a)}" + (f" - {extra}" if extra else ""), cat
        return cmp, desc

    if ctype == "numeric":
        tol = Decimal(str(spec.get("tolerance", "0")))
        thousands = spec.get("thousands_sep")
        desc = f"Numeric comparison (tolerance {tol})"

        def cmp(e, a):
            e = None if is_null(e) else e
            a = None if is_null(a) else a
            r = null_check(e, a)
            if r:
                return r
            de = parse_decimal(e, thousands)
            if de is None:
                return REVIEW, f"Expected value {q(e)} is not numeric - cannot compare numerically", "EXPECTED_NOT_NUMERIC"
            da = parse_decimal(a, thousands)
            if da is None:
                return FAIL, f"Actual target value {q(a)} is not numeric", "TARGET_NOT_NUMERIC"
            if abs(de - da) <= tol:
                return PASS, "", "MATCH"
            return FAIL, f"Numeric mismatch: expected {fmt_decimal(de)}, found {fmt_decimal(da)} (difference {fmt_decimal(da - de)})", "NUMERIC_MISMATCH"
        return cmp, desc

    if ctype == "date":
        ef = spec.get("expected_formats") or DEFAULT_DATE_FORMATS
        af = spec.get("actual_formats") or DEFAULT_DATE_FORMATS
        ef = [ef] if isinstance(ef, str) else ef
        af = [af] if isinstance(af, str) else af
        date_only = bool(spec.get("date_only", False))
        desc = f"Date comparison (actual formats {af}{', date part only' if date_only else ''})"

        def parse(v, fmts):
            for fm in fmts:
                try:
                    d = _dt.datetime.strptime(v.strip(), fm)
                    return d.date() if date_only else d
                except ValueError:
                    continue
            return None

        def cmp(e, a):
            e = None if is_null(e) else e
            a = None if is_null(a) else a
            r = null_check(e, a)
            if r:
                return r
            de = parse(e, ef)
            if de is None:
                return REVIEW, f"Expected value {q(e)} could not be parsed as a date", "EXPECTED_NOT_DATE"
            da = parse(a, af)
            if da is None:
                return FAIL, f"Actual target value {q(a)} is not a valid date in formats {af}", "TARGET_NOT_DATE"
            if de == da:
                return PASS, "", "MATCH"
            return FAIL, f"Date mismatch: expected {de}, found {da}", "DATE_MISMATCH"
        return cmp, desc
    raise ConfigError(f"Unknown compare type '{ctype}' (use string, numeric or date)")


# --------------------------------------------------------------------------------------------
# Engine
# --------------------------------------------------------------------------------------------
class Mapping:
    pass


class Engine:
    def __init__(self, config: dict, header_only: bool = False):
        self.config = config
        self.header_only = header_only
        self.nulls = NullPolicy(config.get("null_handling"))
        self.dataset_specs = config.get("datasets") or {}
        self._datasets: dict[str, Dataset] = {}
        if not isinstance(self.dataset_specs, dict) or not self.dataset_specs:
            raise ConfigError("Config must declare 'datasets' as an object of name -> file spec")
        if not config.get("validations"):
            raise ConfigError("Config must contain a non-empty 'validations' list")

    def get_dataset(self, name: str) -> Dataset:
        if name not in self.dataset_specs:
            raise ConfigError(f"Dataset '{name}' is not declared in 'datasets'")
        if name not in self._datasets:
            self._datasets[name] = load_dataset(name, dict(self.dataset_specs[name]), self.header_only)
        return self._datasets[name]

    def input_paths(self) -> list[Path]:
        paths = [resolve_path(s["path"]) for s in self.dataset_specs.values() if "path" in s]
        if self.config.get("sttm_file"):
            paths.append(resolve_path(self.config["sttm_file"]))
        return paths


def _key_expr(spec, scope):
    if isinstance(spec, str):
        return _compile_col(spec, scope), spec
    return compile_expr(spec, scope), describe(spec)


class ValidationSet:
    def __init__(self, engine: Engine, cfg: dict, position: int):
        self.engine = engine
        self.cfg = cfg
        self.name = str(cfg.get("name") or f"SET_{position}")
        nulls = engine.nulls
        src = cfg.get("source") or {}
        tgt = cfg.get("target") or {}
        if not src.get("dataset") or not tgt.get("dataset"):
            raise ConfigError(f"Validation '{self.name}': source.dataset and target.dataset are required")
        self.src_alias = src.get("alias", "S")
        self.tgt_alias = tgt.get("alias", "T")
        self.src_ds = engine.get_dataset(src["dataset"])
        self.tgt_ds = engine.get_dataset(tgt["dataset"])
        self.src_scope = Scope({self.src_alias: self.src_ds}, self.src_alias, nulls, engine)
        # joins
        self.joins = []
        for j in src.get("joins", []) or []:
            jds = engine.get_dataset(j["dataset"])
            alias = j.get("alias") or j["dataset"]
            jtype = j.get("type", "left").lower()
            if jtype not in ("left", "inner"):
                raise ConfigError(f"Join type '{jtype}' not supported (left, inner)")
            on = j.get("on") or []
            if not on:
                raise ConfigError(f"Join to '{j['dataset']}' requires 'on'")
            left_exprs = [_key_expr(o["left"], self.src_scope)[0] for o in on]
            right_idx = []
            for o in on:
                idx = jds.col_index(o["right"])
                if idx is None:
                    raise ConfigError(f"Join column '{o['right']}' not found in dataset '{j['dataset']}'")
                right_idx.append(idx)
            self.joins.append({"alias": alias, "ds": jds, "type": jtype, "left": left_exprs,
                               "right": right_idx, "cfg": j, "ambiguous": 0, "unmatched": 0})
            self.src_scope.aliases[alias] = jds
        self.src_filter = compile_cond(src["filter"], self.src_scope) if src.get("filter") else None
        self.src_filter_text = src.get("filter_text") or (describe(src["filter"]) if src.get("filter") else "")
        self.src_filter_ref = src.get("filter_sttm_ref", "")
        self.tgt_scope = Scope({self.tgt_alias: self.tgt_ds}, self.tgt_alias, nulls, engine, side="target")
        self.tgt_filter = compile_cond(tgt["filter"], self.tgt_scope) if tgt.get("filter") else None
        self.tgt_filter_text = tgt.get("filter_text") or (describe(tgt["filter"]) if tgt.get("filter") else "")
        self.tgt_filter_ref = tgt.get("filter_sttm_ref", "")
        # keys
        keys = cfg.get("keys") or []
        if not keys:
            raise ConfigError(f"Validation '{self.name}': at least one key is required")
        self.keys = []
        self.key_errors = []
        for k in keys:
            if "source" not in k or "target" not in k:
                raise ConfigError(f"Validation '{self.name}': each key needs 'source' and 'target'")
            for side, scope in (("source", self.src_scope), ("target", self.tgt_scope)):
                for ref in ([k[side]] if isinstance(k[side], str) else collect_cols(k[side])):
                    if scope.resolve(ref) is None:
                        self.key_errors.append(f"[{self.name}] Key column '{ref}' not found in {side} data "
                                               f"(available: {', '.join(scope.aliases[scope.primary].columns[:40])})")
            sf, sd = _key_expr(k["source"], self.src_scope)
            tf, td = _key_expr(k["target"], self.tgt_scope)
            self.keys.append({"name": k.get("name") or td, "src": sf, "tgt": tf, "src_desc": sd,
                              "tgt_desc": td, "sttm_ref": k.get("sttm_ref", ""), "notes": k.get("notes", "")})
        km = cfg.get("key_match") or {}
        self.key_trim = bool(km.get("trim", False))
        self.key_ci = bool(km.get("case_insensitive", False))
        # mappings
        self.mappings: list[Mapping] = []
        seen_ids = set()
        for i, m in enumerate(cfg.get("mappings") or [], start=1):
            mp = Mapping()
            mp.rule_id = str(m.get("rule_id") or f"{self.name}-M{i:03d}")
            if mp.rule_id in seen_ids:
                raise ConfigError(f"Duplicate rule_id '{mp.rule_id}' in validation '{self.name}'")
            seen_ids.add(mp.rule_id)
            mp.sttm_ref = m.get("sttm_ref", "")
            mp.rule_text = m.get("rule_text", "")
            mp.notes = m.get("notes", "")
            mp.target_column = m.get("target_column")
            if not mp.target_column:
                raise ConfigError(f"Mapping {mp.rule_id}: 'target_column' is required")
            mp.tgt_idx = self.tgt_ds.col_index(mp.target_column)
            mp.review = m.get("review")
            expr = m.get("expression")
            if expr is None and not mp.review:
                raise ConfigError(f"Mapping {mp.rule_id}: provide 'expression' or a 'review' reason")
            refs = m.get("source_columns") or (collect_cols(expr) if expr is not None else [])
            mp.source_refs = []
            tables = []
            for r in refs:
                res = self.src_scope.resolve(r)
                table, colname = self.src_scope.label(r)
                mp.source_refs.append((res, colname))
                if table and table not in tables:
                    tables.append(table)
            mp.source_table = m.get("source_table") or ", ".join(tables) or "(derived)"
            mp.source_column_label = ", ".join(c for _, c in mp.source_refs) or "(none)"
            before = len(self.src_scope.warnings)
            mp.expr = compile_expr(expr, self.src_scope) if expr is not None else None
            mp.warnings = self.src_scope.warnings[before:]
            mp.transform_desc = m.get("transformation") or (describe(expr) if expr is not None else "Not evaluated")
            mp.compare, mp.compare_desc = compile_compare(m.get("compare"), nulls)
            mp.counts = Counter()
            mp.categories = Counter()
            self.mappings.append(mp)
        self.unvalidated = cfg.get("unvalidated_rules") or []
        mapped = {self.tgt_ds.columns[mp.tgt_idx] for mp in self.mappings if mp.tgt_idx is not None}
        for k in keys:
            if isinstance(k["target"], str):
                idx = self.tgt_ds.col_index(k["target"])
                if idx is not None:
                    mapped.add(self.tgt_ds.columns[idx])
        self.unmapped_target_cols = [c for c in self.tgt_ds.columns if c not in mapped]

    # ---- configuration diagnostics -------------------------------------------------------
    def diagnostics(self) -> tuple[list[str], list[str]]:
        errors, warnings = list(self.key_errors), []
        for mp in self.mappings:
            if mp.tgt_idx is None:
                warnings.append(f"[{self.name}] {mp.rule_id}: target column '{mp.target_column}' not found in "
                                f"target dataset '{self.tgt_ds.name}' -> all its field validations will FAIL")
            for w in mp.warnings:
                warnings.append(f"[{self.name}] {mp.rule_id}: {w} -> REVIEW")
        for w in self.src_scope.warnings + self.tgt_scope.warnings:
            if not any(w in x for x in warnings) and not any(w.split("'")[1:2] and f"'{w.split(chr(39))[1]}'" in e for e in errors):
                warnings.append(f"[{self.name}] {w}")
        if self.unmapped_target_cols:
            warnings.append(f"[{self.name}] Target columns not covered by any mapping/key: "
                            f"{', '.join(self.unmapped_target_cols)}")
        return errors, warnings

    # ---- key normalisation ---------------------------------------------------------------
    def _norm_key(self, vals):
        out = []
        for v in vals:
            if v is None or self.engine.nulls.is_null(v):
                return None
            v = v.strip() if self.key_trim else v
            out.append(v.casefold() if self.key_ci else v)
        return tuple(out)

    def key_display(self, vals) -> str:
        if len(self.keys) == 1:
            return display(vals[0])
        return " | ".join(f"{k['name']}={display(v)}" for k, v in zip(self.keys, vals))

    # ---- phase 1: load, scope, key classification -----------------------------------------
    def classify(self):
        nulls = self.engine.nulls
        for j in self.joins:
            ix = {}
            for i, row in enumerate(j["ds"].rows):
                vals = [row[c] for c in j["right"]]
                if any(nulls.is_null(v) for v in vals):
                    continue
                ix.setdefault(tuple(vals), []).append(i)
            j["index"] = ix

        self.src_records = {}       # key -> list of (line, ctx, raw key vals)
        self.src_null = []          # (line, reason)
        self.src_excluded = []      # (line, normkey or None, raw key vals, reason)
        self.src_notes = {}         # line -> note
        self.dq_rows = []           # data-quality exceptions discovered while evaluating
        for row, line in zip(self.src_ds.rows, self.src_ds.line_nos):
            ctx = {self.src_alias: row}
            excluded_reason = None
            for j in self.joins:
                try:
                    lk = [e(ctx) for e in j["left"]]
                except ReviewNeeded:
                    lk = [None]
                hits = [] if any(v is None for v in lk) else j["index"].get(tuple(lk), [])
                if not hits:
                    j["unmatched"] += 1
                    ctx[j["alias"]] = None
                    if j["type"] == "inner":
                        excluded_reason = f"No matching row in inner join to '{j['ds'].name}'"
                elif len(hits) > 1:
                    j["ambiguous"] += 1
                    ctx[j["alias"]] = AMBIGUOUS
                else:
                    ctx[j["alias"]] = j["ds"].rows[hits[0]]
            if excluded_reason is None and self.src_filter is not None:
                try:
                    if not self.src_filter(ctx):
                        excluded_reason = f"Excluded by STTM source filter: {self.src_filter_text}"
                except ReviewNeeded as rn:
                    self.src_notes[line] = f"Source filter could not be evaluated ({rn.reason}); record kept in scope"
            try:
                raw = [k["src"](ctx) for k in self.keys]
                nk = self._norm_key(raw)
                key_err = None
            except ReviewNeeded as rn:
                raw, nk, key_err = [None] * len(self.keys), None, rn.reason
            if excluded_reason:
                self.src_excluded.append((line, nk, raw, excluded_reason))
                continue
            if nk is None:
                self.src_null.append((line, key_err or "Null/blank key value in source record"))
                continue
            self.src_records.setdefault(nk, []).append((line, ctx, raw))

        self.tgt_records = {}
        self.tgt_null = []
        self.tgt_excluded = 0
        self.tgt_filter_errors = 0
        for row, line in zip(self.tgt_ds.rows, self.tgt_ds.line_nos):
            ctx = {self.tgt_alias: row}
            if self.tgt_filter is not None:
                try:
                    if not self.tgt_filter(ctx):
                        self.tgt_excluded += 1
                        continue
                except ReviewNeeded:
                    self.tgt_filter_errors += 1
            try:
                raw = [k["tgt"](ctx) for k in self.keys]
                nk = self._norm_key(raw)
                key_err = None
            except ReviewNeeded as rn:
                raw, nk, key_err = [None] * len(self.keys), None, rn.reason
            if nk is None:
                self.tgt_null.append((line, key_err or "Null/blank key value in target record"))
                continue
            self.tgt_records.setdefault(nk, []).append((line, row, raw))

        self.matched, self.missing, self.extra, self.conflict = [], [], [], []
        for k, recs in self.src_records.items():
            t = self.tgt_records.get(k)
            if not t:
                self.missing.append(k)
            elif len(recs) == 1 and len(t) == 1:
                self.matched.append(k)
            else:
                self.conflict.append(k)
        for k in self.tgt_records:
            if k not in self.src_records:
                self.extra.append(k)
        self.excluded_by_key = {}
        for idx, (line, nk, raw, reason) in enumerate(self.src_excluded):
            if nk is not None:
                self.excluded_by_key.setdefault(nk, []).append(idx)
        log.info("[%s] matched=%d missing=%d extra=%d duplicate-conflict=%d src-null-key=%d tgt-null-key=%d "
                 "src-excluded=%d tgt-excluded=%d", self.name, len(self.matched), len(self.missing),
                 len(self.extra), len(self.conflict), len(self.src_null), len(self.tgt_null),
                 len(self.src_excluded), self.tgt_excluded)

    @property
    def field_rows_expected(self) -> int:
        return len(self.matched) * len(self.mappings)

    # ---- phase 2: field-level validation, streamed to the report --------------------------
    def evaluate(self, field_sheet, exceptions: list):
        self.field_stats = Counter()
        self.record_field_stats = {}
        tgt_table = self.tgt_ds.table
        for k in self.matched:
            line_s, ctx, raw = self.src_records[k][0]
            line_t, trow, _ = self.tgt_records[k][0]
            keyd = self.key_display(raw)
            rec = Counter()
            failed_cols = []
            for mp in self.mappings:
                src_val = self._source_value(mp, ctx)
                actual_raw = trow[mp.tgt_idx] if mp.tgt_idx is not None else None
                expected = None
                if mp.review:
                    status, reason, cat = REVIEW, f"STTM rule unclear / not evaluable: {mp.review}", "MAPPING_UNCLEAR"
                elif mp.tgt_idx is None:
                    status, reason, cat = FAIL, f"Target column '{mp.target_column}' not present in target file", "TARGET_COLUMN_MISSING"
                else:
                    try:
                        expected = mp.expr(ctx)
                        status, reason, cat = mp.compare(expected, actual_raw)
                    except ReviewNeeded as rn:
                        status, reason, cat = REVIEW, rn.reason, rn.category
                        if rn.dq:
                            self.dq_rows.append([self.name, "Data Quality", REVIEW, keyd, "Source",
                                                 self.src_ds.path.name, line_s, mp.rule_id,
                                                 mp.source_column_label, rn.reason])
                not_evaluated = mp.review or mp.tgt_idx is None or (status == REVIEW and expected is None)
                exp_disp = "(not evaluated)" if not_evaluated else display(expected)
                if mp.tgt_idx is None:
                    act_disp = "(column not in target)"
                else:
                    act_disp = display(actual_raw)
                field_sheet.write([self.name, keyd, mp.rule_id, mp.source_table, mp.source_column_label,
                                   tgt_table, mp.target_column, src_val, exp_disp, act_disp,
                                   mp.transform_desc, status, reason, mp.sttm_ref])
                mp.counts[status] += 1
                if status != PASS:
                    mp.categories[cat] += 1
                rec[status] += 1
                self.field_stats[status] += 1
                if status == FAIL:
                    failed_cols.append(mp.target_column)
                    exceptions.append([self.name, "Field Mismatch", FAIL, keyd, "Target", self.tgt_ds.path.name,
                                       line_t, mp.rule_id, mp.target_column, reason])
                elif status == REVIEW:
                    exceptions.append([self.name, "Review Item", REVIEW, keyd, "Target", self.tgt_ds.path.name,
                                       line_t, mp.rule_id, mp.target_column, reason])
            self.record_field_stats[k] = (rec, failed_cols)
            if failed_cols:
                exceptions.append([self.name, "Failed Record", FAIL, keyd, "Source/Target",
                                   f"{self.src_ds.path.name} / {self.tgt_ds.path.name}", f"{line_s} / {line_t}", "",
                                   ", ".join(failed_cols), f"Record has {len(failed_cols)} failed field(s)"])

    def _source_value(self, mp, ctx):
        if not mp.source_refs:
            return "(none)"
        parts = []
        for res, colname in mp.source_refs:
            if res is None:
                v = "(column not found)"
            else:
                row = ctx.get(res[0])
                v = NULL_DISPLAY if row is None else ("(ambiguous join)" if row is AMBIGUOUS else display(row[res[1]]))
            parts.append(v if len(mp.source_refs) == 1 else f"{colname}={v}")
        return " | ".join(parts)

    # ---- phase 3: record-level rows --------------------------------------------------------
    def record_rows(self, exceptions: list) -> list[list]:
        rows = []
        sname, tname = self.src_ds.path.name, self.tgt_ds.path.name
        lines = lambda recs: ", ".join(str(r[0]) for r in recs[:5]) + (" ..." if len(recs) > 5 else "")
        merged_excluded = set()
        nmap = len(self.mappings)
        for k in self.src_records:
            srecs = self.src_records[k]
            trecs = self.tgt_records.get(k, [])
            keyd = self.key_display(srecs[0][2])
            if k in self.record_field_stats:
                rec, failed = self.record_field_stats[k]
                if rec[FAIL]:
                    result, reason = FAIL, f"{rec[FAIL]} field(s) failed: {', '.join(failed[:10])}"
                elif rec[REVIEW]:
                    result, reason = REVIEW, f"{rec[REVIEW]} field(s) require review"
                elif nmap == 0:
                    result, reason = REVIEW, "No field mappings defined - only key presence verified"
                else:
                    result, reason = PASS, ""
                note = self.src_notes.get(srecs[0][0])
                if note:
                    reason = (reason + "; " if reason else "") + note
                    if result == PASS:
                        result = REVIEW
                rows.append([self.name, keyd, "Present", "Present", lines(srecs), lines(trecs),
                             sum(rec.values()), rec[PASS], rec[FAIL], rec[REVIEW], result, reason])
            elif not trecs:
                sstat = "Present" if len(srecs) == 1 else f"Duplicate ({len(srecs)} records)"
                reason = "Source record missing in target"
                if len(srecs) > 1:
                    reason += f"; key duplicated in source ({len(srecs)} records)"
                rows.append([self.name, keyd, sstat, "Missing", lines(srecs), "", 0, 0, 0, 0, FAIL, reason])
                exceptions.append([self.name, "Missing in Target", FAIL, keyd, "Source", sname, lines(srecs),
                                   "", "", reason])
            else:
                sstat = "Present" if len(srecs) == 1 else f"Duplicate ({len(srecs)} records)"
                tstat = "Present" if len(trecs) == 1 else f"Duplicate ({len(trecs)} records)"
                parts = []
                if len(srecs) > 1:
                    parts.append(f"duplicate key in source ({len(srecs)} records)")
                if len(trecs) > 1:
                    parts.append(f"duplicate key in target ({len(trecs)} records)")
                reason = "Duplicate key - " + "; ".join(parts) + "; field validation not performed (ambiguous match)"
                rows.append([self.name, keyd, sstat, tstat, lines(srecs), lines(trecs), 0, 0, 0, 0, FAIL, reason])
        for k in self.extra:
            trecs = self.tgt_records[k]
            keyd = self.key_display(trecs[0][2])
            tstat = "Present" if len(trecs) == 1 else f"Duplicate ({len(trecs)} records)"
            excl = self.excluded_by_key.get(k)
            if excl:
                merged_excluded.update(excl)
                sstat = "Present (excluded by STTM rule)"
                reason = f"Target contains a record that should have been excluded: {self.src_ds_excl_reason(excl[0])}"
                slines = ", ".join(str(self.src_excluded[i][0]) for i in excl[:5])
            else:
                sstat, slines = "Not in source", ""
                reason = "Extra record in target - no in-scope source record with this key"
            if len(trecs) > 1:
                reason += f"; key duplicated in target ({len(trecs)} records)"
            rows.append([self.name, keyd, sstat, tstat, slines, lines(trecs), 0, 0, 0, 0, FAIL, reason])
            exceptions.append([self.name, "Extra in Target", FAIL, keyd, "Target", tname, lines(trecs), "", "", reason])
        for line, reason in self.src_null:
            keyd = f"<NULL KEY> (source line {line})"
            rows.append([self.name, keyd, "Null key", "N/A", str(line), "", 0, 0, 0, 0, FAIL, reason])
            exceptions.append([self.name, "Null Key", FAIL, keyd, "Source", sname, line, "", "", reason])
        for line, reason in self.tgt_null:
            keyd = f"<NULL KEY> (target line {line})"
            rows.append([self.name, keyd, "N/A", "Null key", "", str(line), 0, 0, 0, 0, FAIL, reason])
            exceptions.append([self.name, "Null Key", FAIL, keyd, "Target", tname, line, "", "", reason])
        for idx, (line, nk, raw, reason) in enumerate(self.src_excluded):
            if idx in merged_excluded:
                continue
            keyd = self.key_display(raw) if nk is not None else f"<NULL KEY> (source line {line})"
            if nk is not None and nk in self.src_records and nk in self.tgt_records:
                res, why = REVIEW, f"{reason}; the same key also exists in in-scope source records"
                exceptions.append([self.name, "Review Item", REVIEW, keyd, "Source", sname, line, "FILTER", "", why])
            else:
                res, why = PASS, f"{reason}; correctly absent from target" if nk is not None else \
                    f"{reason}; key is null so target absence cannot be keyed"
            rows.append([self.name, keyd, "Excluded by STTM rule", "Not expected", str(line), "", 0, 0, 0, 0, res, why])
        for side, ds in (("Source", self.src_ds), ("Target", self.tgt_ds)):
            for m in ds.malformed:
                keyd = f"<MALFORMED> ({side.lower()} line {m['line']})"
                if isinstance(m["expected"], int):
                    reason = f"Malformed record: expected {m['expected']} fields, found {m['found']} - record could not be validated"
                else:
                    reason = f"Malformed record: expected {m['expected']}, found {m['found']} - record could not be validated"
                rows.append([self.name, keyd, "Malformed" if side == "Source" else "N/A",
                             "Malformed" if side == "Target" else "N/A",
                             str(m["line"]) if side == "Source" else "", str(m["line"]) if side == "Target" else "",
                             0, 0, 0, 0, FAIL, reason])
                exceptions.append([self.name, "Malformed Record", FAIL, keyd, side, ds.path.name, m["line"], "", "", reason])
        # duplicate key exceptions (one row per key per side)
        for side, recs_by_key, fname in (("Source", self.src_records, sname), ("Target", self.tgt_records, tname)):
            for k, recs in recs_by_key.items():
                if len(recs) > 1:
                    exceptions.append([self.name, "Duplicate Key", FAIL, self.key_display(recs[0][2]), side, fname,
                                       lines(recs), "", "", f"Key occurs {len(recs)} times in {side.lower()}"])
        self.record_result_counts = Counter(r[10] for r in rows)
        self.merged_excluded = len(merged_excluded)
        return rows

    def src_ds_excl_reason(self, idx):
        return self.src_excluded[idx][3]

    # ---- metrics ---------------------------------------------------------------------------
    def metrics(self) -> dict:
        s_dup_keys = sum(1 for r in self.src_records.values() if len(r) > 1)
        t_dup_keys = sum(1 for r in self.tgt_records.values() if len(r) > 1)
        s_dup_recs = sum(len(r) for r in self.src_records.values() if len(r) > 1)
        t_dup_recs = sum(len(r) for r in self.tgt_records.values() if len(r) > 1)
        src_in_scope = len(self.src_ds.rows) - len(self.src_excluded)
        tgt_in_scope = len(self.tgt_ds.rows) - self.tgt_excluded
        return {
            "source_files": ", ".join([self.src_ds.path.name] + [j["ds"].path.name for j in self.joins]),
            "target_file": self.tgt_ds.path.name,
            "source_table": self.src_ds.table,
            "target_table": self.tgt_ds.table,
            "source_records_read": self.src_ds.records_read,
            "source_malformed": len(self.src_ds.malformed),
            "source_excluded": len(self.src_excluded),
            "source_in_scope": src_in_scope,
            "source_null_key": len(self.src_null),
            "source_duplicate_keys": s_dup_keys,
            "source_duplicate_key_records": s_dup_recs,
            "target_records_read": self.tgt_ds.records_read,
            "target_malformed": len(self.tgt_ds.malformed),
            "target_excluded": self.tgt_excluded,
            "target_in_scope": tgt_in_scope,
            "target_null_key": len(self.tgt_null),
            "target_duplicate_keys": t_dup_keys,
            "target_duplicate_key_records": t_dup_recs,
            "matched": len(self.matched),
            "missing": len(self.missing),
            "missing_source_records": sum(len(self.src_records[k]) for k in self.missing),
            "extra": len(self.extra),
            "extra_target_records": sum(len(self.tgt_records[k]) for k in self.extra),
            "duplicate_conflicts": len(self.conflict),
            "conflict_source_records": sum(len(self.src_records[k]) for k in self.conflict),
            "conflict_target_records": sum(len(self.tgt_records[k]) for k in self.conflict),
            "mappings": len(self.mappings),
            "mappings_review": sum(1 for m in self.mappings if m.review),
            "unvalidated_rules": len(self.unvalidated),
            "field_total": sum(self.field_stats.values()),
            "field_pass": self.field_stats[PASS],
            "field_fail": self.field_stats[FAIL],
            "field_review": self.field_stats[REVIEW],
            "record_pass": self.record_result_counts[PASS],
            "record_fail": self.record_result_counts[FAIL],
            "record_review": self.record_result_counts[REVIEW],
        }

    def result(self, m: dict) -> tuple[str, str]:
        if m["record_fail"] or m["field_fail"]:
            return FAIL, "Definite mismatches or record-level failures exist"
        if m["record_review"] or m["field_review"] or m["unvalidated_rules"]:
            return REVIEW, "No definite failures, but some rules/records need review"
        if m["field_total"] == 0:
            return REVIEW, "No field validations were performed"
        return PASS, "All in-scope records and mapped fields matched"

    def reconciliation(self, m: dict, record_rows: int, written_fields: int) -> list[tuple]:
        checks = [
            ("Field PASS + FAIL + REVIEW = total field validations",
             m["field_total"], m["field_pass"] + m["field_fail"] + m["field_review"]),
            ("Total field validations = matched records x field mappings",
             m["matched"] * m["mappings"], m["field_total"]),
            ("Field Validation rows written = total field validations", m["field_total"], written_fields),
            ("Source records read = in scope + excluded + malformed",
             m["source_records_read"], m["source_in_scope"] + m["source_excluded"] + m["source_malformed"]),
            ("Source in scope = matched + missing + duplicate-conflict + null-key records",
             m["source_in_scope"], m["matched"] + m["missing_source_records"] + m["conflict_source_records"] + m["source_null_key"]),
            ("Target records read = in scope + excluded by target filter + malformed",
             m["target_records_read"], m["target_in_scope"] + m["target_excluded"] + m["target_malformed"]),
            ("Target in scope = matched + extra + duplicate-conflict + null-key records",
             m["target_in_scope"], m["matched"] + m["extra_target_records"] + m["conflict_target_records"] + m["target_null_key"]),
            ("Record Validation rows = keys + null-key + excluded + malformed rows",
             record_rows,
             m["matched"] + m["missing"] + m["extra"] + m["duplicate_conflicts"] + m["source_null_key"]
             + m["target_null_key"] + m["source_excluded"] - self.merged_excluded + m["source_malformed"]
             + m["target_malformed"]),
        ]
        return [(name, exp, act, "OK" if exp == act else "MISMATCH") for name, exp, act in checks]


# --------------------------------------------------------------------------------------------
# Excel report
# --------------------------------------------------------------------------------------------
FIELD_HEADERS = ["Validation Set", "Primary/Business Key", "Rule ID", "Source Table", "Source Column",
                 "Target Table", "Target Column", "Source Value", "Expected Target Value",
                 "Actual Target Value", "Transformation Rule", "Validation Status",
                 "Failure/Review Reason", "STTM Reference"]
RECORD_HEADERS = ["Validation Set", "Primary/Business Key", "Source Record Status", "Target Record Status",
                  "Source Line/Row", "Target Line/Row", "Fields Validated", "Field PASS", "Field FAIL",
                  "Field REVIEW", "Validation Result", "Failure Reason"]
EXCEPTION_HEADERS = ["Validation Set", "Exception Type", "Status", "Primary/Business Key", "Side", "File",
                     "Line/Row", "Rule ID", "Field", "Details"]
RULELOG_HEADERS = ["Validation Set", "Rule ID", "Rule Type", "STTM Reference", "STTM Rule", "Source Field",
                   "Target Field", "Transformation Applied", "Validation Method", "Records Evaluated",
                   "PASS", "FAIL", "REVIEW", "Result", "Notes"]


def clean_cell(s: str) -> str:
    # XlsxWriter escapes XML control characters itself (as _xHHHH_), so only the length limit is enforced here.
    if len(s) > EXCEL_MAX_CELL_CHARS:
        s = s[:EXCEL_MAX_CELL_CHARS - 40] + " ...[truncated: exceeds Excel cell limit]"
    return s


class Formats:
    def __init__(self, wb):
        self.header = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78",
                                     "border": 1, "text_wrap": True, "valign": "top"})
        self.cell = wb.add_format({"valign": "top"})
        self.num = wb.add_format({"valign": "top", "num_format": "#,##0"})
        self.title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"})
        self.section = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F4E78", "bottom": 2})
        self.label = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        self.value = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
        self.value_num = wb.add_format({"border": 1, "num_format": "#,##0"})
        self.pass_ = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        self.fail = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        self.review = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C5700"})
        self.info = wb.add_format({"bg_color": "#DDEBF7", "font_color": "#1F4E78"})
        self.big_pass = wb.add_format({"bold": True, "font_size": 14, "border": 1, "bg_color": "#C6EFCE", "font_color": "#006100"})
        self.big_fail = wb.add_format({"bold": True, "font_size": 14, "border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006"})
        self.big_review = wb.add_format({"bold": True, "font_size": 14, "border": 1, "bg_color": "#FFEB9C", "font_color": "#9C5700"})

    def status_rules(self, ws, first_row, last_row, col):
        if last_row < first_row:
            return
        for txt, fmt in ((PASS, self.pass_), (FAIL, self.fail), (REVIEW, self.review), (INFO, self.info),
                         ("OK", self.pass_), ("MISMATCH", self.fail)):
            ws.conditional_format(first_row, col, last_row, col,
                                  {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": fmt})


class SheetWriter:
    """Tabular sheet writer with automatic continuation sheets beyond Excel's row limit."""

    def __init__(self, wb, fmts: Formats, base: str, headers: list[str], expected_rows: int,
                 status_cols: list[int], max_widths: dict | None = None, rows_per_sheet: int = EXCEL_MAX_DATA_ROWS):
        self.wb, self.fmts, self.base, self.headers = wb, fmts, base, headers
        self.status_cols = status_cols
        self.rows_per_sheet = rows_per_sheet
        self.max_widths = max_widths or {}
        self.widths = [len(h) + 2 for h in headers]
        self.sheets: list[list] = []
        n_sheets = max(1, math.ceil(expected_rows / rows_per_sheet))
        for _ in range(n_sheets):
            self._add_sheet()
        self.current = 0
        self.total = 0

    def _add_sheet(self):
        name = self.base if not self.sheets else f"{self.base} ({len(self.sheets) + 1})"
        ws = self.wb.add_worksheet(name)
        for c, h in enumerate(self.headers):
            ws.write_string(0, c, h, self.fmts.header)
        ws.freeze_panes(1, 0)
        self.sheets.append([ws, 0])

    @property
    def sheet_names(self) -> list[str]:
        return [ws.get_name() for ws, _ in self.sheets]

    def write(self, values: list):
        if self.sheets[self.current][1] >= self.rows_per_sheet:
            self.current += 1
            if self.current >= len(self.sheets):
                self._add_sheet()
        ws, n = self.sheets[self.current]
        r = n + 1
        for c, v in enumerate(values):
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r, c, v, self.fmts.num)
                ln = len(f"{v:,}")
            else:
                s = clean_cell("" if v is None else str(v))
                ws.write_string(r, c, s, self.fmts.cell)
                ln = len(s)
            if ln > self.widths[c]:
                self.widths[c] = ln
        self.sheets[self.current][1] = n + 1
        self.total += 1

    def finalize(self):
        for ws, n in self.sheets:
            ws.autofilter(0, 0, max(n, 1), len(self.headers) - 1)
            for c, w in enumerate(self.widths):
                ws.set_column(c, c, min(max(w + 2, 10), self.max_widths.get(c, 50)))
            for c in self.status_cols:
                self.fmts.status_rules(ws, 1, n, c)


def _countifs(sheets: list[str], conditions: list[tuple[str, str]], cache: int) -> tuple[str, int]:
    parts = []
    for sh in sheets:
        sq = "'" + sh.replace("'", "''") + "'"
        args = ",".join(f"{sq}!${col}:${col},\"{val.replace(chr(34), chr(34) * 2)}\"" for col, val in conditions)
        parts.append(f"COUNTIFS({args})")
    return "=" + "+".join(parts), cache


def write_report(engine: Engine, sets: list[ValidationSet], report_path: Path, config_path: Path,
                 run_started: float) -> dict:
    import xlsxwriter
    from xlsxwriter.exceptions import FileCreateError

    wb = xlsxwriter.Workbook(str(report_path), {"constant_memory": True, "strings_to_numbers": False,
                                                "strings_to_formulas": False, "strings_to_urls": False})
    fmts = Formats(wb)
    summary_ws = wb.add_worksheet("Summary")

    # Record rows and exceptions depend on field results, but sheet ORDER must be fixed up-front.
    expected_record_rows = 0
    for vs in sets:
        expected_record_rows += (len(vs.src_records) + len(vs.extra) + len(vs.src_null) + len(vs.tgt_null)
                                 + len(vs.src_excluded) + len(vs.src_ds.malformed) + len(vs.tgt_ds.malformed))
    record_sheet = SheetWriter(wb, fmts, "Record Validation", RECORD_HEADERS, expected_record_rows, [10], {11: 80})
    field_sheet = SheetWriter(wb, fmts, "Field Validation", FIELD_HEADERS,
                              sum(vs.field_rows_expected for vs in sets), [11],
                              {7: 45, 8: 45, 9: 45, 10: 60, 12: 80})

    exceptions: list[list] = []
    for vs in sets:
        log.info("[%s] evaluating %d matched records x %d mappings", vs.name, len(vs.matched), len(vs.mappings))
        vs.evaluate(field_sheet, exceptions)
    all_record_rows = []
    for vs in sets:
        vs.record_rows_list = vs.record_rows(exceptions)
        all_record_rows.extend(vs.record_rows_list)
    for r in all_record_rows:
        record_sheet.write(r)

    # data-quality / review exceptions at dataset & rule level
    dq_seen = set()
    for vs in sets:
        exceptions.extend(vs.dq_rows)
        for ds in [vs.src_ds, vs.tgt_ds] + [j["ds"] for j in vs.joins]:
            if ds.name in dq_seen:
                continue
            dq_seen.add(ds.name)
            meta = ds.meta
            for w in meta.get("warnings", []):
                exceptions.append([vs.name, "Data Quality", REVIEW, "", "Dataset", ds.path.name, "", "", "", w])
            if meta.get("blank_lines_skipped"):
                exceptions.append([vs.name, "Data Quality", INFO, "", "Dataset", ds.path.name, "", "", "",
                                   f"{meta['blank_lines_skipped']} blank line(s) ignored"])
            if meta.get("trailing_empty_fields_trimmed"):
                exceptions.append([vs.name, "Data Quality", INFO, "", "Dataset", ds.path.name, "", "", "",
                                   f"{meta['trailing_empty_fields_trimmed']} record(s) had trailing empty fields (trailing delimiter) that were ignored"])
            for d in meta.get("duplicate_columns", []):
                exceptions.append([vs.name, "Data Quality", REVIEW, "", "Dataset", ds.path.name, "", "", "",
                                   f"Duplicate column name renamed: {d}"])
            if meta.get("footer_records_skipped"):
                exceptions.append([vs.name, "Data Quality", INFO, "", "Dataset", ds.path.name, "", "", "",
                                   f"{meta['footer_records_skipped']} trailer/footer record(s) excluded per config"])
            if ds is not vs.src_ds and ds is not vs.tgt_ds and ds.malformed:
                exceptions.append([vs.name, "Data Quality", REVIEW, "", "Dataset", ds.path.name, "", "", "",
                                   f"{len(ds.malformed)} malformed record(s) in joined/reference data"])
        for j in vs.joins:
            if j["ambiguous"]:
                exceptions.append([vs.name, "Data Quality", REVIEW, "", "Join", j["ds"].path.name, "", "JOIN", "",
                                   f"{j['ambiguous']} source record(s) matched more than one row in join to '{j['ds'].name}'"])
        if vs.tgt_filter_errors:
            exceptions.append([vs.name, "Data Quality", REVIEW, "", "Target", vs.tgt_ds.path.name, "", "FILTER", "",
                               f"Target filter could not be evaluated for {vs.tgt_filter_errors} record(s); kept in scope"])
        for line, note in vs.src_notes.items():
            exceptions.append([vs.name, "Review Item", REVIEW, "", "Source", vs.src_ds.path.name, line, "FILTER", "", note])
        for u in vs.unvalidated:
            exceptions.append([vs.name, "Unvalidated STTM Rule", REVIEW, "", "STTM", Path(str(engine.config.get("sttm_file", ""))).name,
                               u.get("sttm_ref", ""), u.get("rule_id", ""), u.get("target", ""),
                               u.get("reason", "Rule could not be validated with the provided files")])
    exc_order = {"Malformed Record": 0, "Null Key": 1, "Duplicate Key": 2, "Missing in Target": 3,
                 "Extra in Target": 4, "Failed Record": 5, "Field Mismatch": 6, "Data Quality": 7,
                 "Review Item": 8, "Unvalidated STTM Rule": 9}
    set_order = {vs.name: i for i, vs in enumerate(sets)}
    exceptions.sort(key=lambda r: (set_order.get(r[0], 99), exc_order.get(r[1], 99)))
    exc_sheet = SheetWriter(wb, fmts, "Exceptions", EXCEPTION_HEADERS, len(exceptions), [2], {9: 90})
    for r in exceptions:
        exc_sheet.write(r)

    # Rule log
    rule_rows = []
    for vs in sets:
        m_keys = vs.metrics()
        key_res = FAIL if (m_keys["source_null_key"] or m_keys["target_null_key"] or m_keys["duplicate_conflicts"]
                           or m_keys["source_duplicate_keys"] or m_keys["target_duplicate_keys"]) else PASS
        for k in vs.keys:
            rule_rows.append([vs.name, "KEY", "Primary/Business Key", k["sttm_ref"], k.get("notes", ""),
                              k["src_desc"], k["tgt_desc"],
                              f"Key match (trim={vs.key_trim}, case-insensitive={vs.key_ci})",
                              "Record matching on key", m_keys["source_in_scope"] + m_keys["target_in_scope"],
                              "", "", "", key_res,
                              f"Null keys: src {m_keys['source_null_key']}, tgt {m_keys['target_null_key']}; "
                              f"duplicate keys: src {m_keys['source_duplicate_keys']}, tgt {m_keys['target_duplicate_keys']}"])
        for j in vs.joins:
            jc = j["cfg"]
            rule_rows.append([vs.name, jc.get("rule_id", "JOIN"), f"Join ({j['type']})", jc.get("sttm_ref", ""),
                              jc.get("rule_text", ""), vs.src_ds.table, j["ds"].table,
                              describe({"op": "join", "on": jc.get("on")}), "Source enrichment",
                              len(vs.src_ds.rows), "", "", "", REVIEW if j["ambiguous"] else PASS,
                              f"Unmatched: {j['unmatched']}; ambiguous (multiple rows): {j['ambiguous']}"])
        if vs.src_filter is not None:
            leaked = vs.merged_excluded
            rule_rows.append([vs.name, "FILTER", "Source filter", vs.src_filter_ref, vs.src_filter_text,
                              vs.src_ds.table, vs.tgt_ds.table, describe(vs.cfg["source"]["filter"]),
                              "Records failing the filter must be absent from target", len(vs.src_ds.rows),
                              "", leaked, len(vs.src_notes), FAIL if leaked else (REVIEW if vs.src_notes else PASS),
                              f"{m_keys['source_excluded']} source record(s) excluded; {leaked} excluded record(s) found in target"])
        if vs.tgt_filter is not None:
            rule_rows.append([vs.name, "TARGET-FILTER", "Target scope filter", vs.tgt_filter_ref, vs.tgt_filter_text,
                              "", vs.tgt_ds.table, describe(vs.cfg["target"]["filter"]),
                              "Limits target records in scope", len(vs.tgt_ds.rows), "", "", vs.tgt_filter_errors,
                              REVIEW if vs.tgt_filter_errors else INFO,
                              f"{vs.tgt_excluded} target record(s) outside validation scope"])
        for mp in vs.mappings:
            evaluated = sum(mp.counts.values())
            if mp.counts[FAIL]:
                res = FAIL
            elif mp.counts[REVIEW] or evaluated == 0:
                res = REVIEW
            else:
                res = PASS
            notes = [mp.notes] if mp.notes else []
            if mp.review:
                notes.append(f"Marked REVIEW: {mp.review}")
            if evaluated == 0:
                notes.append("No matched records - rule not evaluated (not counted as PASS)")
            if mp.tgt_idx is None:
                notes.append("Target column missing in target file")
            notes.extend(mp.warnings)
            if mp.categories:
                notes.append("Issue categories: " + ", ".join(f"{k}={v}" for k, v in mp.categories.most_common()))
            rule_rows.append([vs.name, mp.rule_id, "Field mapping", mp.sttm_ref, mp.rule_text,
                              f"{mp.source_table}.{mp.source_column_label}", f"{vs.tgt_ds.table}.{mp.target_column}",
                              mp.transform_desc, mp.compare_desc, evaluated, mp.counts[PASS], mp.counts[FAIL],
                              mp.counts[REVIEW], res, "; ".join(notes)])
        for u in vs.unvalidated:
            rule_rows.append([vs.name, u.get("rule_id", ""), "Unvalidated STTM rule", u.get("sttm_ref", ""),
                              u.get("rule_text", ""), u.get("source", ""), u.get("target", ""), "Not applied",
                              "Not validated", 0, 0, 0, 0, REVIEW, u.get("reason", "")])
        if vs.unmapped_target_cols:
            rule_rows.append([vs.name, "UNMAPPED", "Coverage", "", "Target columns with no STTM mapping in config",
                              "", ", ".join(vs.unmapped_target_cols), "", "Coverage check", "", "", "", "", INFO,
                              "Confirm these columns are out of STTM scope (audit columns, surrogate keys, etc.)"])
    rule_sheet = SheetWriter(wb, fmts, "Rule Log", RULELOG_HEADERS, len(rule_rows), [13], {4: 60, 7: 60, 14: 80})
    for r in rule_rows:
        rule_sheet.write(r)

    # ---- metrics, reconciliation, overall result ----
    set_metrics, set_results, recon = [], [], []
    written_fields_by_set = Counter()
    for vs in sets:
        written_fields_by_set[vs.name] = sum(vs.field_stats.values())
    for vs in sets:
        m = vs.metrics()
        set_metrics.append(m)
        set_results.append(vs.result(m))
        recon.append(vs.reconciliation(m, len(vs.record_rows_list), written_fields_by_set[vs.name]))
    recon_ok = all(c[3] == "OK" for checks in recon for c in checks) and field_sheet.total == sum(
        m["field_total"] for m in set_metrics)
    order = {PASS: 0, REVIEW: 1, FAIL: 2}
    overall = max((r[0] for r in set_results), key=lambda s: order[s])
    if not recon_ok:
        log.error("Reconciliation check failed - see Summary sheet")
        if overall == PASS:
            overall = REVIEW

    _write_summary(summary_ws, fmts, engine, sets, set_metrics, set_results, recon, overall,
                   field_sheet.sheet_names, record_sheet.sheet_names, exc_sheet.sheet_names, len(exceptions),
                   config_path, run_started)
    for sw in (record_sheet, field_sheet, exc_sheet, rule_sheet):
        sw.finalize()
    summary_ws.activate()

    final_path = report_path
    try:
        wb.close()
    except FileCreateError:
        alt = report_path.with_name(f"{report_path.stem}_{time.strftime('%Y%m%d_%H%M%S')}{report_path.suffix}")
        log.warning("Cannot write %s (is it open in Excel?). Writing %s instead.", report_path.name, alt.name)
        wb.filename = str(alt)
        wb.close()
        final_path = alt

    exc_counts = Counter((r[1], r[2]) for r in exceptions)
    return {
        "report": str(final_path),
        "overall_result": overall,
        "reconciliation_ok": recon_ok,
        "sets": [
            {"name": vs.name, "result": res[0], "result_reason": res[1], "metrics": m,
             "reconciliation": [{"check": c[0], "expected": c[1], "actual": c[2], "status": c[3]} for c in rc],
             "rules": [{"rule_id": mp.rule_id, "target_column": mp.target_column, "sttm_ref": mp.sttm_ref,
                        "pass": mp.counts[PASS], "fail": mp.counts[FAIL], "review": mp.counts[REVIEW],
                        "issue_categories": dict(mp.categories)} for mp in vs.mappings],
             "unmapped_target_columns": vs.unmapped_target_cols}
            for vs, m, res, rc in zip(sets, set_metrics, set_results, recon)
        ],
        "exceptions_by_type": {f"{t} [{s}]": n for (t, s), n in sorted(exc_counts.items())},
        "sheets": {"record": record_sheet.sheet_names, "field": field_sheet.sheet_names,
                   "exceptions": exc_sheet.sheet_names, "rule_log": rule_sheet.sheet_names},
    }


SUMMARY_METRICS = [
    ("Source table", "source_table", False),
    ("Source file(s)", "source_files", False),
    ("Target table", "target_table", False),
    ("Target file", "target_file", False),
    ("Source records read", "source_records_read", True),
    ("Source malformed records", "source_malformed", True),
    ("Source records excluded by STTM filter/join", "source_excluded", True),
    ("Source records in scope", "source_in_scope", True),
    ("Source null-key records", "source_null_key", True),
    ("Source duplicate keys (distinct)", "source_duplicate_keys", True),
    ("Target records read", "target_records_read", True),
    ("Target malformed records", "target_malformed", True),
    ("Target records outside scope (target filter)", "target_excluded", True),
    ("Target records in scope", "target_in_scope", True),
    ("Target null-key records", "target_null_key", True),
    ("Target duplicate keys (distinct)", "target_duplicate_keys", True),
    ("Matched records", "matched", True),
    ("Missing records (source not in target)", "missing", True),
    ("Extra records (target not in source)", "extra", True),
    ("Duplicate-key conflicts (not field-validated)", "duplicate_conflicts", True),
    ("Field mappings validated", "mappings", True),
    ("Mappings marked REVIEW (unclear STTM rule)", "mappings_review", True),
    ("Unvalidated STTM rules (out of scope / missing data)", "unvalidated_rules", True),
]


def _write_summary(ws, fmts, engine, sets, metrics, results, recon, overall, field_sheets, record_sheets,
                   exc_sheets, n_exceptions, config_path, run_started):
    ws.set_column(0, 0, 52)
    ws.set_column(1, max(1, len(sets) + 1), 30)
    r = 0
    ws.write_string(r, 0, "STTM Source-to-Target Data Validation Report", fmts.title)
    r += 2
    ws.write_string(r, 0, "Run Information", fmts.section)
    r += 1
    sttm = engine.config.get("sttm_file")
    info = [
        ("STTM file", Path(sttm).name if sttm else "(not specified)"),
        ("Source file(s)", ", ".join(dict.fromkeys(m["source_files"] for m in metrics))),
        ("Target file(s)", ", ".join(dict.fromkeys(m["target_file"] for m in metrics))),
        ("Validation config", config_path.name),
        ("Report generated", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Run duration (seconds)", f"{time.time() - run_started:.1f}"),
        ("Engine version", __version__),
        ("Null handling", engine.nulls.describe()),
    ]
    assumptions = engine.config.get("assumptions") or []
    if isinstance(assumptions, str):
        assumptions = [assumptions]
    for i, a in enumerate(assumptions, start=1):
        info.append((f"Assumption {i}", str(a)))
    for label, val in info:
        ws.write_string(r, 0, label, fmts.label)
        ws.write_string(r, 1, clean_cell(val), fmts.value)
        r += 1
    ws.write_string(r, 0, "OVERALL VALIDATION RESULT", fmts.label)
    big = {PASS: fmts.big_pass, FAIL: fmts.big_fail, REVIEW: fmts.big_review}[overall]
    ws.write_string(r, 1, overall, big)
    r += 2

    multi = len(sets) > 1
    ws.write_string(r, 0, "Validation Statistics", fmts.section)
    r += 1
    ws.write_string(r, 0, "Metric", fmts.header)
    for i, vs in enumerate(sets):
        ws.write_string(r, 1 + i, vs.name, fmts.header)
    if multi:
        ws.write_string(r, 1 + len(sets), "TOTAL", fmts.header)
    r += 1
    for label, key, numeric in SUMMARY_METRICS:
        ws.write_string(r, 0, label, fmts.label)
        for i, m in enumerate(metrics):
            if numeric:
                ws.write_number(r, 1 + i, m[key], fmts.value_num)
            else:
                ws.write_string(r, 1 + i, clean_cell(str(m[key])), fmts.value)
        if multi and numeric:
            ws.write_number(r, 1 + len(sets), sum(m[key] for m in metrics), fmts.value_num)
        r += 1
    # live formulas over the detail sheets (cached values pre-computed) -> totals reconcile in Excel
    formula_rows = [
        ("Total field validations", field_sheets, "A", None, "field_total"),
        ("Field validations - PASS", field_sheets, "L", PASS, "field_pass"),
        ("Field validations - FAIL", field_sheets, "L", FAIL, "field_fail"),
        ("Field validations - REVIEW", field_sheets, "L", REVIEW, "field_review"),
        ("Records - PASS (Record Validation sheet)", record_sheets, "K", PASS, "record_pass"),
        ("Records - FAIL (Record Validation sheet)", record_sheets, "K", FAIL, "record_fail"),
        ("Records - REVIEW (Record Validation sheet)", record_sheets, "K", REVIEW, "record_review"),
    ]
    for label, sheets, col, status, key in formula_rows:
        ws.write_string(r, 0, label, fmts.label)
        for i, (vs, m) in enumerate(zip(sets, metrics)):
            conds = [("A", vs.name)] + ([(col, status)] if status else [])
            f, cache = _countifs(sheets, conds, m[key])
            ws.write_formula(r, 1 + i, f, fmts.value_num, cache)
        if multi:
            conds = [(col, status)] if status else [("A", "<>Validation Set")]
            if status:
                f, cache = _countifs(sheets, conds, sum(m[key] for m in metrics))
            else:
                f = "=" + "+".join(f"(COUNTA('{s}'!$A:$A)-1)" for s in sheets)
                cache = sum(m[key] for m in metrics)
            ws.write_formula(r, 1 + len(sets), f, fmts.value_num, cache)
        r += 1
    ws.write_string(r, 0, "Exceptions logged (Exceptions sheet)", fmts.label)
    ws.write_formula(r, 1, "=" + "+".join(f"(COUNTA('{s}'!$A:$A)-1)" for s in exc_sheets), fmts.value_num, n_exceptions)
    r += 1
    ws.write_string(r, 0, "Validation result", fmts.label)
    for i, (res, why) in enumerate(results):
        ws.write_string(r, 1 + i, res, {PASS: fmts.big_pass, FAIL: fmts.big_fail, REVIEW: fmts.big_review}[res])
    r += 1
    ws.write_string(r, 0, "Result explanation", fmts.label)
    for i, (res, why) in enumerate(results):
        ws.write_string(r, 1 + i, why, fmts.value)
    r += 2

    ws.write_string(r, 0, "Reconciliation Checks", fmts.section)
    r += 1
    for c, h in enumerate(["Check", "Validation Set", "Expected", "Actual", "Status"]):
        ws.write_string(r, c, h, fmts.header)
    r += 1
    first = r
    for vs, checks in zip(sets, recon):
        for name, exp, act, status in checks:
            ws.write_string(r, 0, name, fmts.value)
            ws.write_string(r, 1, vs.name, fmts.value)
            ws.write_number(r, 2, exp, fmts.value_num)
            ws.write_number(r, 3, act, fmts.value_num)
            ws.write_string(r, 4, status, fmts.value)
            r += 1
    fmts.status_rules(ws, first, r - 1, 4)
    r += 1
    ws.write_string(r, 0, "Status definitions", fmts.section)
    r += 1
    for s, d in ((PASS, "Expected and actual values match."),
                 (FAIL, "A definite mismatch or rule violation exists."),
                 (REVIEW, "Mapping unclear, reference data missing, or rule cannot be confidently evaluated. Never counted as PASS."),
                 (INFO, "Informational note (Exceptions/Rule Log only); does not affect totals.")):
        ws.write_string(r, 0, s, {PASS: fmts.pass_, FAIL: fmts.fail, REVIEW: fmts.review, INFO: fmts.info}[s])
        ws.write_string(r, 1, d, fmts.value)
        r += 1


# --------------------------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------------------------
ROLE_HINTS = {
    "sttm": ("sttm", "mapping", "s2t", "stm", "source_to_target", "sourcetotarget", "map"),
    "source": ("src", "source", "extract", "input", "legacy"),
    "target": ("tgt", "target", "dest", "output", "load", "dw", "warehouse"),
}


def cmd_discover(args) -> int:
    folder = resolve_path(args.input_dir) if args.input_dir else INPUT_DIR
    if not folder.is_dir():
        print(json.dumps({"error": f"Input folder not found: {folder}"}, indent=2))
        return 1
    files = []
    for p in sorted(folder.rglob("*")):
        if not p.is_file() or p.name.lower() == "readme.md" or p.name.startswith(("~$", ".")):
            continue
        ext = p.suffix.lower()
        entry = {"path": str(p.relative_to(PROJECT_ROOT)) if PROJECT_ROOT in p.parents else str(p),
                 "size_bytes": p.stat().st_size, "extension": ext,
                 "supported": ext in EXCEL_EXT or ext in TEXT_EXT}
        low = p.stem.lower()
        guesses = [role for role, hints in ROLE_HINTS.items() if any(h in low for h in hints)]
        if ext in EXCEL_EXT:
            try:
                entry["sheets"] = excel_sheet_names(p)
            except Exception as exc:  # noqa: BLE001 - report, do not crash discovery
                entry["sheet_error"] = str(exc)
        entry["role_guess"] = guesses or ["unknown"]
        files.append(entry)
    out = {"input_folder": str(folder), "file_count": len(files), "files": files,
           "note": "Role guesses are filename heuristics only - confirm STTM/source/target roles before validating."}
    print(json.dumps(out, indent=2))
    return 0


def cmd_inspect_sttm(args) -> int:
    path = resolve_path(args.file)
    if not path.is_file():
        print(f"ERROR: STTM file not found: {path}", file=sys.stderr)
        return 1
    out_path = ensure_safe_output(resolve_path(args.out) if args.out else work_dir() / "sttm_dump.md")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"# STTM dump: {path.name}", "",
             f"Generated {_dt.datetime.now():%Y-%m-%d %H:%M:%S}. Cell values are JSON-quoted so whitespace is visible.",
             "Reference format for config `sttm_ref`: `<Sheet>!R<row>` or `<Sheet>!<Col><row>`.", ""]
    overview = []
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            for i, ws in enumerate(wb.worksheets, start=1):
                state = ws.sheet_state
                lines.append(f"## Sheet {i}: \"{ws.title}\" ({state}, used range {ws.dimensions})")
                merged = [str(m) for m in ws.merged_cells.ranges]
                if merged:
                    lines.append(f"Merged ranges: {', '.join(merged[:200])}" + (" ..." if len(merged) > 200 else ""))
                nrows = 0
                comments = []
                for row in ws.iter_rows():
                    cells = []
                    for c in row:
                        if c.value is not None and str(c.value).strip() != "":
                            v = excel_value_to_str(c.value, c.number_format)
                            cells.append(f"{c.column_letter}={json.dumps(v, ensure_ascii=False)}")
                        if c.comment:
                            comments.append(f"{c.coordinate}: {json.dumps(c.comment.text, ensure_ascii=False)}")
                    if cells:
                        nrows += 1
                        lines.append(f"R{row[0].row}: " + " | ".join(cells))
                if comments:
                    lines.append("Cell comments:")
                    lines.extend(f"  {c}" for c in comments)
                lines.append("")
                overview.append({"sheet": ws.title, "state": state, "non_empty_rows": nrows,
                                 "dimensions": ws.dimensions, "merged_ranges": len(merged)})
        finally:
            wb.close()
    elif ext == ".xls":
        import xlrd
        names = excel_sheet_names(path)
        for i, name in enumerate(names, start=1):
            lines.append(f"## Sheet {i}: \"{name}\"")
            nrows = 0
            for rn, vals in iter_excel_rows(path, name):
                cells = [f"{col_letter(c)}={json.dumps(v, ensure_ascii=False)}" for c, v in enumerate(vals) if v.strip()]
                if cells:
                    nrows += 1
                    lines.append(f"R{rn}: " + " | ".join(cells))
            lines.append("")
            overview.append({"sheet": name, "non_empty_rows": nrows})
    elif ext in TEXT_EXT:
        ds = load_dataset("STTM", {"path": str(path), "header": False})
        lines.append(f"## Sheet 1: \"{path.stem}\" (delimited text, delimiter {ds.meta.get('delimiter')!r})")
        for ln, row in zip(ds.line_nos, ds.rows):
            cells = [f"{col_letter(c)}={json.dumps(v, ensure_ascii=False)}" for c, v in enumerate(row) if v.strip()]
            if cells:
                lines.append(f"R{ln}: " + " | ".join(cells))
        overview.append({"sheet": path.stem, "non_empty_rows": len(ds.rows), "malformed": len(ds.malformed)})
    else:
        print(f"ERROR: unsupported STTM format {ext}", file=sys.stderr)
        return 1
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(json.dumps({"sttm_file": str(path), "dump_file": str(out_path), "dump_lines": len(lines),
                      "sheets": overview}, indent=2))
    return 0


def cmd_profile(args) -> int:
    spec = {"path": args.file}
    for k in ("format", "encoding", "sheet", "quotechar"):
        v = getattr(args, k, None)
        if v is not None:
            spec[k] = v
    if args.delimiter:
        spec["delimiter"] = args.delimiter
    if args.no_quote:
        spec["quotechar"] = None
    if args.header in ("true", "false"):
        spec["header"] = args.header == "true"
    if args.skip_rows:
        spec["skip_rows"] = args.skip_rows
    if args.skip_footer:
        spec["skip_footer"] = args.skip_footer
    if args.sheet is not None and args.sheet.isdigit():
        spec["sheet"] = int(args.sheet)
    nulls = NullPolicy({"null_tokens": ["", "NULL", "null"]})
    ds = load_dataset("PROFILE", spec)
    cols = []
    for i, c in enumerate(ds.columns):
        vals = [r[i] for r in ds.rows]
        nonnull = [v for v in vals if not nulls.is_null(v)]
        distinct = len(set(vals))
        shapes = Counter(shape(v) for v in nonnull[:200000])
        entry = {
            "index": i + 1, "name": c, "non_null": len(nonnull), "null_or_blank": len(vals) - len(nonnull),
            "literal_NULL_tokens": sum(1 for v in vals if v.strip().upper() == "NULL"),
            "distinct": distinct, "min_len": min((len(v) for v in nonnull), default=0),
            "max_len": max((len(v) for v in nonnull), default=0),
            "leading_zero_values": sum(1 for v in nonnull if len(v) > 1 and v[0] == "0" and v[1].isdigit()),
            "leading_or_trailing_spaces": sum(1 for v in nonnull if v != v.strip()),
            "numeric_values": sum(1 for v in nonnull if parse_decimal(v) is not None),
            "top_patterns": [f"{p} (x{n})" for p, n in shapes.most_common(3)],
        }
        if args.show_values:
            seen = list(dict.fromkeys(nonnull))[: args.show_values]
            entry["sample_values"] = seen
        cols.append(entry)
    out = {"file": str(ds.path), "format": ds.meta.get("format"), "encoding": ds.meta.get("encoding"),
           "delimiter": ds.meta.get("delimiter"), "sheet": ds.meta.get("sheet"),
           "header_present": ds.meta.get("header"), "header_detection": ds.meta.get("header_detection"),
           "records_read": ds.records_read, "valid_records": len(ds.rows), "malformed_records": len(ds.malformed),
           "malformed_examples": ds.malformed[:10], "blank_lines": ds.meta.get("blank_lines_skipped"),
           "trailing_empty_fields_trimmed": ds.meta.get("trailing_empty_fields_trimmed"),
           "warnings": ds.meta.get("warnings", []), "duplicate_columns": ds.meta.get("duplicate_columns", []),
           "column_count": len(ds.columns), "columns": cols,
           "privacy_note": "Patterns mask digits as 9 and letters as A/a. Raw values only appear with --show-values."}
    if args.key:
        key_cols = [k.strip() for k in args.key.split(",")]
        idx = [ds.col_index(k) for k in key_cols]
        if None in idx:
            out["key_check"] = {"error": f"Key column(s) not found: {[k for k, i in zip(key_cols, idx) if i is None]}"}
        else:
            cnt = Counter(tuple(r[i] for i in idx) for r in ds.rows)
            out["key_check"] = {"key": key_cols,
                                "null_key_records": sum(n for k, n in cnt.items() if any(nulls.is_null(v) for v in k)),
                                "distinct_keys": len(cnt),
                                "duplicate_keys": sum(1 for n in cnt.values() if n > 1),
                                "duplicate_key_records": sum(n for n in cnt.values() if n > 1)}
    text = json.dumps(out, indent=2, ensure_ascii=False)
    try:
        dest = work_dir() / f"profile_{re.sub(r'[^A-Za-z0-9_.-]', '_', ds.path.stem)}.json"
        ensure_safe_output(dest).write_text(text, encoding="utf-8")
        out_note = f"(saved to {dest})"
    except OSError:
        out_note = ""
    print(text)
    if out_note:
        print(out_note, file=sys.stderr)
    return 0


def load_config(path: Path) -> dict:
    if not path.is_file():
        raise ConfigError(f"Config file not found: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Config is not valid JSON: {exc}") from exc


def build_sets(engine: Engine) -> list[ValidationSet]:
    sets, names = [], set()
    for i, v in enumerate(engine.config["validations"], start=1):
        vs = ValidationSet(engine, v, i)
        if vs.name in names:
            raise ConfigError(f"Duplicate validation name '{vs.name}'")
        names.add(vs.name)
        sets.append(vs)
    return sets


def cmd_check_config(args) -> int:
    cfg_path = resolve_path(args.config)
    result = {"config": str(cfg_path), "ok": False, "errors": [], "warnings": []}
    try:
        cfg = load_config(cfg_path)
        engine = Engine(cfg, header_only=True)
        if cfg.get("sttm_file") and not resolve_path(cfg["sttm_file"]).is_file():
            result["warnings"].append(f"sttm_file not found: {cfg['sttm_file']}")
        sets = build_sets(engine)
        for vs in sets:
            e, w = vs.diagnostics()
            result["errors"].extend(e)
            result["warnings"].extend(w)
        result["datasets"] = {n: {"file": d.path.name, "columns": d.columns, "header": d.meta.get("header"),
                                  "delimiter": d.meta.get("delimiter"), "encoding": d.meta.get("encoding"),
                                  "warnings": d.meta.get("warnings", [])}
                              for n, d in engine._datasets.items()}
        result["validations"] = [{"name": vs.name, "keys": len(vs.keys), "mappings": len(vs.mappings),
                                  "review_mappings": sum(1 for m in vs.mappings if m.review),
                                  "unvalidated_rules": len(vs.unvalidated)} for vs in sets]
        report = resolve_path(cfg.get("report_path", str(DEFAULT_REPORT)))
        ensure_safe_output(report, engine.input_paths())
        result["ok"] = not result["errors"]
    except (ConfigError, InputFileError, KeyError, ValueError, TypeError) as exc:
        result["errors"].append(f"{type(exc).__name__}: {exc}")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0 if result["ok"] else 1


def cmd_validate(args) -> int:
    started = time.time()
    cfg_path = resolve_path(args.config)
    cfg = load_config(cfg_path)
    engine = Engine(cfg)
    report = resolve_path(args.report or cfg.get("report_path") or str(DEFAULT_REPORT))
    if report.suffix.lower() != ".xlsx":
        raise ConfigError("Report path must end with .xlsx")
    ensure_safe_output(report, engine.input_paths())
    report.parent.mkdir(parents=True, exist_ok=True)
    log.info("Validation started with config %s", cfg_path.name)
    sets = build_sets(engine)
    errors = []
    for vs in sets:
        e, w = vs.diagnostics()
        errors.extend(e)
        for msg in w:
            log.warning(msg)
    if errors:
        raise ConfigError("Config has blocking errors (run check-config):\n  " + "\n  ".join(errors))
    for vs in sets:
        vs.classify()
    summary = write_report(engine, sets, report, cfg_path, started)
    summary["generated"] = _dt.datetime.now().isoformat(timespec="seconds")
    summary["config"] = str(cfg_path)
    summary["sttm_file"] = cfg.get("sttm_file")
    summary["duration_seconds"] = round(time.time() - started, 2)
    summary["privacy_note"] = "Aggregate statistics only. Record-level values are in the Excel report."
    dest = ensure_safe_output(work_dir() / "validation_summary.json")
    dest.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    log.info("Report written: %s | overall result: %s | reconciliation %s",
             summary["report"], summary["overall_result"], "OK" if summary["reconciliation_ok"] else "MISMATCH")
    print(json.dumps({"report": summary["report"], "summary_json": str(dest),
                      "overall_result": summary["overall_result"], "reconciliation_ok": summary["reconciliation_ok"],
                      "sets": [{"name": s["name"], "result": s["result"],
                                **{k: s["metrics"][k] for k in ("source_records_read", "target_records_read", "matched",
                                                                "missing", "extra", "field_total", "field_pass",
                                                                "field_fail", "field_review")}}
                               for s in summary["sets"]]}, indent=2))
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Internal STTM validation engine for the sttm-validator agent.")
    p.add_argument("--version", action="version", version=__version__)
    p.add_argument("--verbose", action="store_true")
    p.add_argument("--work-dir", help="Folder for logs/summary/dumps (default: output/work)")
    sub = p.add_subparsers(dest="command", required=True)

    d = sub.add_parser("discover", help="List input files and guess their roles")
    d.add_argument("--input-dir")

    s = sub.add_parser("inspect-sttm", help="Dump the STTM workbook to markdown for the agent to read")
    s.add_argument("file")
    s.add_argument("--out")

    pr = sub.add_parser("profile", help="Profile a data file (structure, quality, masked patterns)")
    pr.add_argument("file")
    pr.add_argument("--format", choices=["delimited", "excel", "fixed"])
    pr.add_argument("--delimiter")
    pr.add_argument("--encoding")
    pr.add_argument("--sheet")
    pr.add_argument("--quotechar")
    pr.add_argument("--no-quote", action="store_true", help="Disable quote handling")
    pr.add_argument("--header", choices=["auto", "true", "false"], default="auto")
    pr.add_argument("--skip-rows", type=int, default=0)
    pr.add_argument("--skip-footer", type=int, default=0)
    pr.add_argument("--key", help="Comma-separated key columns to check for null/duplicate keys")
    pr.add_argument("--show-values", type=int, default=0, help="Show up to N raw sample values per column")

    c = sub.add_parser("check-config", help="Validate a validation config without running it")
    c.add_argument("--config", default=str(WORK["dir"] / "validation_config.json"))

    v = sub.add_parser("validate", help="Run the validation and build the Excel report")
    v.add_argument("--config", default=str(WORK["dir"] / "validation_config.json"))
    v.add_argument("--report", help=f"Report path (default output/{REPORT_NAME})")

    args = p.parse_args(argv)
    if args.work_dir:
        WORK["dir"] = resolve_path(args.work_dir)
        ensure_safe_output(WORK["dir"])
    setup_logging(args.verbose)
    handlers = {"discover": cmd_discover, "inspect-sttm": cmd_inspect_sttm, "profile": cmd_profile,
                "check-config": cmd_check_config, "validate": cmd_validate}
    try:
        return handlers[args.command](args)
    except (ConfigError, InputFileError) as exc:
        log.error("%s: %s", type(exc).__name__, exc)
        return 1
    except ImportError as exc:
        log.error("Missing dependency (%s). Run: python -m pip install -r requirements.txt", exc)
        return 1
    except KeyboardInterrupt:
        log.error("Interrupted")
        return 130
    except Exception as exc:  # noqa: BLE001
        log.exception("Unexpected error: %s", exc)
        return 2


if __name__ == "__main__":
    sys.exit(main())
