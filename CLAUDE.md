# STTM Data Validation: project instructions

This project is a **Claude Code agent + skill** for STTM-driven source-to-target data validation.
It is not a web app or a standalone application.

## Components

- `.claude/agents/sttm-validator.md`: the `sttm-validator` subagent that orchestrates the validation.
- `.claude/skills/sttm-validation/SKILL.md`: the validation playbook (STTM analysis, file profiling,
  config format, expression/compare reference, statuses, reporting). It is preloaded into the agent.
- `scripts/validate_data.py`: internal deterministic engine used by the agent
  (`discover`, `inspect-sttm`, `profile`, `check-config`, `validate`). It is not a user interface.
- `input/`: user-supplied STTM, source and target files. **Read-only.**
- `output/`: generated `STTM_Validation_Report.xlsx`; `output/work/` holds the config, STTM dump,
  profiles, summary JSON and the log.

## Rules for any Claude session in this project

1. When the user asks for STTM / source-to-target validation, delegate to the `sttm-validator`
   agent, or follow the `sttm-validation` skill exactly if you work directly.
2. Never modify, move, re-save or delete anything in `input/`.
3. Never fabricate STTM rules, keys, code mappings, lookup data, expected values or sample business
   data. Unclear means REVIEW; skipped is never PASS.
4. Do not connect to databases or external services, or use credentials, unless the user explicitly
   configures and authorises it.
5. Keep record-level data out of chat and logs beyond minimal illustrative examples.
6. Python: use `.venv\Scripts\python.exe` (Windows) or `.venv/bin/python`. Dependencies are in
   `requirements.txt` (openpyxl, XlsxWriter, xlrd). Pure Python, no pandas.
7. Relative paths in the validation config resolve against the project root.
8. After changing `scripts/validate_data.py`, run
   `.venv\Scripts\python.exe -m py_compile scripts/validate_data.py` and a `check-config` on the
   current config if one exists.
