# AGENTS.md

Single-window Tkinter desktop app (Chinese-language UI/data) that converts arbitrary Excel/PDF/image files into ECount accounting vouchers, plus a large 经营分析报告 Excel engine. All field names, sheet names, and config keys are **Chinese string literals that are load-bearing** — never translate or rename them casually.

**Read `CLAUDE.md` before making non-trivial changes.** It is the authoritative, detailed guide (module map, tab-to-module table, conversion pipeline, report engine, AI backend routing, known traps). This file only lists what an agent most often gets wrong.

## Commands

```bash
pip install -r requirements.txt           # pandas openpyxl pdfplumber zhipuai numpy xlrd
pip install -r requirements-optional.txt  # OCR/image extras (Pillow pytesseract paddleocr easyocr)

python 亿看智能识别系统.py                 # run the GUI (entry point)
python -m py_compile 亿看智能识别系统.py report_generator.py   # pre-commit sanity check
```

### Testing

- **Never run bare `pytest` at the repo root**: `test_ai_key.py` raises `SystemExit` at import time when `ZHIPU_API_KEY` is unset, which aborts collection with an INTERNALERROR.
- Newer tests are pytest-style; older ones are plain `python <file>` scripts with `__main__` blocks. Run individually:

```bash
python -m pytest test_report_generator_repairs.py            # report engine regressions (biggest suite)
python -m pytest test_report_generator_repairs.py -k 品类     # single test
python test_base_data.py                                     # legacy script-style test
```

- `test_template_path_resolution.py` executes the whole main module via `spec.loader.exec_module`, so it catches import-time breakage in `亿看智能识别系统.py` first.

## Must-know gotchas

- `亿看智能识别系统.py` (and likely `report_generator.py`) is stored **UTF-8 with BOM**. Preserve exactly one BOM; a duplicated BOM at byte 0 causes `SyntaxError: invalid non-printable character U+FEFF`.
- Nothing may assume the CWD. Use `resolve_resource_file()` / `resolve_resource_dir()` / `resolve_template_path()` (from the main module) for any new asset — a bare relative path breaks the PyInstaller build and launches from another directory. The template is found via a search list; `Template_通用凭证.xlsx` (repo root) is preferred, fallback is `基础数据/基础数据/Template.xlsx`; with no template the app falls back to a blank-workbook mode instead of failing.
- Optional modules (`report_generator`, `excel_merger`, `folder_processor`, `bank_parser`, image/DANFE GUIs) are imported inside `try/except ImportError` — **a broken import silently removes the tab instead of raising**. If a feature "disappeared", check those import guards at the top of `亿看智能识别系统.py`.
- Worker threads must talk to the UI via `_queue_report_event` / `_report_call_main` (queue drained by `_poll_report_task_queue`), never call Tkinter directly.
- Runtime state (`config.json`, `base_data.db` ~800 MB, `shipping.bd` [the `.bd` extension is intentional], `reconciliation.db`) lives in `APP_DIR` (next to the script). `runtime_paths.py` + `migrate_runtime_state.py` target `%APPDATA%\ecount-system`, but the app still reads `APP_DIR` — the migration is half-finished; don't assume `runtime_paths` is wired in.
- `.gitignore` is a whitelist (`*` ignore, then `!*.py`, `!*.md`, `!*/`). All `.xlsx`, `.db`, `.json` (incl. `config.json`), and templates are untracked working data; never copy/rewrite `base_data.db` wholesale. This working copy has no `.git` directory.
- `亿看智能识别系统/` (subdir with the same name as the repo) and `图片智能识别/` are **stale snapshot copies / raw data**, not live code — greps hit them constantly. Check the path before editing; never edit the snapshot as if it were the app. Root `*.py.rej` files are leftovers from failed patch applications.
- `亿看智能识别系统.spec` (PyInstaller) hardcodes its Analysis path to `C:\Users\123\Downloads\...` — fix before packaging.
- Base-data Excel layout: row 1 is a company name, row 2 is the header, and the last row is an export timestamp that must be dropped (`_clean_dataframe()` in `base_data_manager.py` handles this plus per-table column mapping).

## Env vars

- `ZHIPU_API_KEY` / `YIKAN_AI_API_KEY` — ZhipuAI key fallback so keys need not be stored in the DB.
- `ECOUNT_SYSTEM_HOME` / `YIKAN_APP_HOME` — override the runtime dir used by `runtime_paths.py` (limited effect; app still reads `APP_DIR`).

## Useful entry points beyond the GUI

- `validate_generated_report.py --report Generated_Report.xlsx --year 2026 --month 7 [--base-dir 基础资料] [--year-scope current|all]` — CLI validation of a generated report against source data (no GUI).
- Repo-local skill: `.agents/skills/monthly-report-annotation/` — procedure + script (`auto_annotate_report.py`) for annotating monthly report workbooks.
