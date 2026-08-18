# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**亿看智能识别系统** (Yikan Intelligent Recognition System) — a single-window Tkinter desktop app that turns arbitrary business Excel/PDF/image files into ECount-compatible accounting artifacts, plus a large Excel report engine that builds monthly 经营分析报告 workbooks from ECount exports.

Everything is Chinese-language UI and data. Field names, sheet names, and config keys are Chinese string literals — they are load-bearing, not cosmetic.

## Commands

```bash
pip install -r requirements.txt              # pandas openpyxl pdfplumber zhipuai numpy xlrd
pip install -r requirements-optional.txt     # Pillow pytesseract openai paddleocr easyocr (image/OCR features)

python 亿看智能识别系统.py                    # run the GUI (entry point)
python migrate_runtime_state.py              # move config.json/*.db out of the repo into %APPDATA%\ecount-system

# Validate a generated report against 基础资料 source data (CLI, no GUI)
python validate_generated_report.py --report Generated_Report.xlsx --year 2026 --month 7 [--base-dir 基础资料] [--year-scope current|all]

pyinstaller 亿看智能识别系统.spec             # packaging (spec's Analysis path is hardcoded to a Downloads copy — fix before use)
```

### Tests

pytest 9 is installed and the newer tests are pytest-style; older ones are `python <file>` scripts with a `__main__` block.

```bash
# Never run bare `pytest` at the repo root: test_ai_key.py raises SystemExit at import
# time when ZHIPU_API_KEY is unset, which aborts collection with an INTERNALERROR.
python -m pytest test_report_generator_repairs.py          # report engine regressions (largest suite, ~3.9K lines)
python -m pytest test_report_inventory_sync.py test_danfe_recognition_module.py test_template_path_resolution.py
python -m pytest test_report_generator_repairs.py -k 品类   # single test

python test_base_data.py            # base data import/query
python test_smart_recognition.py    # summary recognition
python test_auto_balance.py         # debit/credit balancing
python test_shipping_module.py      # shipping DB + allocation
python test_recon_sign.py           # reconciliation sign/direction
```

Note: `test_template_path_resolution.py` executes the whole main module (`spec.loader.exec_module`), so it exercises import-time side effects of `亿看智能识别系统.py`. A syntax error or bad import there fails this test first.

### Sanity check before committing

`python -m py_compile 亿看智能识别系统.py report_generator.py` — the two big files are edited most often and a broken one only surfaces at app start. `亿看智能识别系统.py` is stored **UTF-8 with BOM**; preserve exactly one BOM (a duplicated BOM at byte 0 produced `SyntaxError: invalid non-printable character U+FEFF`).

**Environment variables**: `ZHIPU_API_KEY` / `YIKAN_AI_API_KEY` (ZhipuAI key, read as a fallback so keys need not be stored in the DB), `ECOUNT_SYSTEM_HOME` / `YIKAN_APP_HOME` (override the runtime dir used by `runtime_paths.py`).

## Repository conventions

`.gitignore` is a **whitelist**: `*` is ignored, then `!*.py`, `!*.md`, `!*/` re-included. Only Python and Markdown are tracked. Every `.xlsx`, `.db`, `.json` (including `config.json`), and template in the tree is untracked working data — adding one requires `git add -f`, and asking git for "the state of the data" is meaningless. `base_data.db` is ~800 MB locally; never copy or rewrite it wholesale.

`历史上传数据/亿看智能识别系统/亿看智能识别系统/` is a tracked snapshot of an older full copy of the app. Greps hit it constantly — always check the path before editing, and never edit it as if it were live code.

`*.py.rej` files at the root are leftovers from failed patch applications, not source.

## Architecture

### Module map

```
亿看智能识别系统.py     13.6K lines — GUI, all tabs, conversion pipeline (entry point)
report_generator.py     17.7K lines — 经营分析报告 Excel engine (largest module)
shipping_module.py       2.8K lines — 报关清单/柜子 costing, own SQLite (shipping.bd)
base_data_manager.py     2.2K lines — SQLite base data + app settings + caches
treeview_tools.py        1.4K lines — shared TreeView helpers, smart code restoration
export_format_manager.py 1.3K lines — user-defined output column mappings (config.json)
local_llm_analyzer.py    1.3K lines — LLM commentary written back into report workbooks
image_recognition_gui.py 1.1K + image_intelligence.py 1.0K — image → table extraction
summary_intelligence.py  1.0K lines — summary-text field recognition
reconciliation_module.py 966 lines  — StandardReconciler (local ledger vs Yikan)
danfe_recognition_module.py / danfe_xml_parser.py / danfe_recognition_gui.py — Brazilian DANFE invoices
bank_parser.py, excel_merger.py, folder_processor.py, shipping_report_utils.py, runtime_paths.py
```

### The main window is one class

`ExcelConverterGUI` spans lines ~675–13606 of `亿看智能识别系统.py` — every tab, dialog, and worker lives on it. Tabs are added in `_build_ui()` and each has its own `_build_*_tab()`:

| Tab | Builder | Backing module |
|---|---|---|
| Excel凭证转换 | `_build_excel_converter_tab` | template mapping + `summary_intelligence` |
| 摘要匹配 | `_build_summary_match_tab` | in-file (fuzzy match summaries between two files) |
| 智能对账系统 | `_build_reconciliation_tab` | `reconciliation_module`, `bank_parser` |
| 基础数据管理 | `_build_base_data_tab` | `base_data_manager` |
| 报关清单汇总 | `_build_shipping_tab` | `shipping_module` |
| 经营报告 | `_build_report_tab` | `report_generator`, `local_llm_analyzer` |
| 文档识别 (Docs) | `_build_document_recognition_tab` | `danfe_*` |
| 文件夹平铺汇总 / 批量合并 | `_build_folder_processor_tab` / `_build_batch_merge_tab` | `folder_processor`, `excel_merger` |
| 控制台 (Console) | `_build_console_tab` | stdout capture |

Optional modules (`report_generator`, `excel_merger`, `folder_processor`, `bank_parser`, image and DANFE GUIs) are imported in `try/except ImportError` blocks. **A broken import silently removes its tab instead of raising** — if a feature "disappeared", check the import guard at the top of the file first.

Long-running work (report generation, batch runs) is dispatched via `_start_report_background_task()` onto a thread that talks back through a queue drained by `_poll_report_task_queue()`; UI calls from workers must go through `_queue_report_event` / `_report_call_main`, never directly.

### Path resolution

Nothing assumes the CWD. `RESOURCE_ROOT_DIRS` (frozen-exe dir, its parents, `APP_DIR`, `os.getcwd()`) and `RESOURCE_SEARCH_DIRS` (each root + `基础数据/基础数据`, `基础资料`, `基础数据`) drive `resolve_resource_file()` / `resolve_resource_dir()` / `resolve_template_path()`. Use those helpers for any new asset; a bare relative path will break both the PyInstaller build and launches from another directory.

Consequence: `Template.xlsx` is **not** in the repo root — it lives in `基础数据/基础数据/Template.xlsx`. On startup the app prefers `Template_通用凭证.xlsx` (repo root) if present. If no template resolves, the app offers a no-template mode (blank workbook) rather than failing.

### Persistence: three stores, two locations

| Store | What | Where |
|---|---|---|
| `config.json` | `FIELD_RULES`, `FIELD_SYNONYMS`, `HEADER_SCHEMES`, `EXPORT_FORMATS`, `REPORT_SHEET_OUTPUT_FORMATS` | `APP_DIR` (next to the script), auto-backed up to `config.json.bak` |
| `base_data.db` | 7 base tables + `app_config`, `smart_recognition_cache`, `mapping_schemes`, `auto_mapping_cache`, `recognition_rules`, `custom_category`/`custom_record`, `import_log` | `APP_DIR` |
| `shipping.bd` (note the `.bd` extension), `reconciliation.db`, `reconciliation_header_mapping.json` | module-local state | `APP_DIR` |

`runtime_paths.py` defines the intended future home (`%APPDATA%\ecount-system`, or `ECOUNT_SYSTEM_HOME`) and `migrate_runtime_state.py` moves files there, **but the app itself still reads `APP_DIR`** (`CONFIG_FILE = os.path.join(APP_DIR, "config.json")`). Migration is half-finished; don't assume `runtime_paths` is authoritative without wiring it up.

User toggles are persisted per key in `app_config` under `setting_*` names (e.g. `setting_enable_smart_recognition`), loaded in `ExcelConverterGUI.__init__`.

### Base data (`base_data_manager.py`)

`BaseDataManager` owns 7 base tables — `currency`, `department`, `warehouse`, `account_subject`, `product`, `business_partner` (has `local_code`, the bridge for reconciliation), `bank_account` — with `query()`, `search_by_name()`, `lookup_value()`, and add/update/delete.

Source Excel layout: **row 1 is the company name, row 2 is the header (`header=1`), and the last row is an export timestamp that must be dropped**. `_clean_dataframe()` holds the per-table Excel-column → DB-column mapping; new data types need an entry there plus a table in `_init_database()`.

### Conversion pipeline (Excel凭证转换)

`FIELD_RULES` / `FIELD_SYNONYMS` (defaults in the main file, overridable from `config.json`) describe each template column. `normalize_header()` + `score_similarity()` (exact 1.0 → synonym 0.9 → containment 0.85 → difflib ≤0.8, accepted at ≥0.6) auto-map source columns; `convert_value()` coerces per `{"type": "date"|"number"|"text"}` with `max_int_len` / `max_decimal_len` / `max_len`.

Modes: `MODE_GENERAL_VOUCHER` 通用凭证 / `MODE_SALES_OUTBOUND` 销售出库 / `MODE_CUSTOM` / `MODE_ORIGINAL` (no template).

**Field value priority** (do not reorder without checking the preview dialog): manual mapping → recognized dedicated column (`_recognize_from_fields`: 日期/金额/汇率) → summary recognition → configured default.

`summary_intelligence.py` resolves a summary string into 业务类型 / 往来单位 / 科目 / 部门 / 金额 / 日期, using keyword rules from `_init_recognition_rules()` plus a cached snapshot of base data; results are cached in `smart_recognition_cache` (whose `match_items` JSON column stores learned aliases — the "将摘要加入科目匹配项" action writes here).

### Report engine (`report_generator.py`)

`ReportGenerator(base_data_dir)` reads a **flat directory of monthly ECount exports** — by default `基础资料/`, whose filenames look like `利润表_2026-07-01_to_2026-07-31.xlsx`, `销售出库明细表_…`, `实际成本报表_…`, `资产负债表_…`, `会计科目明细表_{应收账款|应付账款|银行存款}_…`, `科目账簿_期间费用…`.

- `_classify_source_file()` buckets each file into `profit / cost / expense / asset / sales / ar / ap / cash` **using both the filename and a peek at the first rows and sheet names** — 科目账簿 expense files are only identifiable by their 6601/6602/6603 account codes inside the sheet. Month keys are `YYYY-MM` via `_determine_period_key()`.
- `self.data[category][month_key] = DataFrame` is the single in-memory shape everything downstream reads.
- `list_available_months(ready_only=True)` intersects `profit`/`cost`/`asset` — a month missing any core statement is not offered in the GUI.
- `generate_report(template_path, output_path, target_year, target_month, year_scope, replenishment_params, cashflow_params, include_ai_placeholders, fail_on_validation_error, fail_on_data_quality_error, allow_generated_report_template, output_sheets)` loads the template workbook, rewrites sheets by **Chinese sheet name** (`利润表`, `费用明细`, `经营指标`, `仪表盘`, `按产品汇总(含合计数)`, …), regenerates charts, and appends 审计日志 / 数据质量检查 sheets. `generate_batch_reports` / `generate_continuous_batch_reports` loop months.
- `year_scope` is `current` | `recent_two_years` | `all` and is normalized from both English keys and Chinese phrases ("跨年", "历年") by `_normalize_year_scope()`.
- Guards worth knowing: `_detect_template_risk()` refuses a template that looks like a previously *generated* report (name contains 经营分析报告, or an 审计日志 sheet with generation records) unless `allow_generated_report_template=True`; a missing template silently switches to blank-workbook mode; `fail_on_data_quality_error` blocks output when `_run_data_quality_checks()` found ERRORs.
- `validate_report_file()` re-checks a produced workbook against the source data — that's what `validate_generated_report.py` wraps.

`local_llm_analyzer.LocalLLMAnalyzer` writes AI commentary into the generated workbook (chunked sheet summaries, optional chart recognition) via ZhipuAI or any OpenAI-compatible local endpoint.

### AI backend routing

AI config is **not** a single provider setting. `app_config` holds:

- `ai_backends` — JSON list of `{name, provider ("zhipu"|"lm_studio"), api_key, base_url, model}` (edited in 设置 → AI 设置 → "模型源配置")
- `ai_task_map` — JSON `{task_id: backend_name}` for the five tasks `smart_summary`, `formula_gen`, `image_rec`, `reconciliation`, `report_analysis` ("功能模块分配" tab)

`_get_ai_backend_for_task()` → `_normalize_ai_backend()` → `_build_ai_context()` resolves a task to a client, falling back to the legacy single-provider keys (`ai_provider`, `ai_api_key`, `ai_base_url`, `ai_model_name`) and then to the `ZHIPU_API_KEY`/`YIKAN_AI_API_KEY` env vars. Defaults: `glm-4-flash` (zhipu), `local-model` @ `http://localhost:1234/v1` (LM Studio); image recognition uses `glm-4v-flash`.

### Export format layer (`export_format_manager.py`)

Per-module user-defined output schemas stored in `config.json` under `EXPORT_FORMATS[module_key] = {"active": name, "use_original": bool, "formats": {name: [ {output, source, default}, … ]}}`. Module keys in use: `main_export`, `summary_match_export`, `image_recognition`, `shipping_product`, `shipping_container` (report sheets use the separate `REPORT_SHEET_OUTPUT_FORMATS` key). `apply_export_format(module_key, headers, rows, base_data_mgr)` returns `(headers, rows, applied)`.

Two special `source` tokens resolve against base data instead of a source column:
- `BD:table|target_col|key_header` — per-row lookup keyed by `code` taken from the row's `key_header` column
- `BDV:table|target_col|key_col|key_val` — fixed lookup, same value for every row

## Known limitations / traps

1. Whole source files are read into memory; no pagination.
2. Numbers over `max_int_len` are truncated silently.
3. First matching keyword wins in `summary_intelligence` business-type rules.
4. One sheet per conversion run.
5. `shipping_module` defaults to `shipping.bd`, not `shipping.db` — the odd extension is intentional and matches the on-disk file.
6. `folder_processor` drag-and-drop uses raw Win32 `ctypes` window subclassing; it is Windows-only and unwinds the window proc on close.
