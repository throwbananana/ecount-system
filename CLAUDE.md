# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**亿看智能识别系统** (Yikan Intelligent Recognition System) is a Tkinter GUI application for converting arbitrary Excel files into standardized accounting voucher template format ("一般凭证") with intelligent column mapping, AI-powered field recognition, and integrated base data management.

**Core Capabilities**:
1. Excel format conversion with intelligent column matching and field transformation rules
2. AI-powered summary recognition (摘要智能识别) - extracts business type, partner info, account codes, amounts, dates from summary text
3. Multi-field recognition from dedicated columns (日期, 金额, 汇率)
4. Base data management with CRUD operations for 7 data types via SQLite
5. Image intelligent recognition (图片智能识别) - OCR and AI-powered table extraction from images with template-based export
6. Accounting reconciliation between local and Yikan formats
7. Shipping/logistics cost management with multi-currency allocation
8. Business intelligence report generation

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Optional: For AI/OCR features
pip install -r requirements-optional.txt

# Run the main application
python 亿看智能识别系统.py

# Run tests
python test_base_data.py              # Base data import/query tests
python test_smart_recognition.py      # Summary recognition tests
python test_multi_field_recognition.py # Multi-field recognition tests
python test_preview_function.py       # Preview window tests
python test_edit_functions.py         # CRUD operation tests
python test_image_recognition.py      # Image recognition tests
python test_auto_balance.py           # Debit/credit balance tests
python test_actual_conversion.py      # Full conversion workflow tests
python test_encoding_fix.py           # Character encoding tests
python test_actual_matching.py        # Column matching tests
python test_recon_sign.py             # Reconciliation sign/direction tests
```

**Environment Variables** (optional):
- `ZHIPU_API_KEY` or `YIKAN_AI_API_KEY` - API key for ZhipuAI cloud service

## Architecture

### Core Module Design

```
亿看智能识别系统.py          # Main GUI application (entry point)
├── base_data_manager.py     # SQLite database manager (7 base tables + system tables)
├── summary_intelligence.py  # AI-powered text recognition engine
├── image_intelligence.py    # Image OCR and AI vision recognition
├── image_recognition_gui.py # Standalone image recognition GUI window
├── export_format_manager.py # Flexible output format mapping system
├── reconciliation_module.py # Standard format accounting reconciliation
├── shipping_module.py       # Shipping cost management (shipping.db)
├── bank_parser.py           # PDF bank statement extraction
├── report_generator.py      # Business intelligence report generation
└── treeview_tools.py        # GUI utilities for TreeView (smart code restoration)
```

### Operating Modes

The main application supports multiple conversion modes:
- `MODE_GENERAL_VOUCHER` = "通用凭证模式" - Standard accounting vouchers
- `MODE_SALES_OUTBOUND` = "销售出库模式" - Sales outbound/export format
- `MODE_CUSTOM` = "自定义模式" - User-defined templates
- `MODE_ORIGINAL` = "原格式模式(不使用模板)" - No template conversion

### 1. Main Application (`亿看智能识别系统.py`)

**Key Components**:
- `FIELD_RULES` (dict): Type and constraints for template columns - `"date"`, `"number"`, `"text"`
- `FIELD_SYNONYMS` (dict): Synonym mappings for auto-matching source columns
- `TemplateHeader` (dataclass): Stores template column name, position, comment
- `load_template_headers()`: Extracts headers from `Template.xlsx` row 1
- `normalize_header()`: Normalizes column names for matching
- `score_similarity()`: Multi-strategy scoring (exact=1.0, contains=0.85, synonym=0.9, difflib)
- `convert_value()`: Applies transformation based on field type

**GUI Structure**:
- Menu bar with "基础数据" and "文件" menus for base data and file operations
- Top section: File selection + sheet picker + mode selector
- Middle section: Scrollable mapping grid (template → source dropdowns)
- Bottom section: Auto-match button, Smart recognition checkbox, Default values button, Convert button

### 2. Base Data Manager (`base_data_manager.py`)

**`BaseDataManager` class**:
- `_init_database()`: Creates 7 tables + import_log table
- `import_all_data()` / `import_single_file()`: Batch/single Excel import
- `_clean_dataframe()`: Excel column → database column mapping with table-specific logic
- `query(table, code)`: Retrieve by code
- `search_by_name(table, keyword)`: Full-text search
- `add_record()` / `update_record()` / `delete_record()`: CRUD operations
- `get_table_columns()`: Returns column names for table

**Database Tables** (Base Data):
| Table | Primary Fields | Records |
|-------|---------------|---------|
| currency | code, name | 1 |
| department | code, name | 2 |
| warehouse | code, name | 3 |
| account_subject | code, name | 254 |
| product | code, name, category, etc. | 1674 |
| business_partner | code, name, type, local_code | 480 |
| bank_account | code, name, bank_name | 21 |

**System Tables**:
- `smart_recognition_cache`: Caches AI recognition results with `match_items` (JSON aliases)
- `app_config`: Persists user settings
- `mapping_schemes`: Stores column mapping presets
- `auto_mapping_cache`: Caches auto-match results per template/mode
- `recognition_rules`: Custom business type, account, department rules
- `import_log`: Tracks all data imports with timestamps

**Excel Import Format**: Header at row 2 (index=1), row 1 is company name, last row is timestamp (filtered out).

### 3. Summary Intelligence (`summary_intelligence.py`)

**`SummaryIntelligence` class**:
- `_init_ai_client()`: Initializes ZhipuAI or OpenAI (LM Studio) client based on config
- `_init_recognition_rules()`: Loads business type keywords and rules
- `_load_base_data_cache()`: Caches partners, accounts, departments from database
- `recognize(summary)`: Main entry - returns dict of recognized fields
- `recognize_from_row(row_data)`: Multi-field recognition from entire row
- `batch_recognize(df)`: Process entire DataFrame

**Recognition Pipeline**:
```
Input row → _recognize_from_fields() → Extract date/amount/rate from columns
         → recognize(summary)       → Extract from summary text:
             ├── _recognize_business_type()  # Keywords → type, account, summary_code
             ├── _recognize_partner()        # Database match + company patterns
             ├── _recognize_account()        # Direct reference + name match
             ├── _recognize_department()     # Database + keywords
             ├── _recognize_amount()         # Regex patterns
             └── _recognize_date()           # Multiple date formats
```

**AI Configuration** (in `default_values` dict):
- `ai_provider`: "zhipu" (cloud) or "lm_studio" (local)
- `ai_api_key`: API key for ZhipuAI
- `ai_base_url`: LM Studio endpoint (default: "http://localhost:1234/v1")
- `ai_model_name`: Local model name

### 4. Image Intelligence (`image_intelligence.py`)

**`ImageIntelligence` class**:
- `check_and_install_dependencies()`: Auto-installs missing dependencies (Pillow, zhipuai, openai)
- `recognize_image()`: Main entry - auto-selects best recognition method
- `recognize_image_with_ai()`: Uses AI vision models (ZhipuAI glm-4v-flash or OpenAI compatible)
- `recognize_with_local_ocr()`: Fallback to local PaddleOCR/EasyOCR
- `batch_recognize()`: Process multiple images
- `merge_results_to_table()`: Combine results from multiple images
- `export_to_excel()`: Export with optional template mapping

**Recognition Methods**:
1. **ZhipuAI Vision** (recommended): Uses `glm-4v-flash` model for accurate table extraction
2. **LM Studio**: Local vision models with OpenAI-compatible API
3. **PaddleOCR**: Local Chinese OCR (requires `paddleocr` package)
4. **EasyOCR**: Multilingual local OCR (requires `easyocr` package)

**Auto-Install Dependencies**:
When dependencies are missing, the module automatically installs:
- `Pillow`: Image processing
- `zhipuai`: ZhipuAI SDK for cloud AI
- `openai`: OpenAI SDK for LM Studio

### 5. Image Recognition GUI (`image_recognition_gui.py`)

**`ImageRecognitionWindow` class**:
- Standalone Tkinter window for image recognition operations
- Features:
  - Import images/folders
  - Image preview with thumbnail
  - Single/batch recognition
  - Table preview of recognized data
  - Raw text view
  - Merged data view (combining multiple images)
  - Export to Excel (direct or template-based)
  - AI settings configuration

**Access**: Menu → 文件 → 图片智能识别

### 6. Export Format Manager (`export_format_manager.py`)

Flexible output format mapping system for converting template columns to custom output schemas.

**Key Functions**:
- `load_export_formats()` / `save_export_formats()`: Persist format definitions to `config.json`
- `get_active_export_format_name(module_key)`: Get current active format
- `apply_export_format(module_key, headers, rows)`: Apply mapping transformation

**Module Keys**: `shipping_product`, `image_recognition`, `general_voucher`

**Format Definition** supports special source syntax:
- `BDV:table|code|display|fallback` - Base Data Value lookup

### 7. Reconciliation Module (`reconciliation_module.py`)

**`StandardReconciler` class** for matching two accounting datasets:
- Accepts DataFrames in "Standard Template Format" (as defined in Template.xlsx)
- Local code → Yikan code mapping via `business_partner.local_code`
- Date parsing with locale-aware strategies (DD/MM/YYYY vs YYYY/MM/DD)
- Amount reconciliation with debit/credit separation

### 8. Shipping Module (`shipping_module.py`)

Shipping & logistics cost management with its own database (`shipping.db`):
- `containers` table: Shipment header (shipment_code, container_no, costs in USD/RMB)
- `products` table: Line items with allocated costs and tax rates
- Multi-currency cost allocation with exchange rate conversion

### 9. Report Generator (`report_generator.py`)

**`ReportGenerator` class** - Comprehensive business intelligence reporting (6,444 lines):
- Profit & Loss statements with trend analysis
- Cash flow analysis (estimated from accounting data)
- Budget execution & variance analysis
- Anomaly alerts and threshold-based warnings
- Product/Category contribution analysis with Pareto charts
- Customer contribution & collections tracking
- Inventory health & slow-moving items identification
- Expense structure analysis
- Multi-entity consolidation (when multiple directories)
- Multi-currency summary
- Data quality audits

**Chart Types**: Pareto (bar + cumulative line), Scatter, Line, Bar charts via openpyxl

### 10. Supporting Modules

- **`bank_parser.py`**: PDF bank statement extraction (BAC, St. Georges Bank formats)
- **`treeview_tools.py`**: GUI utilities including smart code restoration from names (fuzzy matching against base data)
- **`excel_merger.py`**: Batch Excel file consolidation for multi-file processing

## Field Value Priority

When a field has multiple sources, priority is:
1. **Manual mapping** (user selected source column)
2. **Original data fields** (recognized from dedicated columns)
3. **Summary recognition** (extracted from summary text)
4. **Default values** (user-configured fallbacks)

## Configuration Points

### Adding Field Rules

Edit `FIELD_RULES` dict:
```python
FIELD_RULES = {
    "新字段名": {
        "type": "date",  # or "number" or "text"
        "max_int_len": 15,      # for number
        "max_decimal_len": 2,   # for number
        "max_len": 100          # for text
    }
}
```

### Adding Synonyms

Edit `FIELD_SYNONYMS` dict:
```python
FIELD_SYNONYMS = {
    "模板字段名": ["别名1", "别名2", "别名3"],
}
```

### Adding Business Type Rules

Edit `_init_recognition_rules()` in `summary_intelligence.py`:
```python
self.business_type_rules = {
    "新业务类型": {
        "keywords": ["关键词1", "关键词2"],
        "account": "XXXX",
        "type": "1",  # 1=出, 2=入, 3=借, 4=贷
        "summary_code": "XX"
    },
}
```

### Auto-Match Threshold

Default threshold is 0.6. Adjust in auto-match logic:
```python
if best_score >= 0.6 and best_col:  # Lower = more permissive
```

## Template File Requirements

`Template.xlsx` must:
- Have headers in row 1 matching `FIELD_RULES` keys
- Be in same directory as main script
- Formatting is preserved in output

## Configuration Persistence

**`config.json`** stores persistent configuration:
- `FIELD_RULES`: Field type and constraint definitions
- `FIELD_SYNONYMS`: Column name synonyms for auto-matching
- `HEADER_SCHEMES`: Saved column mapping presets
- `EXPORT_FORMATS`: Custom output format definitions per module

## Key Files

```
亿看智能识别系统.py          # Main GUI application (entry point, ~10K lines)
base_data_manager.py          # Database manager module (~1.7K lines)
summary_intelligence.py       # AI text recognition engine (~1K lines)
image_intelligence.py         # Image OCR/AI vision module (~1K lines)
image_recognition_gui.py      # Image recognition GUI window (~1.1K lines)
export_format_manager.py      # Output format mapping (~1.2K lines)
reconciliation_module.py      # Accounting reconciliation (~1K lines)
shipping_module.py            # Shipping cost management (~2.4K lines)
report_generator.py           # Business intelligence reports (~6.4K lines)
bank_parser.py                # PDF bank statement extraction
treeview_tools.py             # GUI utilities with fuzzy matching
excel_merger.py               # Multi-file batch processing
Template.xlsx                 # Required template (headers + comments)
config.json                   # Persistent configuration
base_data.db                  # SQLite database (auto-created)
shipping.db                   # Shipping module database (auto-created)
基础数据/基础数据/            # Source Excel files for base data import
```

## Auto-Match Algorithm

**`score_similarity()`** uses multi-strategy scoring:
1. Exact match → 1.0
2. Synonym match → 0.9
3. Containment → 0.85
4. Difflib ratio → 0.0-0.8 (weighted)

Default threshold: 0.6 (lower = more permissive)

## Known Limitations

1. **Template dependency**: Application requires `Template.xlsx` in working directory
2. **Single sheet processing**: Multi-sheet files require multiple conversion runs
3. **Memory-based**: Entire source file loaded into memory (no pagination)
4. **Number truncation**: Integers exceeding `max_int_len` are silently truncated
5. **First-match wins**: Conflicting keywords in summary use first matched rule

## Data Flow Summary

```
User selects Excel → Choose sheet → Select mode → Auto-match columns (optional)
    → Enable Smart Recognition (optional) → Set Default Values (optional)
    → Preview & Confirm → Convert & Export
```

**Recognition Pipeline** (when Smart Recognition enabled):
```
Input row → _recognize_from_fields() [date/amount/rate columns]
         → recognize(summary) [AI analysis of summary text]
         → Apply field priority: Manual > Original > Recognized > Default
         → Cache results in smart_recognition_cache
```
