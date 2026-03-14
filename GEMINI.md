# Project: 亿看智能识别系统 (Yikan Intelligent Recognition System)

This directory contains a Python-based GUI application designed to automate the conversion of Excel data into a standardized general voucher format. It uses a template-based approach to map, format, and validate data from various source Excel files.

## Key Files

*   **`亿看智能识别系统.py`**: The main application script. It launches a Tkinter GUI that allows users to:
    *   Load a source Excel file and a specific worksheet.
    *   Automatically map source columns to template columns based on name similarity and synonyms.
    *   Manually adjust column mappings.
    *   Convert and export the data into a new Excel file based on `Template.xlsx`.
*   **`Template.xlsx`**: The required template file. The application reads headers and comments from this file to define the target structure and validation rules. **This file must exist in the same directory as the script.**
*   **`11月bac对账建议.xlsx`**: A sample or working Excel file likely used as input data for the conversion process.

## Features

*   **Automatic Column Matching:** Uses fuzzy matching (`difflib`) and a synonym dictionary to guess the correspondence between source and target columns.
*   **Data Formatting:** Automatically formats dates (to `YYYYMMDD`), numbers (precision and scale), and truncates text based on defined rules.
*   **GUI Interface:** User-friendly interface for file selection and mapping verification.
*   **Smart Recognition:** Can optionally extract structured data from "Summary" (摘要) fields using keyword matching against a base database.
*   **Base Data Management:** Includes a module to manage base data (Departments, Currencies, Warehouses, etc.) stored in a local database, which supports the smart recognition feature.

## Dependencies

The project requires the following Python packages:

*   `pandas`
*   `openpyxl`
*   `tkinter` (Standard library)

## Usage

1.  **Install Dependencies:**
    ```bash
    pip install pandas openpyxl
    ```

2.  **Run the Application:**
    ```bash
    python 亿看智能识别系统.py
    ```

3.  **Operation:**
    *   The GUI will open.
    *   Select your source Excel file (e.g., `11月bac对账建议.xlsx`).
    *   Select the worksheet containing the data.
    *   Click "Automatic Recognition Match" (自动识别匹配) to let the system guess column mappings.
    *   Review and adjust mappings in the dropdowns.
    *   Click "Start Conversion and Export" (开始转换并导出) to save the result.

## Field Specifications (Output Rules)

The application enforces the following rules for output data, corresponding to the comments in `Template.xlsx`:

| Field Name (Target) | Description | Format/Length Limit |
| :--- | :--- | :--- |
| **凭证日期** (Document Date) | Transaction date. Defaults to current date if empty. | `YYYYMMDD` |
| **序号** (Serial No.) | Bundles entries into one document if identical. | Max 4 chars |
| **会计凭证No.** (Accounting No.) | Transaction voucher number. | Max 30 chars |
| **摘要** (Summary) | Description of the transaction. | Max 200 chars |
| **类型** (Type) | Debit/Credit type. | 1 char (1=Out, 2=In, 3=Debit, 4=Credit) |
| **科目编码** (Subject Code) | Accounting subject code. | Max 8 chars (Name can be 100 chars) |
| **往来单位编码** (Partner Code) | Business partner code. | Max 30 chars |
| **金额** (Amount) | Transaction amount. | Max 15 digits integer, 2 decimal places |
| **外币金额** (Foreign Amt) | Foreign currency amount. | Max 15 digits integer, 2 decimal places |
| **汇率** (Exchange Rate) | Exchange rate if foreign currency is used. | Max 14 digits integer, 4 decimal places |
| **部门** (Department) | Department code or name. | Code max 14 chars, Name max 50 chars |

## Configuration

The conversion logic is primarily defined in `亿看智能识别系统.py`:

*   **`FIELD_RULES`**: Dictionary defining data types (`date`, `text`, `number`), maximum lengths, and number precision for each target column.
*   **`FIELD_SYNONYMS`**: Dictionary defining lists of potential source column names that map to specific target columns (e.g., mapping "记账日期" to "凭证日期").