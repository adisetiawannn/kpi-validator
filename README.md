# KPI Validator

> Automated, config-driven KPI reporting pipeline built with Python and pandas — aggregates multi-source data, computes period-over-period metrics, and generates formatted Excel reports.

**English** | [Bahasa Indonesia](README.id.md)

[![Python](https://img.shields.io/badge/Python-3.11-blue.svg)](https://www.python.org/)
[![pandas](https://img.shields.io/badge/pandas-2.0+-150458.svg)](https://pandas.pydata.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Stable-brightgreen.svg)]()

---

## Table of Contents

- [Overview](#overview)
- [Tech Stack](#tech-stack)
- [Architecture](#architecture)
- [KPI Coverage](#kpi-coverage)
- [Key Highlights](#key-highlights)
- [Output Sample](#output-sample)
- [Attribution](#attribution)
- [Getting Started](#getting-started)
- [About](#about)
- [License](#license)

---

## Overview

**KPI Validator** automates the monthly KPI reporting workflow for a telecommunications business unit. What used to be a manual, error-prone process of consolidating spreadsheets across multiple data sources is reduced to a single command that validates inputs, computes the full metric suite, and exports a formatted, presentation-ready Excel workbook.

The pipeline handles **4 KPI families across 4 business segments**, computing month-to-date (MtD), year-to-date (YtD), achievement-vs-target, month-over-month (MoM), and year-over-year (YoY) metrics — with graceful handling of missing data, cut-off dates, and outlook projections.

> **Note** — This is a public portfolio version. All company, brand, product, and segmentation identifiers have been replaced with generic placeholders. The calculation logic and architecture are unchanged.

---

## Tech Stack

- **Python 3.11**
- **pandas** — data manipulation & aggregation
- **openpyxl** — Excel generation & formatting
- **xlrd** — legacy Excel read support

---

## Architecture

The pipeline runs as a deterministic **5-step flow**:

```
┌─────────────┐   ┌─────────────┐   ┌─────────────┐   ┌─────────────┐   ┌─────────────┐
│  1.VALIDATE │ → │  2. INPUT   │ → │ 3. PROCESS  │ → │4. AGGREGATE │ → │  5. EXPORT  │
│             │   │             │   │             │   │             │   │             │
│ Check source│   │ Confirm     │   │ Compute 4   │   │ Roll up     │   │ Styled      │
│ files exist │   │ cutoffs &   │   │ KPIs × 4    │   │ segments    │   │ Excel       │
│             │   │ scaling     │   │ segments    │   │ (sums)      │   │ workbook    │
└─────────────┘   └─────────────┘   └─────────────┘   └─────────────┘   └─────────────┘
```

### Project Structure

```
kpi-validator/
├── main.py                  # Entry point — orchestrates the 5-step pipeline
├── config.py                # Single source of truth: paths, sheets, mappings, thresholds
├── requirements.txt
├── .env.example
├── tools/
│   └── generate_sample_data.py   # Synthetic sample data generator
└── src/
    ├── loaders/             # File I/O — Excel & CSV ingestion
    │   ├── sales_loader.py
    │   ├── collection_loader.py
    │   └── target_loader.py
    ├── processors/
    │   ├── input_handler.py        # Interactive cut-off & scaling confirmation
    │   ├── aggregator.py           # Cross-segment roll-ups
    │   ├── validator.py
    │   └── segments/               # Per-KPI calculation engines
    │       ├── bs_processor.py     # Revenue — SEG_A
    │       ├── gs_processor.py     # Revenue — SEG_B
    │       ├── dss_processor.py    # Revenue — SEG_C
    │       ├── dps_processor.py    # Revenue — SEG_D
    │       ├── ngtma_processor.py  # Growth KPI
    │       ├── sales_processor.py  # Sales connectivity (3 sub-indicators)
    │       ├── collection_processor.py  # Collection performance (4 indicators)
    │       └── _historical_helper.py    # Shared YoY historical loader
    └── exporters/
        └── excel_exporter.py   # Formatted workbook generation
```

---

## KPI Coverage

| KPI Family | Indicators | Segments | Metrics |
|------------|-----------|----------|---------|
| **Revenue** | Revenue | SEG_A, SEG_B, SEG_C, SEG_D | MtD, YtD, Achievement, MoM, YoY |
| **Growth** | GROWTH | SEG_A, SEG_B, SEG_C, SEG_D | MtD, YtD, Achievement, MoM, YoY |
| **Sales Connectivity** | PROD_A, PROD_B, PROD_C | SEG_A, SEG_B, SEG_C, SEG_D | MtD, YtD, Achievement, MoM, YoY |
| **Collection** | COLL_A, COLL_B, COLL_C, COLL_D | SEG_A–D + aggregates | Achievement, MoM/YoY |

---

## Key Highlights

A few decisions worth calling out for anyone reading the code:

- **Zero hardcoded business strings in logic** — sheet names, segment codes, and product mappings are all resolved through `config.py`, so a renamed source file or sheet is a one-line change.
- **Multi-source ingestion** — consolidates revenue, growth, sales connectivity, and collection data from both Excel and CSV sources with differing schemas and encodings (UTF-8 BOM, UTF-16 LE, `;`-separated CSVs).
- **Defensive data handling** — `_safe_div()` and empty-DataFrame fallbacks ensure that a missing row or a zero target degrades gracefully into `None`/zeros rather than crashing the run.
- **Outlook projections** — interactive cut-off confirmation lets analysts project incomplete months using manual scaling inputs.
- **Single processor pattern, reused** — the four revenue segment processors and the growth/connectivity/collection engines all follow the same load → compute → derive → return contract, keeping the codebase predictable.
- **Separation of concerns** — loaders never compute, processors never format, exporters never load. Each layer has one job.

---

## Output Sample

The pipeline produces a styled Excel workbook with two sheets — a concise MTD/YTD summary and a full monthly detail view. Color-coded headers, borders, and currency/percentage number formats are applied automatically, so the report is ready to share with no post-processing.

![KPI Validator — Excel output sample](docs/images/kpi-validator-output.png)

> *Data shown is synthetic sample data generated by `tools/generate_sample_data.py`.*

---

## Attribution

- All sample data files under `data/input/` are **synthetic**, generated by `tools/generate_sample_data.py`. No real company data is included in this repository.
- Built with open-source libraries: [pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/), [xlrd](https://xlrd.readthedocs.io/).

---

## Getting Started

### Prerequisites

- Python 3.11+
- Input data files placed in `data/input/` (see `config.py` for expected filenames — synthetic samples are included)

### Installation

```bash
# Clone the repository
git clone https://github.com/adisetiawannn/kpi-validator.git
cd kpi-validator

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Usage

```bash
python main.py --month 6 --year 2026
```

The pipeline will interactively prompt for:
- **Territory** — the reporting unit
- **Per-segment cut-off** (`valid_until`) — last month of confirmed actuals
- **Outlook scaling** — manual projection inputs for incomplete months

Custom file paths can be passed explicitly:

```bash
python main.py --month 6 --year 2026 \
  --revenue "data/input/REVENUE FILE.xlsx" \
  --target "data/input/data target.xlsx" \
  --historical "data/input/data historical.xlsx"
```

The generated report is written to `data/output/`.

---

## About

This project started as an internal automation tool to replace a recurring, multi-hour manual reporting task with a single reproducible command. It is published here as a portfolio piece to demonstrate practical data-pipeline design: config-driven architecture, defensive data handling, and clean separation between ingestion, computation, and presentation layers.

---

## License

This project is licensed under the [MIT License](LICENSE).
