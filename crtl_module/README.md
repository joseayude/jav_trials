# CRTL Importer

Python CLI module to process supplier test results (CRTL) from Aumovio and import them into the TVW Export TD.

Replaces the legacy VBA macro `TVW_Lieferantenimport_04`.

---

## Requirements

- Python 3.11+
- WSL or Linux terminal

## Setup

```bash
cd crtl_module
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
# Single CRTL file
python main.py --crtl path/to/CRTL.xlsx --td path/to/TVW_Export_TD.xlsx

# Multiple CRTL files
python main.py --crtl path/to/CRTL_1.xlsx path/to/CRTL_2.xlsx --td path/to/TVW_Export_TD.xlsx

# Validate only, no output written
python main.py --crtl path/to/CRTL.xlsx --td path/to/TVW_Export_TD.xlsx --dry-run
```

## Project Structure

```
crtl_module/
├── config/                  # Column and result mapping definitions
│   ├── column_mapping.yaml
│   └── result_mapping.yaml
├── models/                  # Pydantic data models
├── excel/                   # Excel reader and writer
├── processing/              # Core business logic
├── validation/              # Input schema validation
├── tests/                   # Unit tests + fixtures
├── main.py                  # CLI entry point
└── requirements.txt
```

## Configuration

Column names and result mappings are defined in `config/` and can be updated without touching code — relevant when Aumovio changes their CRTL structure.

## Running Tests

```bash
pytest tests/ --cov
```

## Required CRTL Columns

The following columns must be present in the Aumovio CRTL file:

| Column | Notes |
|---|---|
| `ForeignID` | Primary match key |
| `VW Requirement` | Fallback match key |
| `Object_UniqueID` | Test case reference |
| `Test Result Logik` | Core test result |
| `051_VerificationMethod` | |
| `Defect ID/ Open Point` | |
| `CustomerReferenceID` | |
| `Planned for delivery` | |
| `Supplier Comment` | |
| `Defect - Summary` | |
| `Product Class` | |
| `Latest Test Planning` | **New — Aumovio to add** |
| `Engineering Judgement` | **New — Aumovio to add** |
| `SW L3 Regression` | **New — Aumovio to add** |