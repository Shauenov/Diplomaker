# Diploma Automation System

Automatic generation of bilingual diploma supplements (Kazakh/Russian) for IT and Accounting students.

**Status**: ✅ Phase 1 (Foundation) ✅ Phase 2 (Data Layer) ✅ Phase 3 (Testing) | Phase 4 Planned

---

## Quick Start

### Prerequisites

- Python 3.8+
- pip

### Installation

1. **Clone/navigate to project directory**
   ```bash
   cd diploma_automation
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify configuration**
   ```bash
   python -c "from config import settings; print(f'Source file: {settings.SOURCE_FILE}')"
   ```

4. **Run tests** (optional but recommended)
   ```bash
   pip install -r requirements-dev.txt
   python -m pytest tests/ -v
   # Expected: 122/126 tests passing (97%)
   ```

### Generate Diplomas

**For IT Program (3F groups)**:
```bash
python -m batch._generate_it --year 2025-2026 --out ./output
```

**For Accounting Program (3D groups)** (coming soon):
```bash
python -m batch._generate_accounting --year 2025-2026 --out ./output
```

The system will:
1. Read student grades from source Excel
2. Convert percentage scores to letter grades, GPA, traditional grades
3. Generate bilingual diplomas (KZ + RU)
4. Save to output directory with student names

---

## Project Structure

```
diploma_automation/
├── config/                      # Configuration management
│   ├── settings.py             # Global constants (paths, thresholds)
│   ├── languages.py            # Language definitions (KZ, RU)
│   ├── programs.py             # Program configs (IT, Accounting)
│   └── __init__.py
│
├── core/                        # Core business logic
│   ├── models.py               # Data classes
│   ├── converters.py           # Grade conversion engine
│   ├── utils.py                # Shared utilities
│   ├── exceptions.py           # Custom exceptions
│   └── __init__.py
│
├── data/                        # Data layer (coming soon)
│   ├── excel_parser.py         # Read source Excel
│   ├── excel_generator.py      # Write output diplomas
│   └── __init__.py
│
├── batch/                       # Batch processing scripts
│   ├── _generate_it.py         # IT program batch generator
│   ├── _generate_accounting.py # Accounting program batch processor
│   └── utils.py
│
├── analysis/                    # Analysis & diagnostic tools
│   ├── validators/             # Data quality checks
│   ├── inspectors/             # Debug & analysis scripts
│   ├── fixers/                 # Data repair scripts
│   └── README.md               # How to use analysis tools
│
├── tests/                       # Test suite (✅ 122/126 passing)
│   ├── conftest.py             # Shared fixtures
│   ├── test_converters.py      # Grade conversion tests (47/47 passing)
│   ├── test_models.py          # Domain model tests (34/34 passing)
│   ├── test_parsers.py         # Excel parsing tests (21/24 passing)
│   └── test_generators.py      # Diploma generation tests (20/21 passing)
│
├── docs/                        # Documentation
│   ├── README.md               # (This file)
│   ├── ARCHITECTURE.md         # System design
│   ├── DATA_SCHEMA.md          # Excel structure
│   └── DEPLOYMENT.md           # Production setup
│
├── output_diplomas/            # Generated files (.gitignored)
├── .gitignore
├── requirements.txt
├── requirements-dev.txt
└── main.py                     # CLI entry point (coming soon)
```

---

## Key Features

### ✅ Phase 1 - Foundation (Complete)

- [x] Bilingual support (Kazakh/Russian)
- [x] Grade conversion (points → letter, GPA, traditional grades)
- [x] Domain models with type hints
- [x] Centralized configuration
- [x] Exception hierarchy

### ✅ Phase 2 - Data Layer (Complete)

- [x] Unified diploma generator
- [x] Excel parser layer (ExcelParser class)
- [x] Excel generator (DiplomaGenerator class)
- [x] Batch processor (_generate_it.py)
- [x] Production verification (200+ diplomas generated)
- [x] IT program support (3F-1, 3F-2, 3F-3, 3F-4)

### ✅ Phase 3 - Testing (Complete)

- [x] Comprehensive test suite (126 tests, 97% passing)
  - [x] Grade converter tests (47/47 passing)
  - [x] Domain model tests (34/34 passing)
  - [x] Excel generator tests (20/21 passing)
  - [x] Excel parser tests (21/24 passing)
- [x] pytest configuration with markers
- [x] Shared test fixtures and utilities
- [x] Testing documentation

### ⏳ Phase 4 - Deployment & Polish (Planned)

- [ ] Production deployment guide
- [ ] Structured logging infrastructure
- [ ] Additional language support (Turkish, English)
- [ ] Web interface for batch upload
- [ ] REST API for system integration
- [ ] PDF diploma export
- [ ] Accounting program support

---

## Configuration

### Global Settings

Edit [config/settings.py](config/settings.py):

```python
# Source Excel file path
SOURCE_FILE = r"path\to\2025-2026 диплом бағалары.xlsx"

# Excel structure (row/column indices)
ROW_SUBJECT_NAMES = 1
ROW_HOURS = 3
ROW_DATA_START = 5

# Grade conversion thresholds
GRADE_THRESHOLDS = {
    95: {"letter": "A", "gpa": 4.0, ...},
    70: {"letter": "C+", "gpa": 2.33, ...},
    ...
}
```

### Program Configuration

Edit [config/programs.py](config/programs.py) to:
- Add new programs
- Change sheet names
- Update subject lists
- Modify page structure

---

## Usage Examples

### 1. Batch Generate All Diplomas

```bash
python -m batch._generate_it
```

### 2. Test with Sample Data

```bash
# Create test file with random scores
python core/converters.py --test

# Generate from test file
python -m batch._generate_it --source test_grades_filled.xlsx --test
```

### 3. Validate Source Data

```bash
cd analysis/validators
python validate_subjects.py
python validate_excel_structure.py
```

### 4. Inspect Generated Output

```bash
cd analysis/inspectors
python analyze_generated_diploma.py
```

### 5. Fix Missing Data

```bash
cd analysis/fixers
python fix_attestation_hours.py  # Add missing hours
```

---

## Grade Conversion Reference

The system converts percentage scores to multiple grade formats:

| Range | Letter | GPA  | Kazakh             | Russian         |
|-------|--------|------|-------------------|-----------------|
| 95-100| A      | 4.0  | 5 (өте жақсы)    | 5 (отлично)     |
| 90-94 | A-     | 3.67 | 5 (өте жақсы)    | 5 (отлично)     |
| 85-89 | B+     | 3.33 | 4 (жақсы)        | 4 (хорошо)      |
| 80-84 | B      | 3.0  | 4 (жақсы)        | 4 (хорошо)      |
| 75-79 | B-     | 2.67 | 4 (жақсы)        | 4 (хорошо)      |
| 70-74 | C+     | 2.33 | 4 (жақсы)        | 4 (хорошо)      |
| 65-69 | C      | 2.0  | 3 (қанағат)      | 3 (удовл)       |
| 50-64 | D      | 1.0  | 3 (қанағат)      | 3 (удовл)       |
| <50   | F      | 0.0  | 2 (неудов)       | 2 (неуд)        |

---

## Troubleshooting

### Problem: "Cannot find source file"

**Solution**: Update `SOURCE_FILE` in [config/settings.py](config/settings.py)

```python
SOURCE_FILE = r"C:\path\to\your\diploma_grades.xlsx"
```

### Problem: "Unknown program"

**Solution**: Use valid program code with `--program` flag:

```bash
python -m batch._generate_it --program IT
```

### Problem: "Excel structure doesn't match"

**Solution**: Run validation before batch generation:

```bash
cd analysis/validators
python validate_excel_structure.py
```

### Problem: Grades not appearing in diplomas

**Solution**: Check if data is being read:

```bash
cd analysis/inspectors
python analyze_generated_diploma.py
```

---

## Development

### Install Development Dependencies

```bash
pip install -r requirements-dev.txt
```

### Run Tests

```bash
pytest tests/ -v --cov=core,config,data
```

### Code Quality Checks

```bash
# Format code
black core/ config/ batch/

# Check style
flake8 core/ config/ batch/

# Type checking
mypy core/ config/ batch/

# Import sorting
isort core/ config/ batch/
```

---

## Documentation

- [ARCHITECTURE.md](docs/ARCHITECTURE.md) - System design, data flow, module details
- [DATA_SCHEMA.md](docs/DATA_SCHEMA.md) - Excel file structure requirements
- [DEPLOYMENT.md](docs/DEPLOYMENT.md) - Production deployment guide
- [Analysis Tools](analysis/README.md) - Using validators, inspectors, fixers

---

## License

Internal use only. Contact IT Department for distribution rights.

---

## Support

### Questions or Issues?

1. Check [analysis/README.md](analysis/README.md) for diagnostic tools
2. Run validators to identify data problems
3. Review logs in `diploma_automation.log`
4. Contact: [IT Support Email]

### Reporting Bugs

Include:
- Steps to reproduce
- Error message (full traceback)
- Source file name
- Python version (`python --version`)

---

## Changelog

### Version 1.0 (Current)

**Phase 1 - Foundation** ✅
- Centralized configuration
- Core data models
- Grade conversion engine
- Shared utilities

**Phase 2 - Consolidation** 🔄 (In progress)
- Unified diploma generator
- Excel parser layer
- Analysis script organization

**Phase 3 - Testing** ⏳ (Coming soon)
- Test suite
- Logging infrastructure

---

## Contributors

- Initial development: IT Automation Team
- Kazakh diploma template: Education Department
- Russian translation: Translation Team

---

*Last updated: February 19, 2026*
