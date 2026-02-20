# Analysis Scripts Directory

This directory contains diagnostic, validation, and data-fixing scripts organized by purpose.

## Structure

### `/validators/`

**Purpose**: Verify data quality and completeness

Scripts that **check** if source data and generated output are correct without making changes.

**Use when**:
- Before running batch generation
- After loading new source Excel file
- When diplomas look incorrect

**Examples**:
- `validate_subjects.py` - Check if all template subjects exist in source
- `validate_data_completeness.py` - Verify all student data is present
- `validate_excel_structure.py` - Check row/column structure matches expected

**Running**:
```bash
cd analysis/validators
python validate_subjects.py
```

---

### `/inspectors/`

**Purpose**: Analyze and debug system state

Scripts that **examine** internals without modifying data. Useful for understanding what the system is doing.

**Use when**:
- Troubleshooting generation issues
- Understanding data flow
- Debugging unexpected behavior

**Examples**:
- `analyze_excel_structure.py` - Display row/column layout
- `analyze_generated_diploma.py` - Show what was written to output Excel
- `analyze_subject_mappings.py` - Check how subjects are being matched
- `analyze_grade_conversions.py` - Verify which grades are assigned

**Running**:
```bash
cd analysis/inspectors
python analyze_generated_diploma.py
```

---

### `/fixers/`

**Purpose**: Repair and update source data

Scripts that **modify** Excel files to fix missing or incorrect data.

⚠️ **Use carefully** - These modify the source file!

**Use when**:
- Source Excel is missing hours/credits
- Need to update structures
- Preparing new academic year data

**Examples**:
- `fix_attestation_hours.py` - Add missing attestation module hours
- `fix_missing_hours.py` - Fill in empty hours/credits cells
- `fix_illformed_subject_names.py` - Normalize inconsistent name formatting

**Running**:
```bash
cd analysis/fixers
python fix_attestation_hours.py  # Creates backup before modifying
```

---

## Recommended Workflow

1. **Validate source data** (before batch generation)
   ```bash
   python validators/validate_subjects.py
   python validators/validate_excel_structure.py
   ```

2. **Run batch generation**
   ```bash
   python -m batch._generate_it --year 2025-2026
   ```

3. **Inspect results** (if problems occur)
   ```bash
   python inspectors/analyze_generated_diploma.py
   python inspectors/analyze_grade_conversions.py
   ```

4. **Fix source data** (if needed)
   ```bash
   python fixers/fix_attestation_hours.py
   python fixers/fix_missing_hours.py
   ```

5. **Revalidate and regenerate**
   ```bash
   python validators/validate_data_completeness.py
   python -m batch._generate_it --year 2025-2026
   ```

---

## Legacy Scripts

The following scripts from the old codebase should be migrated here:

### Should move to `/validators/`:
- `check_subject_matching.py`
- `check_all_subjects.py`
- `check_attestation_data.py`
- `verify_attestation_electives.py`
- `verify_generated_diploma.py`
- `verify_structure_match.py`
- `comprehensive_subject_match.py`

### Should move to `/inspectors/`:
- `inspect_data_source.py`
- `inspect_generated_diploma.py`
- `inspect_data_rows.py`
- `inspect_subjects.py`
- `debug_student_data.py`
- `debug_excel_structure.py`
- `debug_excel_rows.py`

### Should move to `/fixers/`:
- `fill_attestation_hours.py` (v2)
- `fill_grades.py`
- `clean_batch.py`

### Can be deleted (superseded by new architecture):
- `debug_on72_matching.py` (now handled by normalize_key)
- `synthesize_grades.py` (use test generation instead)
- `extract_*.py` scripts (subjects now in config)
- `search_*.py` scripts (subjects searchable via config API)
- `dump_*.py` scripts (debugging info available via inspectors)
- `compare_*.py` scripts (validation now centralized)

---

## Adding New Analysis Scripts

1. Choose category: validator, inspector, or fixer
2. Create file in appropriate subdirectory
3. Use clear prefix: `validate_`, `analyze_`, or `fix_`
4. Add docstring explaining purpose
5. Make it standalone (runnable with `python filename.py`)

Example:
```python
# analysis/validators/validate_new_feature.py
"""
Validate that new feature is working.

Run with: python validate_new_feature.py
"""

if __name__ == "__main__":
    # Validation code here
    print("✅ Feature validated successfully")
```

---

## See Also

- [ARCHITECTURE.md](../docs/ARCHITECTURE.md) - Overall system design
- [README.md](../README.md) - Quick start guide
