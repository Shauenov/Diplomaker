# Technical Specification: Student Diploma Data Parser Agent
**Version:** 1.0  
**Scope:** Accountant specialty (Бухгалтер) — Kazakh-language diploma template  
**Status:** Active

---

## 1. Overview

### 1.1 Purpose

This document specifies the requirements for an agent that reads student grade data from an Excel source file and extracts all data fields needed to fill in diploma supplement templates (приложение к диплому). The agent must map Excel data to template placeholders for each student.

### 1.2 Template Variants (Full System)

There are **4 template variants** that the full system must support:

| # | Specialty | Language | File Reference |
|---|-----------|----------|----------------|
| 1 | **Accountant (Бухгалтер)** — *this TZ* | **Kazakh** | `каз__2_.docx` |
| 2 | Accountant (Бухгалтер) | Russian | `рус.docx` |
| 3 | IT (Computer Science) | Kazakh | `3Ф_ПРИЛОЖ_2026_КАЗАК_ШАБЛОН__3_соңғы___копия.docx` |
| 4 | IT (Computer Science) | Russian | `3Ф_ПРИЛОЖ_2025_РУС_ШАБЛОН_соңғы.docx` |

**This TZ focuses exclusively on Variant 1 (Accountant, Kazakh).**

---

## 2. Source Data File

### 2.1 File

```
2025-2026_диплом_бағалары__ТОЛЫҚ__қызыл_диплом_жазылған_соңғысы_точно__1_.xlsx
```

### 2.2 Relevant Sheets (Accountant Groups)

The workbook contains sheets for both accountant (`3D`) and IT (`3Ғ`) groups. For the accountant Kazakh template, parse only sheets whose names begin with `3D`:

| Sheet Name | Description |
|------------|-------------|
| `3D-1` | Accountant group 1 |
| `3D-2` | Accountant group 2 |

> The agent must iterate over all `3D-*` sheets automatically (do not hardcode sheet names).

### 2.3 Header Row Structure

The spreadsheet uses a **multi-row header** structure with merged cells:

| Row | Content |
|-----|---------|
| Row 1 | Subject sequence numbers |
| Row 2 | Subject group names / module names (merged across sub-columns) |
| Row 3 | Sub-module / learning outcome (ОН) names (merged across sub-columns) |
| Row 5 | Credit hours per subject (e.g., `96с-4к` = 96 hours, 4 credits) |
| Row 6 | Sub-column type identifiers: `п` (points/percentage), `б` (GPA letter), `цэ` (letter grade), `трад` (traditional grade) |
| **Row 7+** | **Student data rows** |

### 2.4 Student Row Fields

For each student, the following columns are always present:

| Column | Field |
|--------|-------|
| **Col A (1)** | Sequential number within group |
| **Col B (2)** | Full student name (Фамилия Имя Отчество) |
| **Col 150** | Year of enrollment (`Год поступления`) — same for all students in sheet |
| **Col 151** | Year of graduation (`Год выпуска`) — same for all students in sheet |
| **Col 152** | Diploma serial number (`Диплом номер`) — **individual per student** |

> **Note:** Columns 150–152 use row 5 as their header (not row 2/3). The value in row 6 for col 150 is the shared year (e.g., `2023`), for col 151 it is the graduation year (e.g., `2026`), and for col 152 the placeholder text `(каждому отдельно)`.

### 2.5 Grade Sub-Columns

Each graded subject has up to 4 consecutive sub-columns:

| Sub-column type | Meaning | Example value |
|-----------------|---------|---------------|
| `п` | Percentage score (0–100) | `86` |
| `б` | GPA score (0.0–4.0) | `3.33` |
| `цэ` | Letter grade (e.g. A, B+, C-) | `B+` |
| `трад` | Traditional Kazakh grade | `4 (жақсы)` |

**The diploma template uses the `трад` (traditional) grade.** This is the value to extract.

Traditional grade values and their meanings:
- `5 (өте жақсы)` — Excellent (95–100%)
- `4 (жақсы)` — Good (75–94%)
- `3 (қанағат)` — Satisfactory (50–74%)
- `2 (қанағатсыз)` — Unsatisfactory (below 50%)

> **Important:** Some subjects only have `п` (3 sub-columns: п, б, цэ — no трад). In this case, derive the traditional grade from the percentage score using the conversion table in section 2.7.

### 2.6 Complete Column Mapping (Accountant, Sheet `3D-1`)

This mapping is **stable across all `3D-*` sheets** (same column positions).

#### 2.6.1 General Education Subjects (Жалпы білім берудің пәндері)

| # in diploma | Subject (Kazakh) | `п` col | `цэ` col | `трад` col | Hours |
|---|---|---|---|---|---|
| 1 | Қазақ тілі | 3 | 5 | 6 | 96с-4к |
| 2 | Қазақ әдебиеті | 7 | 9 | 10 | 96с-4к |
| 3 | Орыс тілі және әдебиеті | 11 | 13 | 14 | 96с-4к |
| 4 | Ағылшын тілі | 15 | 17 | 18 | 216с-9к |
| 5 | Қазақстан тарихы | 19 | 21 | 22 | 96с-4к |
| 6 | Математика | 23 | 25 | 26 | 120с-5к |
| 7 | Информатика | 27 | 29 | 30 | 48с-2к |
| 8 | Алғашқы әскери және технологиялық дайындық | 31 | 33 | 34 | 96с-4к |
| 9 | Дене тәрбиесі | 35 | 37 | 38 | 120с-5к |
| 10 | География | 39 | 41 | 42 | 120с-5к |
| 11 | Биология | 43 | 45 | 46 | 120с-5к |
| 12 | Физика | 47 | 49 | 50 | 72с-3к |
| 13 | Графика және жобалау | 51 | 53 | 54 | 72с-3к |

#### 2.6.2 Basic Modules (Базалық модульдер)

| # in diploma | Module (Kazakh) | `п` col | `цэ` col | `трад` col | Hours |
|---|---|---|---|---|---|
| 14 | БМ 01 Дене қасиеттерін дамыту және жетілдіру | 55 | 57 | 58 | 240с-10к |
| 15 | БМ 02 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану | 59 | 61 | 62 | 72с-3к |
| 16 | БМ 03 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану | 63 | 65 | 66 | 72с-3к |
| 17 | БМ 04 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану | 67 | 69 | 70 | 24с-1к |

#### 2.6.3 Professional Module 1 — КМ 1 (Бизнестің мақсаттары мен түрлерін түсіну)

| # in diploma | Learning Outcome (Kazakh) | `п` col | `трад` col | Hours |
|---|---|---|---|---|
| 18 | КМ 1 *(module header — no grade row; skip or use aggregated)* | — | — | — |
| 19 | ОН 1.1 Бизнестің мақсаттары мен түрлерін, олардың негізгі мүдделі тараптармен және сыртқы ортамен өзара әрекеттесуін түсіну | 71 | *derived* | 96с-4к |
| 20 | ОН 1.2 Көрсеткіштік және логарифмдік функциялар... | 74 | *derived* | 96с-4к |
| 21 | ОН 1.3 Қаржылық есептіліктің мәні мен мақсатын түсіну... | 77 | *derived* | 120с-5к |
| 22 | ОН 1.4 Маркетингтің негізгі тұжырымдамаларды түсіну... | 80 | 83 | 72с-3к |

> **Note for ОН 1.1–1.3:** These have only 3 sub-columns (п, б, цэ — **no трад column**). Derive traditional grade from the `п` (percentage) value using the conversion table in section 2.7.

#### 2.6.4 Professional Module 2 — КМ 2 (Кәсіптік салада тілдік дағдылар)

| # in diploma | Learning Outcome (Kazakh) | `п` col | `трад` col | Hours |
|---|---|---|---|---|
| 23 | КМ 2 *(module header — skip)* | — | — | — |
| 24 | ОН 2.1 Академиялық деңгейде Ағылшын тілінің дағдылары | 84 | 87 | 168с-7к |
| 25 | ОН 2.2 Кәсіби салада Ағылшын тілінің B2 деңгейінде | 88 | 91 | 48с-2к |
| 26 | ОН 2.3 Іскерлік мақсатта қазақ тілін қолдану | 92 | 95 | 48с-2к |
| 27 | ОН 2.4 Іскерлік мақсатта түрік тілін қолдану | 96 | 99 | 72с-3к |

#### 2.6.5 Professional Module 3 — КМ 3 (Бухгалтерлік есептілік)

| # in diploma | Learning Outcome (Kazakh) | `п` col | `трад` col | Hours |
|---|---|---|---|---|
| 28 | КМ 3 *(module header — skip)* | — | — | — |
| 29 | ОН 3.1 Басқару ақпаратының сипатын түсіну... | 100 | *derived* | 120с-5к |
| 30 | ОН 3.2 Еңбек қатынастарына қатысты заңды түсіну... | 103 | *derived* | 48с-2к |
| 31 | ОН 3.3 Іскерлік шешім қабылдау математикалық құралдары... | 106 | *derived* | 72с-3к |
| 32 | ОН 3.4 Негізгі экономикалық принциптер... | 109 | *derived* | 120с-5к |

> **Note for ОН 3.1–3.4:** These have only 3 sub-columns (п, б, цэ — **no трад column**). Derive from `п`.

#### 2.6.6 Professional Module 4 — КМ 4 (Шаруашылық-қаржылық талдау)

| # in diploma | Learning Outcome (Kazakh) | `п` col | `трад` col | Hours |
|---|---|---|---|---|
| 33 | КМ 4 *(module header — skip)* | — | — | — |
| 34 | ОН 4.1 Инвестициялар мен қаржыландыруды бағалау... | 112 | *derived* | 72с-3к |
| 35 | ОН 4.2 Ұйымдарға өнімділікті басқару... | 115 | *derived* | 120с-5к |
| 36 | ОН 4.3 Салық жүйесінің жұмыс істеуі... | 118 | *derived* | 72с-3к |
| 37 | ОН 4.4 ХҚЕС стандарттарына сәйкес операцияларды есепке алу | 121 | *derived* | 72с-3к |
| 38 | ОН 4.5 Бизнес статистикадағы негізгі түсініктер... | 121 | *derived* | 72с-3к |
| 39 | ОН 4.6 Бухгалтерлік есептің ақпараттық жүйелері | 124 | *derived* | 72с-3к |
| 40 | ОН 4.7 Аудит ұғымының, функцияларының анықтамасы... | 127 | 130 | 120с-5к |

> **Note for ОН 4.1–4.6:** No `трад` column in the source. Derive from `п`.

#### 2.6.7 Professional Module 5 — КМ 5 (Қаржы менеджменті)

| # in diploma | Learning Outcome (Kazakh) | `п` col | `трад` col | Hours |
|---|---|---|---|---|
| 41 | КМ 5 *(module header — skip)* | — | — | — |
| 42 | ОН 5.1 Қаржылық басқару функциясының рөлі... | 131 | *derived* | 72с-3к |
| 43 | ОН 5.2 Инвестицияларға тиімді бағалау жүргізу... | 134 | *derived* | 72с-3к |

#### 2.6.8 Professional Practice and Final Attestation

| # in diploma | Item | `п` col | `цэ` col | `трад` col | Hours |
|---|---|---|---|---|---|
| 44 | Кәсіптік практика (Professional practice) | 137 | 139 | 140 | 432с-18к |
| 45 | Қорытынды аттестаттау (Final attestation by qualification 4S04110102) | 141 | 143 | *derived* | — |

> **Note for Final Attestation:** The `трад` column (col 144) contains the average GPA (`Орташа баллы`), not the traditional grade. Derive the traditional grade from the percentage score in col 141.

### 2.7 Traditional Grade Derivation Table

When a `трад` column does not exist, convert the percentage score (`п` column) to a traditional grade using this lookup table:

| Score Range | Traditional Grade (Kazakh) |
|-------------|---------------------------|
| 95–100 | `5 (өте жақсы)` |
| 75–94 | `4 (жақсы)` |
| 50–74 | `3 (қанағат)` |
| 0–49 | `2 (қанағатсыз)` |

> If the `п` value is 0 or empty (student did not take the subject), output empty string for that field.

### 2.8 Average GPA

The average score columns (`Орташа баллы`) start at col 144. These contain calculated average values (п, б, цэ) across all subjects. Extract:

| Field | Column |
|-------|--------|
| Average score (percentage) | 144 |
| Average GPA (0.0–4.0) | 145 |
| Average letter grade | 146 |

---

## 3. Output Data Structure

### 3.1 Per-Student Object

The parser must produce one structured data object per student. The schema is:

```json
{
  "student_info": {
    "sheet": "3D-1",
    "row_number": 7,
    "sequential_number": 1,
    "full_name": "Абайғалиева Нұршат Алматқызы",
    "year_enrollment": 2023,
    "year_graduation": 2026,
    "diploma_serial": "2293467"
  },
  "institution": {
    "name_kz": "Жамбыл инновациялық жоғары колледжі",
    "specialty_code": "04110100",
    "specialty_name_kz": "Есеп және аудит",
    "qualification_code": "4S04110102",
    "qualification_name_kz": "Бухгалтер"
  },
  "general_subjects": [
    {
      "number": 1,
      "name_kz": "Қазақ тілі",
      "hours": 96,
      "credits": 4,
      "score_pct": 81,
      "letter_grade": "B",
      "traditional_grade": "4 (жақсы)"
    }
    // ... subjects 1–13
  ],
  "basic_modules": [
    {
      "number": 14,
      "name_kz": "БМ 01 Дене қасиеттерін дамыту және жетілдіру",
      "hours": 240,
      "credits": 10,
      "score_pct": 75,
      "letter_grade": "B-",
      "traditional_grade": "4 (жақсы)"
    }
    // ... modules 14–17
  ],
  "professional_modules": [
    {
      "module_number": 1,
      "module_name_kz": "КМ 1 Бизнестің мақсаттары мен түрлерін түсіну",
      "learning_outcomes": [
        {
          "number": "1.1",
          "name_kz": "ОН 1.1 Бизнестің мақсаттары мен түрлерін...",
          "hours": 96,
          "credits": 4,
          "score_pct": 86,
          "letter_grade": "B+",
          "traditional_grade": "4 (жақсы)"
        }
        // ... ОН 1.1–1.4
      ]
    }
    // КМ 2, КМ 3, КМ 4, КМ 5
  ],
  "professional_practice": {
    "name_kz": "Кәсіптік практика",
    "designation": "КМ1 ОН1.3; КМ3 ОН3.1, ОН3.2, ОН3.3, ОН3.4; КМ4 ОН 4.1–4.7; КМ5 ОН 5.1, ОН 5.2",
    "hours": 432,
    "credits": 18,
    "score_pct": 86,
    "letter_grade": "B+",
    "traditional_grade": "4 (жақсы)"
  },
  "final_attestation": {
    "qualification": "4S04110102 - Бухгалтер",
    "score_pct": 90,
    "letter_grade": "A-",
    "traditional_grade": "5 (өте жақсы)"
  },
  "averages": {
    "score_pct": 81.5,
    "gpa": 3.0,
    "letter_grade": "B"
  }
}
```

---

## 4. Template Placeholder Mapping

### 4.1 Template: Accounting, Kazakh (`каз__2_.docx`)

The template is a pre-filled `.docx` file (Word document) structured as the official diploma supplement. It contains **text boxes and table cells** with placeholder values that must be replaced for each student.

#### 4.1.1 Identifying Placeholders

The template contains **green-colored text** (hex `#92D050`) marking all values that must be replaced. The parser agent must identify all `<w:t>` elements whose parent `<w:r>` has a color property set to `92D050` and treat these as editable placeholders.

#### 4.1.2 Page 1 — Cover Page Fields

| Green placeholder content | Source field | Notes |
|--------------------------|--------------|-------|
| Diploma serial number (e.g., `2293467`) | `student_info.diploma_serial` | Appears **twice** (two text boxes at top — duplicate for printing front/back) |
| Student full name (e.g., `Абайғалиева Нұршат Алматқызы`) | `student_info.full_name` | Appears twice |
| Enrollment year (e.g., `2023`) | `student_info.year_enrollment` | Appears twice |
| Graduation year (e.g., `2026`) | `student_info.year_graduation` | Appears twice |
| Institution name (`Жамбыл инновациялық жоғары колледжінде`) | `institution.name_kz` | Fixed value, do not change unless different per sheet |
| Specialty code+name (`04110100 – Есеп және аудит`) | `institution.specialty_code + specialty_name_kz` | Fixed value |
| Qualification code+name (`4S04110102 - Бухгалтер`) | `institution.qualification_code + qualification_name_kz` | Fixed value |

#### 4.1.3 Page 1 — General Education Grades Table

The table has 17 rows (subjects 1–13 + basic modules 14–17). Each row contains:

| Column in table | Source field |
|----------------|--------------|
| Subject number (1–17) | Fixed sequential number |
| Subject name | Fixed Kazakh name (from template) |
| Hours | Fixed hours value |
| Letter grade (`цэ`) | `letter_grade` per subject |
| Traditional grade (`трад`) | `traditional_grade` per subject |

The agent replaces the letter grade and traditional grade cells for each subject row.

#### 4.1.4 Page 2 — Professional Modules Grades (КМ 1, КМ 2, КМ 3 partial)

Rows 18–29 cover:
- КМ 1 header row (no grade)
- ОН 1.1, 1.2, 1.3, 1.4
- КМ 2 header row (no grade)
- ОН 2.1, 2.2, 2.3, 2.4
- КМ 3 header row (no grade)
- ОН 3.1

Same column structure: letter grade + traditional grade cells per row.

#### 4.1.5 Page 3 — Professional Modules Grades (КМ 3 continued, КМ 4, КМ 5 partial)

Rows 30–41 cover:
- ОН 3.2, 3.3, 3.4
- КМ 4 header row (no grade)
- ОН 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7
- КМ 5 header row (no grade)
- ОН 5.1

#### 4.1.6 Page 4 — КМ 5 continued + Practice + Attestation + Electives

Rows 42–50 cover:
- ОН 5.2
- Кәсіптік практика (professional practice)
- Қорытынды аттестаттау (final attestation)
- Electives Ф1–Ф5 *(always shown as subject names; grades may be blank if not taken)*

---

## 5. Parsing Logic — Step-by-Step

### Step 1: Load Workbook

```python
import openpyxl
wb = openpyxl.load_workbook(source_file, data_only=True)
```

Use `data_only=True` to get computed formula results rather than formula strings.

### Step 2: Identify Student Sheets

```python
accountant_sheets = [s for s in wb.sheetnames if s.startswith('3D')]
```

### Step 3: Parse Header Once Per Sheet

For a given sheet:
1. Read row 6 to get sub-column type labels.
2. Build a lookup dictionary: `{column_index: 'п'|'б'|'цэ'|'трад'}`.
3. Read enrollment year from cell `(row=6, col=150)` and graduation year from cell `(row=6, col=151)`.

### Step 4: Iterate Student Rows

Student rows begin at row 7. Stop when column B (name) is empty.

```python
for row in ws.iter_rows(min_row=7, values_only=True):
    if not row[1]:  # Column B is name
        break
    student = parse_student_row(row, headers)
```

### Step 5: Extract Grade Values Per Subject

For each subject defined in section 2.6:

```python
def get_grade(row, п_col, трад_col=None):
    п_val = row[п_col - 1]  # zero-indexed
    if трад_col:
        трад_val = row[трад_col - 1]
    else:
        трад_val = derive_traditional_grade(п_val)
    цэ_col_idx = п_col + 2  # letter grade is always п+2
    цэ_val = row[цэ_col_idx - 1]
    return {
        'score_pct': п_val,
        'letter_grade': цэ_val,
        'traditional_grade': трад_val
    }
```

### Step 6: Traditional Grade Derivation

```python
def derive_traditional_grade(score):
    if score is None or score == '' or score == 0:
        return ''
    score = float(score)
    if score >= 95:
        return '5 (өте жақсы)'
    elif score >= 75:
        return '4 (жақсы)'
    elif score >= 50:
        return '3 (қанағат)'
    else:
        return '2 (қанағатсыз)'
```

### Step 7: Extract Metadata

```python
diploma_serial = row[151]  # Col 152, zero-indexed
```

> The diploma serial number is in **col 152** of the student's data row.  
> The enrollment year is in **row 6, col 150** (shared value per sheet).  
> The graduation year is in **row 6, col 151** (shared value per sheet).

### Step 8: Build Output Object

Construct the JSON object as specified in section 3.1.

---

## 6. Template Filling Logic

### 6.1 Input

- One student data object (from section 3.1)
- Template `.docx` file: `каз__2_.docx`

### 6.2 Process

1. Copy the template `.docx` file to a new output path (do not modify the original).
2. Unzip the `.docx` and load `word/document.xml`.
3. Find all `<w:r>` elements with color `92D050` in their `<w:rPr>`.
4. For each green run, identify its parent context (text box, table cell, paragraph) to determine which field it represents.
5. Replace the `<w:t>` text content with the corresponding student data value.
6. Repack the modified XML back into a new `.docx`.

### 6.3 Field Detection Strategy

Because the template document uses text boxes (floating frames), the agent must identify placeholders not only by color but also by their position in the document and their neighboring static text (labels).

Use the following approach:

1. Extract the full text of each green run.
2. Match against known placeholder patterns:

| Green text pattern | Maps to |
|-------------------|---------|
| 7-digit number (e.g., `2293467`) | Diploma serial number |
| Year 4 digits (e.g., `2023`) preceded by label containing "поступления" or "түскен" | Enrollment year |
| Year 4 digits (e.g., `2026`) preceded by label containing "выпуска" or "шыққан" | Graduation year |
| Full name (surname + first name) | Student full name |
| Grade value like `4 (жақсы)` or `5 (өте жақсы)` or `3 (қанағат)` | Traditional grade in grade table |
| Letter grade like `A`, `B+`, `C-` | Letter grade in grade table |
| Percentage score like `86`, `92` | Percentage score in grade table |

3. For grade table cells: identify by row position. The table rows correspond sequentially to the diploma items numbered 1–50 (as per section 4.1.3–4.1.6). Iterate table rows in document order and assign grades from the student data object in order.

### 6.4 Output File Naming

Generate output files using the pattern:

```
{student_full_name}_{sheet_name}_diploma_kaz.docx
```

Example:
```
Абайғалиева_Нұршат_Алматқызы_3D-1_diploma_kaz.docx
```

---

## 7. Edge Cases and Error Handling

### 7.1 Missing Grade Values

If a grade cell in the Excel is `None`, `0`, or empty string:
- Set `score_pct = None`, `letter_grade = ''`, `traditional_grade = ''`
- In the template, replace the placeholder with an empty string (blank cell)

### 7.2 Formula Errors (`#REF!`, `#VALUE!`)

When `data_only=True` is used and a cell contains a formula error:
- The openpyxl value will be `None` or a special error object
- Treat these as missing grade (section 7.1)
- Log a warning: `WARNING: REF error for student {name}, column {col}`

### 7.3 Merged Cells

The Excel file uses merged cells in the header rows. When iterating with `iter_rows(values_only=True)`, merged cell values appear only in the top-left cell of the merge; other cells return `None`. The header parsing must account for this when building the column-to-subject mapping.

### 7.4 Name Encoding

All student names and subject names contain Kazakh characters (including letters like Ғ, Ү, Қ, Ң, Ә, І, Ө). Ensure all file I/O uses `utf-8` encoding.

### 7.5 GPA Values as Strings

Some GPA values in the Excel may be stored as strings with comma decimal separators (e.g., `"3,67"`) rather than floats (`3.67`). Normalize before use:

```python
def normalize_gpa(val):
    if isinstance(val, str):
        return float(val.replace(',', '.'))
    return val
```

### 7.6 Diploma Serial Number

The diploma serial number in col 152 is individual per student. If the cell contains the placeholder text `(каждому отдельно)` instead of a number, this means the serial numbers have not yet been assigned. In this case:
- Leave the serial number blank in the output
- Log a warning: `WARNING: Diploma serial not yet assigned for student {name}`

### 7.7 Students with Incomplete Records

If a student row exists (name is present) but all grade columns are 0 or empty, the student may not have completed all subjects yet. Still generate the output object, but flag:

```json
{ "status": "incomplete", "missing_grades": [1, 2, 5, ...] }
```

---

## 8. Output Formats

### 8.1 JSON (for agent pipeline)

A JSON file per sheet containing an array of student objects:

```
3D-1_parsed_students.json
3D-2_parsed_students.json
```

### 8.2 Filled DOCX Files

One `.docx` per student, generated from the Kazakh accountant template.

### 8.3 Summary Log

A plain-text or CSV log file listing:

| Column | Content |
|--------|---------|
| sheet | Sheet name |
| row | Row number |
| student_name | Full name |
| status | `ok` / `incomplete` / `error` |
| notes | Warning messages |

---

## 9. Subject Name Reference (Kazakh Diploma Template)

The following are the exact Kazakh subject names as they must appear in the diploma supplement (copied verbatim from the template `каз__2_.docx`). The parser must use these exact strings — do not use the Excel header text which may have formatting differences.

### Page 1 — Items 1–17

```
1.  Қазақ тілі
2.  Қазақ әдебиеті
3.  Орыс тілі және әдебиеті
4.  Ағылшын тілі
5.  Қазақстан тарихы
6.  Математика
7.  Информатика
8.  Алғашқы әскери және технологиялық дайындық
9.  Дене тәрбиесі
10. География
11. Биология
12. Физика
13. Графика және жобалау
14. БМ 01 Дене қасиеттерін дамыту және жетілдіру
15. БМ 02 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану
16. БМ 03 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану
17. БМ 04 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану
```

### Page 2 — Items 18–29

```
18. КМ 1 Бизнестің мақсаттары мен түрлерін, негізгі мүдделі тараптармен өзара әрекеттесуін түсіну
19. ОН 1.1 Бизнестің мақсаттары мен түрлерін, олардың негізгі мүдделі тараптармен және сыртқы ортамен өзара әрекеттесуін түсіну
20. ОН 1.2 Көрсеткіштік және логарифмдік функциялар, сызықтық теңдеулер мен матрицалар жүйелері, сызықтық теңсіздіктер және сызықтық бағдарламалау, ықтималдық математикасын білу, бизнес және қаржылық қолдану мәселелерінде ақпаратты талдау және түсіндіру үшін ұғымдарды қолдана білу
21. ОН 1.3 Қаржылық есептіліктің мәні мен мақсатын түсіну, қаржылық ақпараттың сапалық сипаттамаларын анықтау, қаржылық есептілікті дайындау
22. ОН 1.4 Маркетингтің негізгі тұжырымдамаларды түсіну, маркетингтік ортаны зерттеу, тұтынушылар мен ұйымның сатып алу тәртібін түсіну, нарықтарды сегменттеу және өнімдерді орналастыру, жаңа өнімдерді әзірлеу үшін қолданылатын құралдар мен әдістерді білу
23. КМ 2 Кәсіптік салада тілдік дағдыларды қолдану
24. ОН 2.1 Академиялық деңгейде Ағылшын тілінің оқылым, айтылым және жазылым дағдыларын еркін меңгеру
25. ОН 2.2 Кәсіби салада Ағылшын тілінің айтылым және жазылым дағдыларын B2 деңгейінде еркін меңгеру
26. ОН 2.3 Іскерлік мақсатта қазақ тілін қолдану
27. ОН 2.4 Іскерлік мақсатта түрік тілін қолдану
28. КМ 3 Бухгалтерлік (қаржылық) есептілікті жасауға қатысу
29. ОН 3.1 Басқару ақпаратының сипатын, мақсатын түсіну, шығындарды есепке алу, жоспарлау, бизнестің тиімділігін бақылау
```

### Page 3 — Items 30–41

```
30. ОН 3.2 Еңбек қатынастарына қатысты заңды түсіну, компаниялардың қалай басқарылатындығын және реттелетінін сипаттау және түсіну
31. ОН 3.3 Іскерлік шешім қабылдау процесін қолдайтын жалпы математикалық құралдарды қолдану, аналитикалық әдістерді әртүрлі бизнес қолданбаларында қолдану
32. ОН 3.4 Негізгі экономикалық принциптерді, макроэкономикалық мәселелерді және көрсеткіштерді есептеуді білу, фискалдық және ақша-несие саясатының макроэкономикаға әсер ету механизмдерді талдау
33. КМ 4 Ұйымның және оның бөлімшелерінің шаруашылық-қаржылық қызметін кешенді талдауға қатысу
34. ОН 4.1 Инвестициялар мен қаржыландыруды бағалаудың баламалы тәсілдерін салыстыру, қаржы саласындағы проблемаларды шешудің әртүрлі тәсілдерінің орындылығын бағалау
35. ОН 4.2 Ұйымдарға өнімділікті басқару және өлшеу үшін қажет ақпаратты, технологиялық жүйелерді анықтау, шығындарды есепке алу және басқару есебі әдістерін қолдану
36. ОН 4.3 Салық жүйесінің жұмыс істеуі мен көлемін және оны басқаруды түсіну
37. ОН 4.4 ХҚЕС стандарттарына сәйкес операцияларды есепке алу, Қаржылық есептерді талдау және түсіну
38. ОН 4.5 Бизнес статистикадағы негізгі түсініктерді, деректер материалдарын жинау, қорытындылау және талдау әдістерін білу
39. ОН 4.6 Бухгалтерлік есептің ақпараттық жүйелері
40. ОН 4.7 Аудит ұғымының, функцияларының, корпоративтік басқарудың, оның ішінде этика мен кәсіби мінез-құлықтың анықтамасы, Халықаралық аудит стандарттарын (АХС) қолдану
41. КМ 5 Қаржы менеджментіне экономикалық ортаның әсерін бағалау
```

### Page 4 — Items 42–50

```
42. ОН 5.1 Қаржылық басқару функциясының рөлі мен мақсатын түсіну, Қаржы менеджментіне экономикалық ортаның әсерін бағалау
43. ОН 5.2 Инвестицияларға тиімді бағалау жүргізу, Бизнесті қаржыландырудың балама көздерін анықтау және бағалау
44. Кәсіптік практика КМ3. ОН3.2, ОН3.3; КМ4. ОН4.3; КМ5. ОН5.2, ОН5.3; КМ7. ОН7.1, ОН7.2, ОН7.3; КМ8. ОН8.1, ОН8.2, ОН8.3; КМ9. ОН9.1, ОН9.2, ОН9.3.
45. Қорытынды аттестаттау :
46. Ф1 Факультативтік ағылшын тілі
47. Ф2 Факультативтік түрік тілі
48. Ф3 Факультативтік Бизнес және бухгалтерлік есептегі жағдайлар (Cases in Business and Accounting)
49. Ф4 Факультативтік Бизнес деректерін талдау (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))
50. Ф5 Факультативтік кәсіпкерлік қызмет негіздері (Enterpreneurship)
```

---

## 10. Credit Hours Reference

The following credit hours are fixed per subject for all students in the accountant program (extracted from row 5 of the source file):

| Item # | Hours | Credits |
|--------|-------|---------|
| 1 Қазақ тілі | 96 | 4 |
| 2 Қазақ әдебиеті | 96 | 4 |
| 3 Орыс тілі және әдебиеті | 96 | 4 |
| 4 Ағылшын тілі | 216 | 9 |
| 5 Қазақстан тарихы | 96 | 4 |
| 6 Математика | 120 | 5 |
| 7 Информатика | 48 | 2 |
| 8 Алғашқы әскери дайындық | 96 | 4 |
| 9 Дене тәрбиесі | 120 | 5 |
| 10 География | 120 | 5 |
| 11 Биология | 120 | 5 |
| 12 Физика | 72 | 3 |
| 13 Графика және жобалау | 72 | 3 |
| 14 БМ 01 | 240 | 10 |
| 15 БМ 02 | 72 | 3 |
| 16 БМ 03 | 72 | 3 |
| 17 БМ 04 | 24 | 1 |
| 19 ОН 1.1 | 96 | 4 |
| 20 ОН 1.2 | 96 | 4 |
| 21 ОН 1.3 | 120 | 5 |
| 22 ОН 1.4 | 72 | 3 |
| 24 ОН 2.1 | 168 | 7 |
| 25 ОН 2.2 | 48 | 2 |
| 26 ОН 2.3 | 48 | 2 |
| 27 ОН 2.4 | 72 | 3 |
| 29 ОН 3.1 | 120 | 5 |
| 30 ОН 3.2 | 48 | 2 |
| 31 ОН 3.3 | 72 | 3 |
| 32 ОН 3.4 | 120 | 5 |
| 34 ОН 4.1 | 72 | 3 |
| 35 ОН 4.2 | 120 | 5 |
| 36 ОН 4.3 | 72 | 3 |
| 37 ОН 4.4 | 72 | 3 |
| 38 ОН 4.5 | 72 | 3 |
| 39 ОН 4.6 | 72 | 3 |
| 40 ОН 4.7 | 120 | 5 |
| 42 ОН 5.1 | 72 | 3 |
| 43 ОН 5.2 | 72 | 3 |
| 44 Кәсіптік практика | 432 | 18 |

---

## 11. Institution Constants (Hardcoded for Accountant Program)

These values are **fixed** and do not come from the Excel file. The agent should embed them as constants:

```python
INSTITUTION_KZ = "Жамбыл инновациялық жоғары колледжі"
SPECIALTY_CODE = "04110100"
SPECIALTY_NAME_KZ = "Есеп және аудит"
QUALIFICATION_CODE = "4S04110102"
QUALIFICATION_NAME_KZ = "Бухгалтер"
```

---

## 12. Dependencies and Environment

- **Python 3.9+**
- `openpyxl` — Excel parsing (`pip install openpyxl`)
- `python-docx` — DOCX reading/writing (`pip install python-docx`)
- Standard library: `json`, `os`, `re`, `logging`

---

## 13. Scope Exclusions for This TZ

The following are **out of scope** for this specification:
- Russian-language accountant template (separate TZ)
- IT specialty templates (separate TZ)
- Grade calculation or validation logic (parser reads values as-is)
- PDF generation
- Elective (Факультатив) grades — these are optional and typically blank
- Red diploma (қызыл диплом) detection logic — referenced in the source filename but not required for basic parsing