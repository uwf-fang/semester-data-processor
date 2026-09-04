# Assessment Data Processor (`assessment-data-processor`)

`assessment-data-processor` is an automated Python package designed for university faculty, course coordinators, and academic departments to streamline course data management and accreditation reporting. 

It automates:
1. **Roster Consolidation**: Merging multi-section section rosters, extracting 5-digit section CRNs, and normalizing student major abbreviations (`CS`, `SD`, `Cyber`, `CE`, `EE`).
2. **Gradebook Merging**: Merging Canvas gradebook exports while preserving **100% of Canvas grade columns intact** and mapping roster attributes.
3. **Accreditation Assessment Reporting**: Evaluating targeted assignments against ABET / SACSCOC learning outcomes and automatically generating direct formatted Excel workbooks (`assessment_report.xlsx`), summary CSVs (`assessment_summary.csv`), per-student Pivot-Table-ready CSV datasets (`assessment_results.csv`), and Markdown reports (`assessment_report.md`).

---

## Flexible Reporting Options: Direct Excel vs. Manual Pivot Table

Faculty and department coordinators can choose between two flexible reporting methods:

- **Method A: Direct Excel Report (`assessment_report.xlsx`) — Instant & Pre-Formatted**
  Simply open `assessment_report.xlsx` generated automatically by the tool. It includes styled Executive Summary cards, outcome totals, and per-student tables ready for instant department review out-of-the-box.

- **Method B: Custom Pivot Table from CSV (`assessment_results.csv`) — Manual Flexibility**
  If you prefer custom spreadsheet analysis in **Microsoft Excel** or **Google Sheets**, open `assessment_results.csv` and create custom interactive Pivot Tables using the built-in numeric helper fields (`Pass Count` and `Total Student Count`).

👉 **[Pivot Table Reporting Guide](PIVOT_TABLE_GUIDE.md)** for step-by-step Pivot Table instructions.

---

## Installation & Setup Options

First, clone the repository to your local environment:
```bash
git clone https://github.com/uwf-csd/assessment-data-processor.git
cd assessment-data-processor
```

You can then install and run `assessment-data-processor` locally using either standard **`pip`** or **`uv`**:

### Option 1: Standard `pip` Local Installation
```bash
# Install locally in editable mode (recommended)
pip install -e .

# Or install from built Wheel (.whl) package
pip install dist/assessment_data_processor-0.1.0-py3-none-any.whl
```

### Option 2: Local Installation with `uv`
```bash
# Install from local directory using uv
uv pip install .
```

### Option 3: Instant Execution with `uv run` (Zero Setup Needed)
If you use `uv`, you can execute the tool directly from the local directory without running a manual install:
```bash
uv run assessment-processor run-all cop4020-26spring
```

---

## Step-by-Step User Guide: How to Prepare Your Files & Run the Tool

### Step 1: Create a Dedicated Course Semester Directory
Create a directory named after your course and semester in your workspace. 

*Example Directory Name*: `cop4020-26spring` (COP 4020 Programming Languages - Spring 2026)

```bash
mkdir cop4020-26spring
```

---

### Step 2: Create a `raw/` Subdirectory & Copy Input Files
Inside your course directory (`cop4020-26spring`), create a subdirectory named **`raw/`**. 

Copy all raw input files exported from Banner/MyUWF and Canvas into `cop4020-26spring/raw/`:
- **Raw Roster CSVs**: Official university section rosters (e.g. `CourseRoster_202601_16197.csv`, `CourseRoster_202601_17089.csv`). The section CRN (`16197`, `17089`) is automatically extracted from the filename.
- **Raw Canvas Gradebook CSV**: Gradebook export from Canvas (e.g. `2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv`).

> ⚠️ **Important Data Rule**: Never modify or rename raw input files inside `raw/`. The tool reads raw data pristine and generates processed output files directly in the root of the course directory.

---

### Step 3: Create the Assessment Plan (`assessment_plan.json`)
Create a file named **`assessment_plan.json`** directly inside `cop4020-26spring/` (or place it in `cop4020-26spring/raw/`).

This JSON file defines which assignment to assess, passing thresholds, target pass rates, and evaluated majors. It supports **single or multiple learning outcomes / performance indicators**:

```json
{
  "course_name": "COP 4020 - Programming Languages",
  "semester": "Spring 2026",
  "target_majors": ["Computer Science"],
  "assessments": [
    {
      "accreditation_type": "ABET",
      "code": "LO.PI 1.1",
      "assessment_name": "Project 1",
      "passing_operator": ">=",
      "passing_threshold_percent": 70.0,
      "target_pass_rate_percent": 75.0,
      "notes": "ABET outcome assessment based on Project 1 for CS majors."
    },
    {
      "accreditation_type": "ABET",
      "code": "LO.PI 1.2",
      "assessment_name": "Project 2",
      "passing_operator": ">=",
      "passing_threshold_percent": 70.0,
      "target_pass_rate_percent": 75.0,
      "notes": "ABET outcome assessment based on Project 2 for CS majors."
    }
  ]
}
```

#### Field Explanations:
- `course_name`: Official title of the course.
- `semester`: Term and year (e.g., `Spring 2026`).
- `target_majors`: List of programs to evaluate (e.g. `["Computer Science"]`). Leave empty `[]` to evaluate all enrolled students.
- `accreditation_type`: `ABET` or `SACSCOC`.
- `code`: Learning outcome / performance indicator ID (e.g. `"LO.PI 1.1"`).
- `assessment_name`: Name/title of the assessed assignment (e.g. `"Project 1"`). Merges assignment query and assessment name into a single clean attribute. Multi-stage fuzzy matching aligns headers automatically without third-party dependencies!
- `passing_operator`: Operator for passing rule (`>=`, `>`, `<=`, `<`, `==`).
- `passing_threshold_percent`: Minimum percentage score to pass (e.g. `70.0`).
- `target_pass_rate_percent`: Department target pass rate benchmark (e.g. `75.0`).

---

### Step 4: Run the Assessment Processor Tool

#### Method A: Run with Standard `pip` Installed Package
```bash
assessment-processor run-all cop4020-26spring
```

#### Method B: Run with `uv run` (Zero Installation Required)
```bash
uv run assessment-processor run-all cop4020-26spring
```

---

## Demonstration Walkthrough: Before vs. After Directory Structure

### BEFORE Running the Tool
Your course directory only contains raw input files inside `raw/` and your `assessment_plan.json`:

```
cop4020-26spring/
├── raw/
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   └── 2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv
└── assessment_plan.json
```

---

### AFTER Running the Tool
The tool reorganizes the course directory, preserving raw files intact while generating processed output files directly in the root of `cop4020-26spring/`:

```
cop4020-26spring/
├── raw/                                           # RAW DATA PRESERVED UNTOUCHED
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   └── 2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv
├── assessment_plan.json                            # Accreditation configuration plan
├── combined_roster.csv                             # GENERATED: Merged multi-section roster
├── merged_gradebook.csv                            # GENERATED: Consolidated gradebook + roster attributes
├── assessment_results.csv                          # GENERATED: Detailed per-student assessment CSV (Pivot Table ready)
├── assessment_summary.csv                          # GENERATED: Aggregated outcome summary CSV (Pass metrics only)
├── assessment_report.xlsx                          # GENERATED: Formatted multi-tab Excel workbook
└── assessment_report.md                            # GENERATED: Formatted Markdown report
```

---

## Generated Output Files Explained

1. **`combined_roster.csv`**:
   - Merges all section rosters found in `raw/`.
   - Extracts 5-digit section CRNs (`16197`, `17089`).
   - Normalizes student programs (`Computer Science` $\rightarrow$ `CS`, `Software Design` $\rightarrow$ `SD`, `Cybersecurity` $\rightarrow$ `Cyber`, etc.).
   - Computes student status (`Active` vs. `Withdrawn`).

2. **`merged_gradebook.csv`**:
   - Merges Canvas gradebook CSV with `combined_roster.csv`.
   - Preserves **100% of Canvas grade columns intact**.
   - Attaches `Section Number`, `Major`, `Status`, `Level`, and `Cohort` columns to each student row.
   - Dynamic row parser automatically detects `Points Possible` rows and handles presence/absence of `"Manual Posting"` rows without skipping student data.

3. **`assessment_results.csv`**:
   - Detailed per-student dataset listing every student's Program, Name, Section ID, Assessed Assignment, Assessed Score, and Pass/Fail Status.
   - Includes helper count columns (`Pass Count` = 1/0 and `Total Student Count` = 1) for instant **Excel / Google Sheets Pivot Tables**.
   - See the **[Pivot Table Reporting Guide](PIVOT_TABLE_GUIDE.md)** for step-by-step instructions.

4. **`assessment_summary.csv`**:
   - Aggregated CSV table saving summary data similar to the Pivot Table output.
   - Summarizes `Passed Count`, `Total Evaluated`, `Pass Rate (%)`, and `Outcome Status` (`MET` / `NOT MET`) per section and total across all sections.
   - Excludes all `Fail` columns for clean, focused outcome reporting.

5. **`assessment_report.xlsx`**:
   - Multi-tab Excel workbook containing:
     - **Tab 1 (`Assessment Summary & Plan`)**: Assessment plan configuration details, followed by a simplified assessment outcome summary (Code, Assessment Name, Assessed Assignment, Passed Count, Total Evaluated, Pass Ratio, Pass Rate %, MET/NOT MET status).
     - **Tab 2 (`Student Assessment Results`)**: Complete per-student evaluation table with custom header styling and auto-adjusted column widths.

6. **`assessment_report.md`**:
   - Professional Markdown assessment report containing:
     - Executive summary & total roster enrollment breakdown by major and section.
     - Outcome pass/fail ratios and percentages per section and total sum across all sections.
     - Complete student outcome result tables.

---

## Command Line Interface (CLI) Commands

With `pip` (or `uv run`):
```bash
# Run full pipeline
assessment-processor run-all cop4020-26spring

# Run individual pipeline steps separately
assessment-processor process-rosters cop4020-26spring
assessment-processor process-gradebook cop4020-26spring
assessment-processor extract-assessments cop4020-26spring
```

---

## Python API Usage

You can also import and execute the workflow inside Python scripts or Jupyter Notebooks:

```python
from assessment_data_processor import process_rosters, process_gradebook, extract_assessments

course_dir = "cop4020-26spring"

# Execute workflow steps programmatically
process_rosters(course_dir)
process_gradebook(course_dir)
extract_assessments(course_dir)
```

---

## Running Unit Tests

Run the automated verification test suite:

```bash
# Using standard unittest
python3 -m unittest discover tests

# Or using uv
uv run python -m unittest discover tests
```

---

## Building Package Distributions

To build fresh wheel (`.whl`) and source distribution (`.tar.gz`) packages:

```bash
uv build
```
Build artifacts will be saved in `dist/`.
