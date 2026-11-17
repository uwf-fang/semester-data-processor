# Semester Data Processor (`semester-data-processor`)

[![Main CI](https://github.com/uwf-fang/semester-data-processor/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/uwf-fang/semester-data-processor/actions/workflows/ci.yml?query=branch%3Amain+event%3Apush)
[![Release Checks](https://github.com/uwf-fang/semester-data-processor/actions/workflows/release-checks.yml/badge.svg)](https://github.com/uwf-fang/semester-data-processor/actions/workflows/release-checks.yml)
[![PyPI version](https://img.shields.io/pypi/v/semester-data-processor)](https://pypi.org/project/semester-data-processor/)

`semester-data-processor` is an automated Python package designed for university faculty, course coordinators, and academic departments to streamline course data management, output dataset documentation, and accreditation reporting. 

It automates:
1. **Roster Consolidation**: Merging multi-section section rosters, extracting 5-digit section CRNs, and normalizing student major abbreviations (`CS`, `SD`, `Cyber`, `CE`, `EE`).
2. **Gradebook Merging**: Merging Canvas gradebook exports while preserving **100% of Canvas grade columns intact** and mapping roster attributes.
3. **Dataset Documentation & Processing Audit Log**: Automatically generating and maintaining dataset `README.md` documentation in each course directory with raw input file injection timestamps and pipeline run history.
4. **Accreditation Assessment Reporting**: Evaluating targeted assignments against ABET / SACSCOC / general accreditation learning outcomes and automatically generating standardized Academic Year Excel workbooks (`[Course]-[Scope]-Individual Data-[LO].xlsx`) with FERPA-compliant `S1, S2, ...` anonymous IDs and live formulas, summary CSVs (`assessment_summary.csv`), per-student Pivot-Table-ready CSV datasets (`assessment_results.csv`), and Markdown reports (`assessment_report.md`).

---

## Flexible Reporting Options: Standardized Academic Year Excel vs. Custom Pivot Table

Faculty and department coordinators can choose between two flexible reporting methods:

- **Method A: Standardized Academic Year Excel Report (`[Course]-[Scope]-Individual Data-[LO].xlsx`) — Instant & Department-Ready**
  Automatically generates the official annual reporting workbook formatted for university accreditation. Includes section instructor attribution (from `<term>-sections.txt`), anonymous student IDs (`S1, S2, ...`), Fall and Spring 2-semester comparison blocks, and live dynamic Excel formulas (`COUNTIF`, `SUM`, pass percentages, and MET/NOT MET status).

- **Method B: Custom Pivot Table from CSV (`assessment_results.csv`) — Manual Flexibility**
  If you prefer custom spreadsheet analysis in **Microsoft Excel** or **Google Sheets**, open `assessment_results.csv` (which retains real student identifiers and section attributes) and create custom interactive Pivot Tables using the built-in numeric helper fields (`Pass Count` and `Total Student Count`).

👉 **[Pivot Table Reporting Guide](PIVOT_TABLE_GUIDE.md)** for step-by-step Pivot Table instructions.

---

## Installation & Setup Options

First, clone the repository to your local environment:
```bash
git clone https://github.com/uwf-csd/semester-data-processor.git
cd semester-data-processor
```

You can then install and run `semester-data-processor` locally using either standard **`pip`** or **`uv`**:

### Option 1: Standard `pip` Local Installation
```bash
# Install locally in editable mode (recommended)
pip install -e .

# Or install from built Wheel (.whl) package
pip install dist/semester_data_processor-0.1.0-py3-none-any.whl
```

### Option 2: Local Installation with `uv`
```bash
# Install from local directory using uv
uv pip install .
```

### Option 3: Instant Execution with `uv run` (Zero Setup Needed)
If you use `uv`, you can execute the tool directly from the local directory without running a manual install:
```bash
uv run semester-processor run-all copxxxx-26spring
```

---

## Step-by-Step User Guide: How to Prepare Your Files & Run the Tool

### Step 1: Create a Dedicated Course Academic Year Directory
Create a directory named after your course and academic year (e.g. `25-26`) or semester offering. 

*Example Directory Name*: `cop4020-25-26` (COP 4020 Course Title - Academic Year 2025-2026)

```bash
mkdir cop4020-25-26
```

---

### Step 2: Create a `raw/` Subdirectory & Copy Input Files
Inside your course directory (`cop4020-25-26`), create a subdirectory named **`raw/`**. 

Copy all raw input files exported from Banner/MyUWF and Canvas into `cop4020-25-26/raw/`:
- **Raw Roster CSVs**: Official university section rosters (e.g. `CourseRoster_202601_16197.csv`, `CourseRoster_202601_17089.csv`). The 6-digit term code (`202601`) and section CRN (`16197`, `17089`) are automatically extracted from the filename.
- **Section Instructor Text Files**: Plain text mapping file named `<term>-sections.txt` (e.g., `202601-sections.txt` or `202508-sections.txt`). Lists instructors and their section CRNs line-by-line:
  ```
  Fang 16197 17089
  Zhang 18451 18626
  ```
- **Raw Canvas Gradebook CSVs**: All instructors’ gradebook exports from Canvas (e.g. `2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv`).

> ⚠️ **Important Data Rule**: Never modify or rename raw input files inside `raw/`. The tool reads raw data pristine and generates processed output files directly in the root of the course directory.

---

### Multiple instructors and overlapping exports

Put every instructor's export in the same `raw/` directory. Every CSV with `grade` in its filename (case-insensitive) is read. The tool unions assignment columns, aligns identical titles after removing Canvas numeric ID suffixes, and merges students by `SIS User ID`. Roster-only students remain in the output.

Each student's percentage uses points possible from that student's source export, so `40/50` and `80/100` both produce 80%. The Points Possible row is detected with or without Manual Posting metadata. Identical overlapping student records are counted once; complementary blank cells are filled. Conflicting nonblank values or points possible for an overlapping student/assignment raise an error identifying the files and column. Duplicate cleaned titles within one export also raise an error. Resolve conflicting source selections in a separate input directory while preserving originals; no export is silently preferred.

Use matching assignment titles across instructors for automatic alignment. `assignment_column` can select an exact cleaned title; an ambiguous query with multiple scored matches raises an error. Missing or invalid points possible for numeric scores raises an error rather than assuming 100 points. Re-run `run-all` after adding rosters; `extract-assessments` refreshes gradebooks whenever raw exports are available.

### Step 3: Create the Assessment Plan (`assessment_plan.json`)
Create a file named **`assessment_plan.json`** directly inside `cop4020-25-26/` (or place it in `cop4020-25-26/raw/`).

This JSON file defines the academic year evaluation rules, passing thresholds, target pass rates, and evaluated majors. It supports **single or multiple learning outcomes / performance indicators** across any accreditation framework:

```json
{
  "course_name": "COP 4020 - Programming Languages",
  "academic_year": "25-26",
  "target_majors": ["Computer Science"],
  "assessments": [
    {
      "accreditation_type": "ABET",
      "code": "LO6.PI3",
      "assessment_name": "Project 1",
      "assignment_column": "Project 1",
      "target_majors": ["Computer Science"],
      "passing_operator": ">=",
      "passing_threshold_percent": 70.0,
      "target_pass_rate_percent": 70.0,
      "notes": "Design and implement a parser for a given grammar."
    }
  ]
}
```

The example evaluates Project 1 for Computer Science across the academic year. Each outcome retains its own population, denominator, target, and section totals. Repeat an assignment in separate items to report different majors independently. No pooled pass rate is calculated across unrelated outcomes.

Detailed CSV rows retain all active students for audit. `Target Major` identifies the item's scope and `Included in Assessment` indicates whether a row contributes to its summary. Sum the count columns to reproduce the summary; do not count all detail rows. A target population with no active students is reported as `NO DATA`. Missing scores among included active students continue to count as failures.

#### Field Explanations:
- `course_name`: Official title of the course.
- `academic_year`: Two-semester academic year (e.g. `"25-26"` or `"2025-2026"`). Maps to Fall (`202508`) and Spring (`202601`).
- `target_majors`: Set this list at the top level or inside each assessment item to select programs (full names or abbreviations such as `CS` and `Cyber`, case-insensitive). `[]` selects all active majors.
- `accreditation_type`: `ABET`, `SACSCOC`, or any accreditation framework name.
- `code`: Learning outcome / performance indicator ID (e.g. `"LO6.PI3"`).
- `assessment_name`: Name/title of the assessed assignment (e.g. `"Project 1"`).
- `assignment_column`: Assignment query matched against Canvas gradebook columns via fuzzy matching.
- `passing_operator`: Operator for passing rule (`>=`, `>`, `<=`, `<`, `==`).
- `passing_threshold_percent`: Minimum percentage score to pass (e.g. `70.0`).
- `target_pass_rate_percent`: Department target pass rate benchmark (e.g. `70.0`).
- `notes`: Description or outcome explanation copied into the Excel report.

---

### Step 4: Run the Semester Processor Tool

#### Method A: Run with Standard `pip` Installed Package
```bash
semester-processor run-all copxxxx-26spring
```

#### Method B: Run with `uv run` (Zero Installation Required)
```bash
uv run semester-processor run-all copxxxx-26spring
```

---

## Demonstration Walkthrough: Before vs. After Directory Structure

### BEFORE Running the Tool
Your course directory only contains raw input files inside `raw/` and your `assessment_plan.json`:

```
cop4020-25-26/
├── raw/
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   ├── 202601-sections.txt
│   └── 2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv
└── assessment_plan.json
```

---

### AFTER Running the Tool
The tool processes the course directory, preserving raw files intact while generating processed output files directly in the root of `cop4020-25-26/`:

```
cop4020-25-26/
├── raw/                                           # RAW DATA PRESERVED UNTOUCHED
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   ├── 202601-sections.txt
│   └── 2026-09-02T1237_Grades-Programming_Languages_-_16197_COP4020_202601.csv
├── assessment_plan.json                            # Accreditation configuration plan
├── combined_roster.csv                             # GENERATED: Merged multi-section roster with Anonymous IDs
├── merged_gradebook.csv                            # GENERATED: Consolidated gradebook + roster attributes
├── assessment_results.csv                          # GENERATED: Detailed per-student assessment CSV (Pivot Table ready)
├── assessment_summary.csv                          # GENERATED: Aggregated outcome summary CSV (Pass metrics only)
├── COP4020-CS-ONLY-Individual Data-LO6PI3.xlsx      # GENERATED: Standardized Academic Year Excel workbook
├── assessment_report.md                            # GENERATED: Formatted Markdown report
└── README.md                                       # GENERATED: Dataset overview & Processing Audit Log
```

---

## Generated Output Files Explained

1. **`combined_roster.csv`**:
   - Merges all section rosters found in `raw/`.
   - Extracts 5-digit section CRNs (`16197`, `17089`).
   - Assigns section-relative anonymous IDs (`S1, S2, ...`) for active students (FERPA compliant).
   - Maps instructor names from `<term>-sections.txt`.
   - Normalizes student programs (`Computer Science` $\rightarrow$ `CS`, `Software Design` $\rightarrow$ `SD`, `Cybersecurity` $\rightarrow$ `Cyber`, etc.).
   - Computes student status (`Active` vs. `Withdrawn`).

2. **`merged_gradebook.csv`**:
   - Combines all Canvas gradebook CSVs with `combined_roster.csv`.
   - Preserves **100% of Canvas grade columns intact**.
   - Carries forward `Anonymous ID`, `Instructor`, `Section Number`, `Major`, `Status`, `Level`, and `Cohort` columns.
   - Dynamic row parser automatically detects `Points Possible` rows and handles presence/absence of `"Manual Posting"` rows without skipping student data.

3. **`assessment_results.csv`**:
   - Detailed per-student dataset listing every student's Program, Name, Anonymous ID, Instructor, Section ID, Assessed Assignment, Assessed Score, and Pass/Fail Status.
   - Includes helper count columns (`Pass Count` = 1/0 and `Total Student Count` = 1 for included students; both are zero outside the item’s target majors) for instant **Excel / Google Sheets Pivot Tables**.
   - See the **[Pivot Table Reporting Guide](PIVOT_TABLE_GUIDE.md)** for step-by-step instructions.

4. **`assessment_summary.csv`**:
   - Aggregated CSV table saving summary data similar to the Pivot Table output.
   - Summarizes `Passed Count`, `Total Evaluated`, `Pass Rate (%)`, and `Outcome Status` (`MET` / `NOT MET`) per section and total across all sections.
   - Excludes all `Fail` columns for clean, focused outcome reporting.

5. **`[Course]-[Scope]-Individual Data-[LO].xlsx`**:
   - Standardized Academic Year Assessment Excel report matching official institutional templates.
   - Anonymizes students using `S1, S2, ...` per section.
   - Side-by-side section tables for Fall (Rows 3–28) and Spring (Rows 32–45).
   - Live dynamic Excel formulas for section pass totals (`=COUNTIF(...)`), valid submissions (`=SUM(COUNTIF(...), COUNTIF(...))`), Semester Totals, Grand Totals, Pass Percentage (`=IF(D>0, ROUND(C/D*100, 0), 0)`), and Met PI check (`=IF(C>=70, "Y", "N")`).
   - Learning outcome description and evaluation criteria at the bottom.

6. **`assessment_report.md`**:
   - Professional Markdown assessment report containing:
     - Executive summary, Academic Year, active term offering, & total roster enrollment breakdown by major and section.
     - Outcome pass/fail ratios and percentages per section and total sum across all sections.
     - Complete student outcome result tables.

7. **`README.md`**:
   - Dataset documentation detailing data products and output usage.
   - Features a **Processing Log & Updates** audit trail section recording execution timestamps, raw file modification/injection times, and run statistics.

---

## Command Line Interface (CLI) Commands

With `pip` (or `uv run`):
```bash
# Run full pipeline
semester-processor run-all copxxxx-26spring

# Run individual pipeline steps separately
semester-processor process-rosters copxxxx-26spring
semester-processor process-gradebook copxxxx-26spring
semester-processor extract-assessments copxxxx-26spring

# Generate/update dataset output README & processing log
semester-processor generate-readme copxxxx-26spring
```

---

## Python API Usage

You can also import and execute the workflow inside Python scripts or Jupyter Notebooks:

```python
from semester_data_processor import process_rosters, process_gradebook, extract_assessments, generate_output_readme

course_dir = "copxxxx-26spring"

# Execute workflow steps programmatically
process_rosters(course_dir)
process_gradebook(course_dir)
extract_assessments(course_dir)
generate_output_readme(course_dir, action_name="Full Pipeline Run")
```

---

## Running Unit Tests

GitHub Actions builds the wheel and source distribution, checks the package metadata, and runs the test suite against the installed wheel on Ubuntu with Python 3.10 through 3.15. Main CI runs on pushes and pull requests targeting `main`. Release Checks runs the same checks at the release tag when a GitHub release is published. The PyPI badge shows the package version available on PyPI.

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

---

## License

This software is licensed under the standard [MIT License](LICENSE).

