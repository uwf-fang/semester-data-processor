# Semester Data Processor (`semester-data-processor`)

`semester-data-processor` is an automated Python package designed for university faculty, course coordinators, and academic departments to streamline course data management, output dataset documentation, and accreditation reporting. 

It automates:
1. **Roster Consolidation**: Merging multi-section section rosters, extracting 5-digit section CRNs, and normalizing student major abbreviations (`CS`, `SD`, `Cyber`, `CE`, `EE`).
2. **Gradebook Merging**: Merging Canvas gradebook exports while preserving **100% of Canvas grade columns intact** and mapping roster attributes.
3. **Dataset Documentation & Processing Audit Log**: Automatically generating and maintaining dataset `README.md` documentation in each course directory with raw input file injection timestamps and pipeline run history.
4. **Accreditation Assessment Reporting**: Evaluating targeted assignments against ABET / SACSCOC learning outcomes and automatically generating direct formatted Excel workbooks (`assessment_report.xlsx`), summary CSVs (`assessment_summary.csv`), per-student Pivot-Table-ready CSV datasets (`assessment_results.csv`), and Markdown reports (`assessment_report.md`).

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

### Step 1: Create a Dedicated Course Semester Directory
Create a directory named after your course and semester in your workspace. 

*Example Directory Name*: `copxxxx-26spring` (COP XXXX Course Title - Spring 2026)

```bash
mkdir copxxxx-26spring
```

---

### Step 2: Create a `raw/` Subdirectory & Copy Input Files
Inside your course directory (`copxxxx-26spring`), create a subdirectory named **`raw/`**. 

Copy all raw input files exported from Banner/MyUWF and Canvas into `copxxxx-26spring/raw/`:
- **Raw Roster CSVs**: Official university section rosters (e.g. `CourseRoster_202601_16197.csv`, `CourseRoster_202601_17089.csv`). The section CRN (`16197`, `17089`) is automatically extracted from the filename.
- **Raw Canvas Gradebook CSVs**: All instructors’ gradebook exports from Canvas (e.g. `2026-09-02T1237_Grades-Course_Title_-_16197_COPXXXX_202601.csv`).

> ⚠️ **Important Data Rule**: Never modify or rename raw input files inside `raw/`. The tool reads raw data pristine and generates processed output files directly in the root of the course directory.

---

### Multiple instructors and overlapping exports

Put every instructor's export in the same `raw/` directory. Every CSV with `grade` in its filename (case-insensitive) is read. The tool unions assignment columns, aligns identical titles after removing Canvas numeric ID suffixes, and merges students by `SIS User ID`. Roster-only students remain in the output.

Each student's percentage uses points possible from that student's source export, so `40/50` and `80/100` both produce 80%. The Points Possible row is detected with or without Manual Posting metadata. Identical overlapping student records are counted once; complementary blank cells are filled. Conflicting nonblank values or points possible for an overlapping student/assignment raise an error identifying the files and column. Duplicate cleaned titles within one export also raise an error. Resolve conflicting source selections in a separate input directory while preserving originals; no export is silently preferred.

Use matching assignment titles across instructors for automatic alignment. `assignment_column` can select an exact cleaned title; an ambiguous query with multiple scored matches raises an error. Missing or invalid points possible for numeric scores raises an error rather than assuming 100 points. Re-run `run-all` after adding rosters; `extract-assessments` refreshes gradebooks whenever raw exports are available.

### Step 3: Create the Assessment Plan (`assessment_plan.json`)
Create a file named **`assessment_plan.json`** directly inside `copxxxx-26spring/` (or place it in `copxxxx-26spring/raw/`).

This JSON file defines which assignment to assess, passing thresholds, target pass rates, and evaluated majors. It supports **single or multiple learning outcomes / performance indicators**:

```json
{
  "course_name": "COP XXXX - Course Title",
  "semester": "Spring 2026",
  "assessments": [
    {
      "accreditation_type": "ABET",
      "code": "LO.PI 1.1",
      "assessment_name": "Project 1",
      "target_majors": ["Computer Science"],
      "passing_operator": ">=",
      "passing_threshold_percent": 70.0,
      "target_pass_rate_percent": 75.0,
      "notes": "ABET outcome assessment based on Project 1 for CS majors."
    },
    {
      "accreditation_type": "ABET",
      "code": "LO.PI 1.2",
      "assessment_name": "Project 2",
      "target_majors": ["Cybersecurity"],
      "passing_operator": ">=",
      "passing_threshold_percent": 70.0,
      "target_pass_rate_percent": 75.0,
      "notes": "ABET outcome assessment based on Project 2 for Cybersecurity majors."
    }
  ]
}
```

The example evaluates Project 1 for Computer Science and Project 2 for Cybersecurity in one combined Markdown/Excel report. Each outcome retains its own population, denominator, target, and section totals. Repeat an assignment in separate items to report different majors independently. No pooled pass rate is calculated across unrelated outcomes.

Detailed CSV/Excel rows retain all active students for audit. `Target Major` identifies the item's scope and `Included in Assessment` indicates whether a row contributes to its summary. Sum the count columns to reproduce the summary; do not count all detail rows. A target population with no active students is reported as `NO DATA`. Missing scores among included active students continue to count as failures.

#### Field Explanations:
- `course_name`: Official title of the course.
- `semester`: Term and year (e.g., `Spring 2026`).
- `target_majors`: Set this list **inside each assessment item** to select its programs (full names or abbreviations such as `CS` and `Cyber`, case-insensitive). `[]` explicitly selects all active majors. An omitted item field inherits the optional top-level `target_majors`, preserving old plans; if neither is set, all active majors are evaluated. Invalid values such as a string or `null` raise an error.
- `accreditation_type`: `ABET` or `SACSCOC`.
- `code`: Learning outcome / performance indicator ID (e.g. `"LO.PI 1.1"`).
- `assessment_name`: Name/title of the assessed assignment (e.g. `"Project 1"`). Merges assignment query and assessment name into a single clean attribute. Multi-stage fuzzy matching aligns headers automatically without third-party dependencies!
- `passing_operator`: Operator for passing rule (`>=`, `>`, `<=`, `<`, `==`).
- `passing_threshold_percent`: Minimum percentage score to pass (e.g. `70.0`).
- `target_pass_rate_percent`: Department target pass rate benchmark (e.g. `75.0`).

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
copxxxx-26spring/
├── raw/
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   └── 2026-09-02T1237_Grades-Course_Title_-_16197_COPXXXX_202601.csv
└── assessment_plan.json
```

---

### AFTER Running the Tool
The tool reorganizes the course directory, preserving raw files intact while generating processed output files directly in the root of `copxxxx-26spring/`:

```
copxxxx-26spring/
├── raw/                                           # RAW DATA PRESERVED UNTOUCHED
│   ├── CourseRoster_202601_16197.csv
│   ├── CourseRoster_202601_17089.csv
│   └── 2026-09-02T1237_Grades-Course_Title_-_16197_COPXXXX_202601.csv
├── assessment_plan.json                            # Accreditation configuration plan
├── combined_roster.csv                             # GENERATED: Merged multi-section roster
├── merged_gradebook.csv                            # GENERATED: Consolidated gradebook + roster attributes
├── assessment_results.csv                          # GENERATED: Detailed per-student assessment CSV (Pivot Table ready)
├── assessment_summary.csv                          # GENERATED: Aggregated outcome summary CSV (Pass metrics only)
├── assessment_report.xlsx                          # GENERATED: Formatted multi-tab Excel workbook
├── assessment_report.md                            # GENERATED: Formatted Markdown report
└── README.md                                       # GENERATED: Dataset overview & Processing Audit Log
```

---

## Generated Output Files Explained

1. **`combined_roster.csv`**:
   - Merges all section rosters found in `raw/`.
   - Extracts 5-digit section CRNs (`16197`, `17089`).
   - Normalizes student programs (`Computer Science` $\rightarrow$ `CS`, `Software Design` $\rightarrow$ `SD`, `Cybersecurity` $\rightarrow$ `Cyber`, etc.).
   - Computes student status (`Active` vs. `Withdrawn`).

2. **`merged_gradebook.csv`**:
   - Combines all Canvas gradebook CSVs with `combined_roster.csv`.
   - Preserves **100% of Canvas grade columns intact**.
   - Attaches `Section Number`, `Major`, `Status`, `Level`, and `Cohort` columns to each student row.
   - Dynamic row parser automatically detects `Points Possible` rows and handles presence/absence of `"Manual Posting"` rows without skipping student data.

3. **`assessment_results.csv`**:
   - Detailed per-student dataset listing every student's Program, Name, Section ID, Assessed Assignment, Assessed Score, and Pass/Fail Status.
   - Includes helper count columns (`Pass Count` = 1/0 and `Total Student Count` = 1 for included students; both are zero outside the item’s target majors) for instant **Excel / Google Sheets Pivot Tables**.
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

