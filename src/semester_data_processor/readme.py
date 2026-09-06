"""
readme.py - Generates and maintains dataset documentation and processing log in <course_dir>/README.md.
"""

import os
import glob
import csv
import json
import datetime
import re


def get_file_timestamp(filepath: str) -> str:
    """Returns ISO-formatted modification timestamp for a file."""
    if not os.path.exists(filepath):
        return "N/A"
    mtime = os.path.getmtime(filepath)
    dt = datetime.datetime.fromtimestamp(mtime)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def parse_existing_log(readme_path: str) -> list:
    """Extracts existing log rows from an existing dataset README.md."""
    if not os.path.exists(readme_path):
        return []

    with open(readme_path, "r", encoding="utf-8") as f:
        content = f.read()

    log_section_match = re.search(r"## Processing Log & Updates\s*\n\n(.*?)(?=\n\n## |\Z)", content, re.DOTALL)
    if not log_section_match:
        return []

    log_block = log_section_match.group(1).strip()
    lines = log_block.split("\n")

    log_entries = []
    for line in lines:
        if line.startswith("|") and not line.startswith("| Timestamp") and not line.startswith("|---"):
            cols = [c.strip() for c in line.split("|")[1:-1]]
            if len(cols) >= 4:
                log_entries.append({
                    "timestamp": cols[0],
                    "action": cols[1],
                    "raw_sources": cols[2],
                    "details": cols[3]
                })
    return log_entries


def generate_output_readme(course_dir: str, action_name: str = "Initial Ingestion", log_note: str = "") -> str:
    """
    Generates or updates <course_dir>/README.md describing dataset outputs and maintaining a processing log.
    """
    course_name = os.path.basename(os.path.normpath(course_dir))
    raw_dir = os.path.join(course_dir, "raw")
    readme_path = os.path.join(course_dir, "README.md")

    # Read existing log history if present
    existing_entries = parse_existing_log(readme_path)

    # Inspect raw directory files & injection timestamps
    raw_files_info = []
    if os.path.exists(raw_dir):
        for fname in sorted(os.listdir(raw_dir)):
            if fname.startswith("."):
                continue
            fpath = os.path.join(raw_dir, fname)
            if os.path.isfile(fpath):
                t_str = get_file_timestamp(fpath)
                raw_files_info.append(f"`{fname}` (Injected/Modified: {t_str})")

    raw_sources_summary = "<br>".join(raw_files_info) if raw_files_info else "No raw input files found"

    # Inspect output files stats
    roster_path = os.path.join(course_dir, "combined_roster.csv")
    gb_path = os.path.join(course_dir, "merged_gradebook.csv")
    ass_results_path = os.path.join(course_dir, "assessment_results.csv")
    ass_summary_path = os.path.join(course_dir, "assessment_summary.csv")
    plan_path = os.path.join(course_dir, "assessment_plan.json")
    if not os.path.exists(plan_path):
        plan_path = os.path.join(raw_dir, "assessment_plan.json")

    roster_stats = "Not generated yet"
    if os.path.exists(roster_path):
        with open(roster_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.DictReader(f))
            tot = len(reader)
            active = sum(1 for r in reader if r.get("Status") == "Active")
            withdrawn = sum(1 for r in reader if r.get("Status") == "Withdrawn")
            sections = sorted(list(set(r.get("Section Number", "Unknown") for r in reader)))
            roster_stats = f"{tot} total students ({active} Active, {withdrawn} Withdrawn) across section(s) {', '.join(sections)}"

    gb_stats = "Not generated yet"
    if os.path.exists(gb_path):
        with open(gb_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.DictReader(f))
            cols_cnt = len(reader[0].keys()) if reader else 0
            gb_stats = f"{len(reader)} student records integrated with {cols_cnt} gradebook & roster columns"

    ass_stats = "Not generated yet"
    if os.path.exists(ass_summary_path):
        with open(ass_summary_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.DictReader(f))
            met_cnt = sum(1 for r in reader if r.get("Outcome Status") == "MET" and r.get("Section ID") == "Total (All Sections)")
            tot_outcomes = sum(1 for r in reader if r.get("Section ID") == "Total (All Sections)")
            ass_stats = f"{tot_outcomes} learning outcome(s) evaluated ({met_cnt} MET)"

    # Build new log entry
    current_time_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not log_note:
        if action_name == "process-rosters" or action_name == "Process Rosters":
            log_note = f"Roster merged: {roster_stats}"
        elif action_name == "process-gradebook" or action_name == "Process Gradebook":
            log_note = f"Gradebook merged: {gb_stats}"
        elif action_name == "extract-assessments" or action_name == "Extract Assessments":
            log_note = f"Assessments evaluated: {ass_stats}"
        elif action_name == "run-all" or action_name == "Full Pipeline Run":
            log_note = f"Executed full pipeline: {roster_stats}; {gb_stats}"
        else:
            log_note = "Pipeline dataset update executed successfully."

    new_entry = {
        "timestamp": current_time_str,
        "action": action_name,
        "raw_sources": raw_sources_summary,
        "details": log_note
    }

    # Avoid duplicate identical log entry if executed within same minute with same action
    if not existing_entries or existing_entries[-1]["action"] != action_name or existing_entries[-1]["details"] != log_note:
        existing_entries.append(new_entry)

    # Build README content
    lines = []
    lines.append(f"# Dataset & Semester Processing Documentation — `{course_name}`\n")
    lines.append("## Executive Dataset Overview\n")
    lines.append(f"This directory contains processed semester course data products and accreditation assessment reports for **`{course_name}`**.\n")
    lines.append(f"- **Combined Roster Summary**: {roster_stats}")
    lines.append(f"- **Merged Gradebook Status**: {gb_stats}")
    lines.append(f"- **Accreditation Assessment Status**: {ass_stats}\n")

    lines.append("## Output Data Files & Product Description\n")
    lines.append("The following data products are generated during pipeline runs:\n")
    lines.append("1. **`combined_roster.csv`**:")
    lines.append("   - Consolidated student roster merged across all section CRNs in `raw/`.")
    lines.append("   - Normalizes student program abbreviations (`CS`, `SD`, `Cyber`, `CE`, `EE`), extracts 5-digit section CRNs, and tracks student withdrawal status (`Active` vs. `Withdrawn`).\n")

    lines.append("2. **`merged_gradebook.csv`**:")
    lines.append("   - Full Canvas gradebook export integrated with student roster metadata (`Section Number`, `Major`, `Status`, `Level`, `Cohort`).")
    lines.append("   - Preserves 100% of raw Canvas grade columns intact without altering original student scores.\n")

    lines.append("3. **`assessment_results.csv`**:")
    lines.append("   - Detailed per-student assessment outcome evaluation table.")
    lines.append("   - Includes numerical helper columns (`Pass Count` = 1/0, `Total Student Count` = 1) ready for custom Pivot Table analysis in Excel or Google Sheets.\n")

    lines.append("4. **`assessment_summary.csv`**:")
    lines.append("   - Aggregated outcome assessment summary metrics per section and total sum across all sections.")
    lines.append("   - Documents evaluated student counts, pass rates (%), target pass rate benchmarks, and outcome status (`MET` / `NOT MET`).\n")

    lines.append("5. **`assessment_report.xlsx`** & **`assessment_report.md`**:")
    lines.append("   - Executive accreditation reports in formatted Excel workbook (`.xlsx`) and Markdown (`.md`) formats.")
    lines.append("   - Features styled outcome cards, enrollment breakdowns by section & major, and outcome evaluation totals for ABET / SACSCOC review.\n")

    lines.append("6. **`assessment_plan.json`**:")
    lines.append("   - Accreditation configuration plan specifying evaluated learning outcomes, assignment columns, passing rules, and targeted majors.\n")

    lines.append("7. **`raw/` Subdirectory**:")
    lines.append("   - Preserves pristine raw section roster CSVs and Canvas gradebook exports. Original raw data is strictly kept untouched.\n")

    lines.append("## Processing Log & Updates\n")
    lines.append("The table below maintains a record of data ingestion, roster updates, gradebook integrations, and assessment pipeline execution runs:\n")

    lines.append("| Timestamp | Action / Stage | Raw Input Files & Ingestion Time | Update Details / Metrics |")
    lines.append("|---|---|---|---|")
    for entry in existing_entries:
        lines.append(f"| {entry['timestamp']} | {entry['action']} | {entry['raw_sources']} | {entry['details']} |")

    lines.append("\n---\n*Documentation automatically generated and maintained by `semester-data-processor`.*")

    content = "\n".join(lines)
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"Generated dataset README documentation & log: {readme_path}")
    return readme_path
