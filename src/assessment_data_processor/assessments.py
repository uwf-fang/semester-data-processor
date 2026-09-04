"""
assessments.py - Extracts focused ABET / SACSCOC assessment metrics according to assessment_plan.json.
Supports multiple learning outcomes, performance indicators (PIs), and accreditation types in a single plan.
Includes multi-stage fuzzy matching for assignment names with decorative titles or formatting differences.
Includes Pass Count and Total Student Count helper columns in assessment_results.csv for Excel/Google Sheets Pivot Tables.
Generates assessment_results.csv and assessment_report.md.
"""

import sys
import os
import glob
import csv
import json
import re
import difflib
import statistics
from collections import defaultdict, Counter
from .gradebook import process_gradebook


PROGRAM_ABBR_MAP = {
    "computer science": "CS",
    "software design & development": "SD",
    "cybersecurity": "Cyber",
    "computer engineering": "CE",
    "electrical engineering": "EE"
}


def get_program_abbr(program_name: str) -> str:
    """Returns concise program abbreviation (CS, SD, Cyber, CE, EE)."""
    p_clean = program_name.strip().lower()
    return PROGRAM_ABBR_MAP.get(p_clean, program_name.strip())


def format_markdown_table(headers, rows):
    """Formats a list of headers and list of rows into Markdown table syntax."""
    if not headers or not rows:
        return ""
    col_widths = [len(str(h)) for h in headers]
    for row in rows:
        for i, val in enumerate(row):
            col_widths[i] = max(col_widths[i], len(str(val)))

    header_str = "| " + " | ".join(str(headers[i]).ljust(col_widths[i]) for i in range(len(headers))) + " |"
    sep_str = "| " + " | ".join("-" * col_widths[i] for i in range(len(headers))) + " |"

    row_strs = []
    for row in rows:
        r_str = "| " + " | ".join(str(row[i]).ljust(col_widths[i]) for i in range(len(headers))) + " |"
        row_strs.append(r_str)

    return "\n".join([header_str, sep_str] + row_strs)


def check_passing_condition(pct: float, operator: str, threshold: float) -> bool:
    """Evaluates passing rule based on operator and threshold percent."""
    if operator in [">=", "ge"]:
        return pct >= threshold
    elif operator in [">", "gt"]:
        return pct > threshold
    elif operator in ["<=", "le"]:
        return pct <= threshold
    elif operator in ["<", "lt"]:
        return pct < threshold
    elif operator in ["==", "eq"]:
        return abs(pct - threshold) < 1e-6
    return pct >= threshold


def load_or_create_plan(course_dir: str) -> dict:
    """Loads assessment_plan.json from course_dir or raw/."""
    plan_path = os.path.join(course_dir, "assessment_plan.json")
    raw_plan_path = os.path.join(course_dir, "raw", "assessment_plan.json")

    target_path = None
    if os.path.exists(plan_path):
        target_path = plan_path
    elif os.path.exists(raw_plan_path):
        target_path = raw_plan_path

    if target_path:
        print(f"Loading assessment plan configuration: {target_path}")
        with open(target_path, mode="r", encoding="utf-8") as f:
            return json.load(f)

    default_plan = {
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
            }
        ]
    }
    with open(plan_path, mode="w", encoding="utf-8") as f:
        json.dump(default_plan, f, indent=2)
    return default_plan


def normalize_assignment_title(title: str) -> str:
    """Normalizes assignment titles by removing Canvas ID suffixes, zero-padding, punctuation, and extra whitespace."""
    title_clean = re.sub(r'\(\d+\)$', '', title.strip())
    title_clean = re.sub(r'\b0+(\d+)\b', r'\1', title_clean)
    title_clean = re.sub(r'[-_:]', ' ', title_clean)
    return ' '.join(title_clean.split()).lower()


def find_matching_columns(fieldnames: list, target_query: str) -> list:
    """
    Multi-stage fuzzy and pattern matching to align target assignment queries with gradebook columns.
    """
    target_clean = normalize_assignment_title(target_query)
    exclude_keywords = ["current score", "final score", "unposted", "current grade", "final grade"]

    exact_matches = []
    for f in fieldnames:
        f_clean = normalize_assignment_title(f)
        if any(k in f_clean for k in exclude_keywords):
            continue
        if target_clean in f_clean:
            exact_matches.append(f)

    if exact_matches:
        return exact_matches

    target_tokens = set(target_clean.split())
    scored_candidates = []

    for f in fieldnames:
        f_clean = normalize_assignment_title(f)
        if any(k in f_clean for k in exclude_keywords):
            continue

        f_tokens = set(f_clean.split())

        if target_tokens.issubset(f_tokens):
            score = 0.90
        else:
            score = difflib.SequenceMatcher(None, target_clean, f_clean).ratio()

        if score >= 0.55:
            scored_candidates.append((f, score))

    scored_candidates.sort(key=lambda x: x[1], reverse=True)
    if scored_candidates:
        best_match = scored_candidates[0][0]
        score = scored_candidates[0][1]
        print(f"Fuzzy Matched target query '{target_query}' -> '{best_match}' (Confidence: {round(score * 100, 1)}%)")
        return [best_match]

    return []


def extract_assessments(course_dir: str) -> str:
    merged_gb_path = os.path.join(course_dir, "merged_gradebook.csv")
    raw_dir = os.path.join(course_dir, "raw")

    if not os.path.exists(merged_gb_path):
        print(f"Merged gradebook not found at {merged_gb_path}. Running process_gradebook first...")
        process_gradebook(course_dir)

    plan = load_or_create_plan(course_dir)
    assessments_config = plan.get("assessments", [])
    target_majors_filter = plan.get("target_majors", [])
    target_majors_lower = [m.strip().lower() for m in target_majors_filter]

    student_records = []
    with open(merged_gb_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        for row in reader:
            student_records.append(row)

    gradebook_files = sorted(glob.glob(os.path.join(raw_dir, "*Grades*.csv"))) or sorted(glob.glob(os.path.join(raw_dir, "*[Gg]rade*.csv")))
    points_possible_map = {}
    if gradebook_files:
        with open(gradebook_files[0], mode="r", encoding="utf-8-sig") as f:
            reader_raw = list(csv.reader(f))
            if len(reader_raw) >= 3:
                raw_headers = [col.strip() for col in reader_raw[0]]
                pts_row = reader_raw[2]
                for h, p in zip(raw_headers, pts_row):
                    clean_h = re.sub(r'\s*\(\d+\)$', '', h)
                    try:
                        pts_val = float(p.strip())
                        points_possible_map[h] = pts_val
                        points_possible_map[clean_h] = pts_val
                    except ValueError:
                        points_possible_map[h] = None
                        points_possible_map[clean_h] = None

    total_enrolled = len(student_records)
    active_students = [s for s in student_records if s.get("Status", "").lower() == "active"]
    withdrawn_students = [s for s in student_records if s.get("Status", "").lower() == "withdrawn"]

    if not active_students:
        active_students = student_records

    per_student_export_rows = []
    report_assessment_results = []

    for item_cfg in assessments_config:
        accreditation_type = item_cfg.get("accreditation_type", "ABET").upper()
        code = str(item_cfg.get("code", item_cfg.get("lo_pi_code", ""))).strip()

        raw_ass_name = item_cfg.get("assessment_name") or item_cfg.get("assignment_name") or ""
        raw_col_query = item_cfg.get("assignment_column") or raw_ass_name or "Project 1"
        target_query = raw_col_query.strip()

        if not raw_ass_name:
            raw_ass_name = target_query

        if not code:
            m = re.search(r'(LO\.?PI\s*[\d\.]+)', raw_ass_name, re.IGNORECASE)
            if m:
                code = m.group(1).strip()

        # Clean leading code prefixes from raw_ass_name so Code and Assignment Name aren't redundant
        clean_ass_name = re.sub(r'^\[?LO\.?PI\s*[\d\.]+\]?\s*[-:]?\s*', '', raw_ass_name, flags=re.IGNORECASE).strip()
        full_assessment_name = clean_ass_name if clean_ass_name else target_query

        op = item_cfg.get("passing_operator", ">=")
        pass_threshold = float(item_cfg.get("passing_threshold_percent", 70.0))
        target_rate = float(item_cfg.get("target_pass_rate_percent", 75.0))
        notes = item_cfg.get("notes", "")

        matched_cols = find_matching_columns(fieldnames, target_query)
        if not matched_cols:
            print(f"Warning: Assignment column '{target_query}' not found in gradebook.")
            continue

        clean_matched = [re.sub(r'\s*\(\d+\)$', '', c) for c in matched_cols]
        matched_col_display = ", ".join(clean_matched) if len(clean_matched) > 1 else clean_matched[0]
        pts_possible = points_possible_map.get(matched_cols[0], points_possible_map.get(clean_matched[0], 100.0))

        per_student_rows = []
        for s in active_students:
            fn = s.get("First Name", "")
            ln = s.get("Last Name", "")
            if not fn or not ln:
                std_name = s.get("Student", "")
                parts = std_name.split(",")
                if len(parts) == 2:
                    ln, fn = parts[0].strip(), parts[1].strip()
                else:
                    fn, ln = std_name, ""

            prog_full = s.get("Major", s.get("Program", "Unknown"))
            prog_abbr = get_program_abbr(prog_full)
            sec_num = s.get("Section Number", s.get("Section", "Unknown"))

            score_val = None
            for col_candidate in matched_cols:
                val_str = s.get(col_candidate, "").strip()
                if val_str != "" and val_str.lower() != "n/a":
                    try:
                        score_val = float(val_str)
                        break
                    except ValueError:
                        pass

            if score_val is not None and pts_possible and pts_possible > 0:
                pct = (score_val / pts_possible) * 100.0
                is_pass = check_passing_condition(pct, op, pass_threshold)
                pass_status = "Pass" if is_pass else "Fail"
                display_score = f"{score_val:.2f}" if score_val % 1 != 0 else f"{int(score_val)}"
                score_pct_str = f"{round(pct, 1)}%"
            else:
                pct = 0.0
                is_pass = False
                pass_status = "Fail"
                display_score = "N/A"
                score_pct_str = "N/A"

            pts_possible_str = f"{pts_possible:g}" if isinstance(pts_possible, (int, float)) and pts_possible is not None else "100"
            pass_criterion_str = f"{op} {pass_threshold:g}%"

            row_dict = {
                "Accreditation": accreditation_type,
                "Code": code,
                "Assessment Name": full_assessment_name,
                "Program": prog_abbr,
                "Full Major": prog_full,
                "First Name": fn,
                "Last Name": ln,
                "Section ID": sec_num,
                "Assessed Score": display_score,
                "Points Possible": pts_possible_str,
                "Score Percentage (%)": score_pct_str,
                "Passing Criterion": pass_criterion_str,
                "Pass/Fail Status": pass_status,
                "Pass Count": 1 if pass_status == "Pass" else 0,
                "Total Student Count": 1
            }
            per_student_rows.append({
                **row_dict,
                "score_val": score_val,
                "pct": pct,
                "is_pass": is_pass,
                "program_full": prog_full,
                "program_abbr": prog_abbr,
                "first_name": fn,
                "last_name": ln,
                "section_id": sec_num,
                "score_display": display_score,
                "status": pass_status
            })
            per_student_export_rows.append(row_dict)

        per_student_rows.sort(key=lambda r: (r["program_abbr"], r["last_name"].lower(), r["first_name"].lower()))

        if target_majors_lower:
            eval_students = [
                r for r in per_student_rows
                if r["program_full"].strip().lower() in target_majors_lower
                or r["program_abbr"].strip().lower() in target_majors_lower
            ]
        else:
            eval_students = per_student_rows

        sec_group = defaultdict(list)
        for r in eval_students:
            sec_group[r["section_id"]].append(r)

        sec_summary_rows = []
        tot_pass = 0
        tot_eval = len(eval_students)

        for sec_id in sorted(sec_group.keys()):
            s_list = sec_group[sec_id]
            s_total = len(s_list)
            s_pass = sum(1 for r in s_list if r["is_pass"])
            tot_pass += s_pass
            s_ratio = f"{s_pass}/{s_total}"
            s_pct = (s_pass / s_total * 100.0) if s_total > 0 else 0.0
            sec_summary_rows.append({
                "section_id": sec_id,
                "pass_ratio": s_ratio,
                "pass_count": s_pass,
                "total_count": s_total,
                "pass_percent": f"{round(s_pct, 1)}%"
            })

        total_ratio = f"{tot_pass}/{tot_eval}"
        total_pct = (tot_pass / tot_eval * 100.0) if tot_eval > 0 else 0.0
        overall_status = "MET" if total_pct >= target_rate else "NOT MET"

        report_assessment_results.append({
            "cfg": item_cfg,
            "code": code,
            "full_assessment_name": full_assessment_name,
            "matched_col": matched_col_display,
            "pts_possible": pts_possible,
            "per_student_rows": per_student_rows,
            "eval_students": eval_students,
            "sec_summary_rows": sec_summary_rows,
            "total_ratio": total_ratio,
            "total_pct": round(total_pct, 1),
            "tot_pass": tot_pass,
            "tot_eval": tot_eval,
            "overall_status": overall_status
        })

    # Detailed Per-Student assessment_results.csv
    results_csv_output_path = os.path.join(course_dir, "assessment_results.csv")
    if per_student_export_rows:
        headers = ["Program", "First Name", "Last Name", "Section ID", "Assessed Score", "Points Possible", "Score Percentage (%)", "Passing Criterion", "Pass/Fail Status", "Pass Count", "Total Student Count", "Accreditation", "Code", "Assessment Name", "Full Major"]
        with open(results_csv_output_path, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for r in per_student_export_rows:
                writer.writerow({k: r[k] for k in headers})
        print(f"Generated detailed student assessment results CSV: {results_csv_output_path}")

    # Aggregated Summary CSV: assessment_summary.csv (Pivot Table-like summary without Fail columns)
    summary_csv_output_path = os.path.join(course_dir, "assessment_summary.csv")
    if report_assessment_results:
        summary_headers = [
            "Accreditation",
            "Code",
            "Assessment Name",
            "Assessed Assignment",
            "Passing Criterion",
            "Target Benchmark",
            "Target Major",
            "Section ID",
            "Passed Count",
            "Total Evaluated",
            "Pass Rate (%)",
            "Outcome Status"
        ]
        summary_export_rows = []
        for res in report_assessment_results:
            cfg = res["cfg"]
            acc_type = cfg.get("accreditation_type", "ABET").upper()
            code = res.get("code", "")
            ass_name = res.get("full_assessment_name", cfg.get("assessment_name", "ABET Outcome"))
            matched_col = res["matched_col"]
            pass_criterion = f"{cfg.get('passing_operator', '>=')} {cfg.get('passing_threshold_percent', 70):g}%"
            target_bm = f"{cfg.get('target_pass_rate_percent', 75):g}% Pass Rate"
            target_major = ", ".join(target_majors_filter) if target_majors_filter else "All Majors"
            target_rate = float(cfg.get("target_pass_rate_percent", 75.0))

            for sr in res["sec_summary_rows"]:
                s_pct_val = float(sr["pass_percent"].rstrip("%"))
                s_status = "MET" if s_pct_val >= target_rate else "NOT MET"
                summary_export_rows.append({
                    "Accreditation": acc_type,
                    "Code": code,
                    "Assessment Name": ass_name,
                    "Assessed Assignment": matched_col,
                    "Passing Criterion": pass_criterion,
                    "Target Benchmark": target_bm,
                    "Target Major": target_major,
                    "Section ID": sr["section_id"],
                    "Passed Count": sr["pass_count"],
                    "Total Evaluated": sr["total_count"],
                    "Pass Rate (%)": sr["pass_percent"],
                    "Outcome Status": s_status
                })

            summary_export_rows.append({
                "Accreditation": acc_type,
                "Code": code,
                "Assessment Name": ass_name,
                "Assessed Assignment": matched_col,
                "Passing Criterion": pass_criterion,
                "Target Benchmark": target_bm,
                "Target Major": target_major,
                "Section ID": "Total (All Sections)",
                "Passed Count": res["tot_pass"],
                "Total Evaluated": res["tot_eval"],
                "Pass Rate (%)": f"{res['total_pct']}%",
                "Outcome Status": res["overall_status"]
            })

        with open(summary_csv_output_path, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=summary_headers)
            writer.writeheader()
            for r in summary_export_rows:
                writer.writerow(r)
        print(f"Generated aggregated assessment summary CSV: {summary_csv_output_path}")

    # Formatted assessment_report.md
    report_output_path = os.path.join(course_dir, "assessment_report.md")
    acc_types_set = sorted(list(set(cfg.get("accreditation_type", "ABET").upper() for cfg in assessments_config))) if assessments_config else ["ABET"]
    acc_types_str = " / ".join(acc_types_set)

    with open(report_output_path, mode="w", encoding="utf-8") as f:
        f.write(f"# {acc_types_str} Assessment Report - {os.path.basename(course_dir)}\n\n")
        f.write("## Executive Summary\n\n")
        f.write(f"- **Course**: {plan.get('course_name', course_dir)}\n")
        f.write(f"- **Semester**: {plan.get('semester', '')}\n")
        f.write(f"- **Total Roster Registrations**: {total_enrolled} ({len(active_students)} Active, {len(withdrawn_students)} Withdrawn)\n")
        f.write(f"- **Target Major Evaluated**: {', '.join(target_majors_filter) if target_majors_filter else 'All Majors'}\n")
        f.write(f"- **Total Learning Outcomes / Performance Indicators Assessed**: {len(assessments_config)}\n\n")

        f.write("### Roster Enrollment Breakdown by Major & Section\n\n")
        recon_headers = ["Section ID", "Major / Program", "Active Students", "Withdrawn Students", "Total Registered"]
        recon_rows = []
        sec_maj_group = defaultdict(lambda: {"active": 0, "withdrawn": 0})
        for s in student_records:
            sec_id = s.get("Section Number", s.get("Section", "Unknown"))
            maj_id = s.get("Major", s.get("Program", "Unknown"))
            st = s.get("Status", "").lower()
            if st == "withdrawn":
                sec_maj_group[(sec_id, maj_id)]["withdrawn"] += 1
            else:
                sec_maj_group[(sec_id, maj_id)]["active"] += 1

        for (sec_id, maj_id), counts in sorted(sec_maj_group.items()):
            act = counts["active"]
            wd = counts["withdrawn"]
            recon_rows.append([sec_id, maj_id, act, wd, act + wd])

        recon_rows.append([
            "**Total All Sections**",
            "**All Majors Combined**",
            f"**{len(active_students)}**",
            f"**{len(withdrawn_students)}**",
            f"**{total_enrolled}**"
        ])
        f.write(format_markdown_table(recon_headers, recon_rows) + "\n\n")

        for res in report_assessment_results:
            cfg = res["cfg"]
            acc_type = cfg.get("accreditation_type", "ABET").upper()
            ass_name = cfg.get("assessment_name", "ABET Outcome")
            
            f.write(f"## [{acc_type}] {ass_name} Assessment Summary\n\n")
            f.write(f"- **Assessed Assignment**: `{res['matched_col']}`\n")
            f.write(f"- **Passing Criterion**: Score {cfg.get('passing_operator', '>=')} {cfg.get('passing_threshold_percent', 70)}%\n")
            f.write(f"- **Target Benchmark**: {cfg.get('target_pass_rate_percent', 75)}% Pass Rate\n")
            f.write(f"- **Evaluated Population**: {res['tot_eval']} CS Majors (out of {len(active_students)} total active students across all majors)\n")
            f.write(f"- **Overall Pass/Fail Ratio**: **{res['total_ratio']}** (**{res['total_pct']}%**)\n")
            f.write(f"- **Outcome Status**: **{res['overall_status']}**\n\n")

            f.write("### Pass/Fail Ratio & Percentage by Section (CS Majors Only)\n\n")
            sec_headers = ["Section ID", "Pass / Total Ratio", "Passed Count", "Total CS Students", "Pass Rate (%)"]
            sec_rows = []
            for sr in res["sec_summary_rows"]:
                sec_rows.append([
                    sr["section_id"],
                    sr["pass_ratio"],
                    sr["pass_count"],
                    sr["total_count"],
                    sr["pass_percent"]
                ])
            sec_rows.append([
                "**Total Sum (All Sections)**",
                f"**{res['total_ratio']}**",
                f"**{res['tot_pass']}**",
                f"**{res['tot_eval']}**",
                f"**{res['total_pct']}%**"
            ])
            f.write(format_markdown_table(sec_headers, sec_rows) + "\n\n")

            f.write("### Complete Student Assessment Results Table (All Active Students)\n\n")
            std_headers = ["Program", "First Name", "Last Name", "Section ID", f"Assessed Score ({res['matched_col']})", "Status"]
            std_rows = []
            for st in res["per_student_rows"]:
                std_rows.append([
                    st["program_abbr"],
                    st["first_name"],
                    st["last_name"],
                    st["section_id"],
                    st["score_display"],
                    st["status"]
                ])
            f.write(format_markdown_table(std_headers, std_rows) + "\n\n")

    print(f"Generated assessment Markdown report: {report_output_path}")

    # Direct Excel Workbook Generation: assessment_report.xlsx
    excel_output_path = export_excel_report(
        course_dir=course_dir,
        plan=plan,
        report_assessment_results=report_assessment_results,
        per_student_export_rows=per_student_export_rows,
        total_enrolled=total_enrolled,
        active_count=len(active_students),
        withdrawn_count=len(withdrawn_students)
    )

    return report_output_path


def export_excel_report(course_dir: str, plan: dict, report_assessment_results: list, per_student_export_rows: list, total_enrolled: int, active_count: int, withdrawn_count: int) -> str:
    """Generates formatted Excel workbook (assessment_report.xlsx) with Plan details, simplified summary, and student results."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Warning: 'openpyxl' module not found. Skipping .xlsx Excel report generation. Install via 'pip install openpyxl' or run with 'uv run'.")
        return ""

    wb = openpyxl.Workbook()

    # Sheet 1: Assessment Summary & Plan
    ws_summary = wb.active
    ws_summary.title = "Assessment Summary & Plan"
    ws_summary.views.sheetView[0].showGridLines = True

    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
    sub_font = Font(name="Calibri", size=11, bold=True, color="333333")
    regular_font = Font(name="Calibri", size=11)

    met_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    met_font = Font(name="Calibri", size=11, bold=True, color="375623")
    not_met_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    not_met_font = Font(name="Calibri", size=11, bold=True, color="C65911")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    row_idx = 1
    ws_summary.cell(row=row_idx, column=1, value=f"{plan.get('course_name', os.path.basename(course_dir))} — Assessment Executive Summary").font = title_font
    row_idx += 1
    ws_summary.cell(row=row_idx, column=1, value=f"Semester: {plan.get('semester', '')} | Evaluated Major: {', '.join(plan.get('target_majors', [])) or 'All Majors'} | Active Students: {active_count}").font = regular_font
    row_idx += 2

    # Section 1: Assessment Plan Details
    ws_summary.cell(row=row_idx, column=1, value="1. Assessment Plan Details").font = section_font
    row_idx += 1

    plan_headers = ["Code", "Accreditation", "Assessed Assignment", "Passing Criterion", "Target Benchmark", "Notes"]
    for col_idx, h in enumerate(plan_headers, 1):
        cell = ws_summary.cell(row=row_idx, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row_idx += 1

    assessments_config = plan.get("assessments", [])
    for item in assessments_config:
        code_val = item.get("code", item.get("lo_pi_code", ""))
        acc_val = item.get("accreditation_type", "ABET").upper()
        raw_name = item.get("assessment_name") or item.get("assignment_name") or item.get("assignment_column", "")
        clean_name = re.sub(r'^\[?LO\.?PI\s*[\d\.]+\]?\s*[-:]?\s*', '', raw_name, flags=re.IGNORECASE).strip()
        ass_name_val = clean_name if clean_name else raw_name

        crit_val = f"{item.get('passing_operator', '>=')} {item.get('passing_threshold_percent', 70):g}%"
        bm_val = f"{item.get('target_pass_rate_percent', 75):g}% Pass Rate"
        notes_val = item.get("notes", "")

        row_vals = [code_val, acc_val, ass_name_val, crit_val, bm_val, notes_val]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [1, 2, 4, 5]:
                cell.alignment = Alignment(horizontal="center")
        row_idx += 1

    row_idx += 2

    # Section 2: Assessment Summary (Simplified Title & Non-Redundant Columns)
    ws_summary.cell(row=row_idx, column=1, value="2. Assessment Summary").font = section_font
    row_idx += 1

    sum_headers = ["Code", "Assessed Assignment", "Passed Count", "Total Evaluated", "Pass Ratio", "Pass Rate (%)", "Outcome Status"]
    for col_idx, h in enumerate(sum_headers, 1):
        cell = ws_summary.cell(row=row_idx, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row_idx += 1

    for res in report_assessment_results:
        cfg = res["cfg"]
        code_val = res.get("code", "")
        matched_col = res.get("matched_col", "")
        tot_pass = res.get("tot_pass", 0)
        tot_eval = res.get("tot_eval", 0)
        ratio_str = f"{tot_pass}/{tot_eval}"
        pct_str = f"{res.get('total_pct', 0.0)}%"
        status_str = res.get("overall_status", "MET")

        row_vals = [code_val, matched_col, tot_pass, tot_eval, ratio_str, pct_str, status_str]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx in [1, 3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center")
            if col_idx == 7:
                cell.alignment = Alignment(horizontal="center")
                if status_str == "MET":
                    cell.fill = met_fill
                    cell.font = met_font
                else:
                    cell.fill = not_met_fill
                    cell.font = not_met_font
        row_idx += 1

    row_idx += 2

    # Section 3: Vertical Outcome Breakdown Cards
    ws_summary.cell(row=row_idx, column=1, value="3. Section Outcome Breakdown Cards").font = section_font
    row_idx += 1

    for res in report_assessment_results:
        code_val = res.get("code", "")
        matched_col = res.get("matched_col", "")

        ws_summary.cell(row=row_idx, column=1, value=f"Outcome Card: {code_val} ({matched_col})").font = sub_font
        row_idx += 1

        card_headers = ["Section ID", "Passed Count", "Total Evaluated", "Pass Ratio", "Pass Rate (%)", "Status"]
        for col_idx, h in enumerate(card_headers, 1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row_idx += 1

        target_rate = float(res["cfg"].get("target_pass_rate_percent", 75.0))
        for sr in res["sec_summary_rows"]:
            s_pct_val = float(sr["pass_percent"].rstrip("%"))
            s_status = "MET" if s_pct_val >= target_rate else "NOT MET"
            row_vals = [sr["section_id"], sr["pass_count"], sr["total_count"], sr["pass_ratio"], sr["pass_percent"], s_status]
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx in [1, 2, 3, 4, 5]:
                    cell.alignment = Alignment(horizontal="center")
                if col_idx == 6:
                    cell.alignment = Alignment(horizontal="center")
                    cell.fill = met_fill if s_status == "MET" else not_met_fill
                    cell.font = met_font if s_status == "MET" else not_met_font
            row_idx += 1

        tot_vals = ["Total (All Sections)", res["tot_pass"], res["tot_eval"], res["total_ratio"], f"{res['total_pct']}%", res["overall_status"]]
        for col_idx, val in enumerate(tot_vals, 1):
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.border = thin_border
            if col_idx in [1, 2, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center")
            if col_idx == 6:
                cell.alignment = Alignment(horizontal="center")
                cell.fill = met_fill if res["overall_status"] == "MET" else not_met_fill
                cell.font = met_font if res["overall_status"] == "MET" else not_met_font
        row_idx += 2

    # Auto-adjust column widths
    for col in ws_summary.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            max_len = max(max_len, len(val_str))
        ws_summary.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Sheet 2: Student Assessment Results
    ws_results = wb.create_sheet(title="Student Assessment Results")
    ws_results.views.sheetView[0].showGridLines = True

    if per_student_export_rows:
        res_headers = ["Program", "First Name", "Last Name", "Section ID", "Assessed Score", "Points Possible", "Score Percentage (%)", "Passing Criterion", "Pass/Fail Status", "Pass Count", "Total Student Count", "Accreditation", "Code", "Assessment Name", "Full Major"]
        for col_idx, h in enumerate(res_headers, 1):
            cell = ws_results.cell(row=1, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, r in enumerate(per_student_export_rows, 2):
            for c_idx, h in enumerate(res_headers, 1):
                val = r.get(h, "")
                cell = ws_results.cell(row=r_idx, column=c_idx, value=val)
                cell.font = regular_font
                cell.border = thin_border
                if h in ["Program", "Section ID", "Pass/Fail Status", "Pass Count", "Total Student Count", "Accreditation", "Code"]:
                    cell.alignment = Alignment(horizontal="center")

        for col in ws_results.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                max_len = max(max_len, len(val_str))
            ws_results.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output_path = os.path.join(course_dir, "assessment_report.xlsx")
    wb.save(output_path)
    print(f"Generated Excel assessment report workbook: {output_path}")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 -m assessment_data_processor.assessments <course_dir>")
        sys.exit(1)

    extract_assessments(sys.argv[1])
