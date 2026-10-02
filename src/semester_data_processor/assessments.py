"""
assessments.py - Extracts focused ABET / SACSCOC assessment metrics according to assessment_plan.json.
Supports multiple learning outcomes, performance indicators (PIs), and accreditation types in a single plan.
Includes multi-stage fuzzy matching for assignment names with decorative titles or formatting differences.
Includes Pass Count and Total Student Count helper columns in assessment_results.csv for Excel/Google Sheets Pivot Tables.
Generates assessment_results.csv and assessment_report.md.
"""

import sys
import os
import csv
import json
import re
import difflib
import math
from collections import defaultdict
from .gradebook import process_gradebook, discover_gradebooks, read_gradebooks
from .rosters import load_instructor_mapping, load_section_term_mapping
from .readme import generate_output_readme


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


def assessment_majors(plan: dict, item: dict) -> list:
    """Item scope overrides the legacy plan default; [] explicitly means all."""
    majors = item.get("target_majors", plan.get("target_majors", []))
    if not isinstance(majors, list) or any(not isinstance(m, str) or not m.strip() for m in majors):
        raise ValueError("target_majors must be a list of nonempty major names; use [] for all majors")
    return [m.strip() for m in majors]


def major_label(majors: list) -> str:
    return ", ".join(majors) if majors else "All Majors"


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

    term_info = resolve_course_and_academic_year(course_dir, {})
    default_plan = {
        "course_name": f"{term_info['course_code']} - Course Title",
        "academic_year": term_info.get("academic_year_label", "2025-26"),
        "assessments": [
            {
                "accreditation_type": "ABET",
                "code": "LO1.PI1",
                "assessment_name": "Project 1",
                "assignment_column": "Project 1",
                "target_majors": ["Computer Science"],
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

    literal_matches = [f for f in fieldnames if normalize_assignment_title(f) == target_clean]
    if literal_matches:
        return literal_matches

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

    if discover_gradebooks(raw_dir) or not os.path.exists(merged_gb_path):
        print("Refreshing merged gradebook before assessment extraction...")
        process_gradebook(course_dir)

    plan = load_or_create_plan(course_dir)
    assessments_config = plan.get("assessments", [])
    for item in assessments_config:
        assessment_majors(plan, item)
    plan_scope = "; ".join(dict.fromkeys(major_label(assessment_majors(plan, item)) for item in assessments_config)) or "No assessments"

    student_records = []
    with open(merged_gb_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        for row in reader:
            student_records.append(row)

    points_possible_map = {}
    if discover_gradebooks(raw_dir):
        _, _, points_possible_map = read_gradebooks(raw_dir)

    total_enrolled = len(student_records)
    active_students = [s for s in student_records if s.get("Status", "Active").strip().lower() in ("", "active")]
    withdrawn_students = [s for s in student_records if s.get("Status", "").strip().lower() == "withdrawn"]

    instructor_map = load_instructor_mapping(course_dir)

    per_student_export_rows = []
    report_assessment_results = []

    for item_cfg in assessments_config:
        target_majors = assessment_majors(plan, item_cfg)
        target_majors_lower = [m.lower() for m in target_majors]
        target_major = major_label(target_majors)
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
            raise ValueError(f"Assignment column {target_query!r} not found in gradebook")

        clean_matched = [re.sub(r'\s*\(\d+\)$', '', c) for c in matched_cols]
        matched_col_display = ", ".join(clean_matched) if len(clean_matched) > 1 else clean_matched[0]
        pts_possible = None

        # Pre-assign section-relative Anonymous IDs (S1, S2, ...) for evaluated students
        eval_candidates = []
        for s in active_students:
            prog_full = s.get("Major", s.get("Program", "Unknown"))
            prog_abbr = get_program_abbr(prog_full)
            included = not target_majors_lower or prog_full.strip().lower() in target_majors_lower or prog_abbr.lower() in target_majors_lower
            if included:
                eval_candidates.append(s)

        sec_eval_map = defaultdict(list)
        for s in eval_candidates:
            sec_num = s.get("Section Number", s.get("Section", "Unknown"))
            sec_eval_map[sec_num].append(s)

        sec_anon_map = {}
        for sec_num, s_list in sec_eval_map.items():
            s_list.sort(key=lambda s: (s.get("Last Name", "").lower(), s.get("First Name", "").lower()))
            for idx, s in enumerate(s_list, 1):
                sid = s.get("SIS User ID", s.get("UWFID", ""))
                sec_anon_map[(sec_num, sid)] = f"S{idx}"

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
            instructor_name = instructor_map.get(sec_num, s.get("Instructor", ""))
            sid = s.get("SIS User ID", s.get("UWFID", ""))
            anon_id = sec_anon_map.get((sec_num, sid), "")

            included = not target_majors_lower or prog_full.strip().lower() in target_majors_lower or prog_abbr.lower() in target_majors_lower
            score_val = None
            pts_possible = None
            for col_candidate in matched_cols:
                maximum = points_possible_map.get((s.get("SIS User ID", ""), col_candidate))
                if pts_possible is None:
                    pts_possible = maximum
                try:
                    candidate_score = float(s.get(col_candidate, "").strip())
                except ValueError:
                    continue
                if not math.isfinite(candidate_score):
                    raise ValueError(f"Non-finite score in {col_candidate!r}")
                if maximum is None or maximum <= 0:
                    raise ValueError(f"Missing or invalid Points Possible for {col_candidate!r}, "
                                     f"SIS ID {s.get('SIS User ID', '')}")
                if score_val is not None:
                    raise ValueError(f"Multiple scored columns match {target_query!r} for "
                                     f"SIS ID {s.get('SIS User ID', '')}; use an exact assignment_column")
                score_val = candidate_score
                pts_possible = maximum

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

            pts_possible_str = f"{pts_possible:g}" if isinstance(pts_possible, (int, float)) and pts_possible is not None else "N/A"
            pass_criterion_str = f"{op} {pass_threshold:g}%"

            row_dict = {
                "Anonymous ID": anon_id,
                "Instructor": instructor_name,
                "Target Major": target_major,
                "Included in Assessment": "Yes" if included else "No",
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
                "Pass Count": 1 if included and pass_status == "Pass" else 0,
                "Total Student Count": 1 if included else 0
            }
            per_student_rows.append({
                **row_dict,
                "Anonymous ID": anon_id,
                "Instructor": instructor_name,
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
        overall_status = ("MET" if total_pct >= target_rate else "NOT MET") if tot_eval else "NO DATA"

        report_assessment_results.append({
            "cfg": item_cfg,
            "target_major": target_major,
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
    headers = ["Anonymous ID", "Instructor", "Program", "First Name", "Last Name", "Section ID", "Assessed Score", "Points Possible", "Score Percentage (%)", "Passing Criterion", "Pass/Fail Status", "Pass Count", "Total Student Count", "Accreditation", "Code", "Assessment Name", "Full Major", "Target Major", "Included in Assessment"]
    with open(results_csv_output_path, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for r in per_student_export_rows:
            writer.writerow({k: r[k] for k in headers})
    print(f"Generated detailed student assessment results CSV: {results_csv_output_path}")

    # Aggregated Summary CSV: assessment_summary.csv (Pivot Table-like summary without Fail columns)
    summary_csv_output_path = os.path.join(course_dir, "assessment_summary.csv")
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
        target_major = res["target_major"]
        target_rate = float(cfg.get("target_pass_rate_percent", 75.0))

        for sr in res["sec_summary_rows"]:
            s_pct_val = sr["pass_count"] / sr["total_count"] * 100.0
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

    term_info = resolve_course_and_academic_year(course_dir, plan)

    with open(report_output_path, mode="w", encoding="utf-8") as f:
        f.write(f"# {acc_types_str} Assessment Report - {os.path.basename(course_dir)}\n\n")
        f.write("## Executive Summary\n\n")
        f.write(f"- **Course**: {plan.get('course_name', course_dir)}\n")
        f.write(f"- **Academic Year**: {term_info['academic_year_label']}\n")
        if plan.get("semester"):
            f.write(f"- **Semester**: {plan.get('semester')}\n")
        elif term_info.get("is_both"):
            f.write(f"- **Active Term**: Fall & Spring ({term_info['fall_term']} & {term_info['spring_term']})\n")
        else:
            active_sem_name = "Fall" if term_info["has_fall"] else "Spring"
            f.write(f"- **Active Term**: {active_sem_name} ({term_info['active_term']})\n")
        f.write(f"- **Total Roster Registrations**: {total_enrolled} ({len(active_students)} Active, {len(withdrawn_students)} Withdrawn)\n")
        f.write(f"- **Target Major Evaluated**: {plan_scope}\n")
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
            f.write(f"- **Evaluated Population**: {res['tot_eval']} students — {res['target_major']} (out of {len(active_students)} total active students across all majors)\n")
            f.write(f"- **Overall Pass/Fail Ratio**: **{res['total_ratio']}** (**{res['total_pct']}%**)\n")
            f.write(f"- **Outcome Status**: **{res['overall_status']}**\n\n")

            f.write(f"### Pass/Fail Ratio & Percentage by Section ({res['target_major']})\n\n")
            sec_headers = ["Section ID", "Pass / Total Ratio", "Passed Count", "Total Evaluated Students", "Pass Rate (%)"]
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
            std_headers = ["Program", "First Name", "Last Name", "Section ID", f"Assessed Score ({res['matched_col']})", "Status", "Included in Assessment"]
            std_rows = []
            for st in res["per_student_rows"]:
                std_rows.append([
                    st["program_abbr"],
                    st["first_name"],
                    st["last_name"],
                    st["section_id"],
                    st["score_display"],
                    st["status"],
                    st["Included in Assessment"]
                ])
            f.write(format_markdown_table(std_headers, std_rows) + "\n\n")

    print(f"Generated assessment Markdown report: {report_output_path}")

    # Generate Academic Year Assessment Excel Reports (new format)
    generated_excel_paths = export_academic_year_excel_report(
        course_dir=course_dir,
        plan=plan,
        report_assessment_results=report_assessment_results,
        instructor_map=instructor_map
    )

    generate_output_readme(course_dir, action_name="Extract Assessments")

    return report_output_path


def resolve_course_and_academic_year(course_dir: str, plan: dict) -> dict:
    """
    Resolves course code, academic year, and active semester term.
    Validates against raw roster filenames and directory naming to prevent misaligned reports,
    issuing clear warnings if assessment_plan.json has conflicting metadata.
    """
    raw_dir = os.path.join(course_dir, "raw")
    dir_name = os.path.basename(os.path.abspath(course_dir)).lower()

    # 1. Course Code Resolution
    m_dir_course = re.search(r'([A-Za-z]{3}\s*\d{4})', dir_name)
    dir_course_code = m_dir_course.group(1).replace(" ", "").upper() if m_dir_course else ""

    plan_course_name = plan.get("course_name", "")
    m_plan_course = re.search(r'([A-Za-z]{3}\s*\d{4})', plan_course_name)
    plan_course_code = m_plan_course.group(1).replace(" ", "").upper() if m_plan_course else ""

    if dir_course_code and plan_course_code and dir_course_code != plan_course_code:
        print(f"\nWARNING: assessment_plan.json course_name '{plan_course_name}' ({plan_course_code}) conflicts with directory '{os.path.basename(course_dir)}' ({dir_course_code})!")
        print(f"Using directory course code '{dir_course_code}' for report naming.\n")
        course_code = dir_course_code
    elif plan_course_code:
        course_code = plan_course_code
    elif dir_course_code:
        course_code = dir_course_code
    else:
        course_code = "COURSE"

    # 2. Extract ground truth term codes from raw/
    raw_terms = set()
    if os.path.exists(raw_dir):
        for fname in os.listdir(raw_dir):
            m_r = re.search(r'CourseRoster_(20\d{2}(?:08|01|05))', fname, re.IGNORECASE)
            if m_r:
                raw_terms.add(m_r.group(1))
            m_s = re.search(r'(20\d{2}(?:08|01|05))[-_]sections', fname, re.IGNORECASE)
            if m_s:
                raw_terms.add(m_s.group(1))
            m_g = re.search(r'_(20\d{2}(?:08|01|05))\.csv', fname, re.IGNORECASE)
            if m_g:
                raw_terms.add(m_g.group(1))

    # Also detect from directory name, e.g. 25fall -> 202508, 26spring -> 202601
    dir_term = ""
    m_dir_term = re.search(r'(20\d{2}(?:08|01|05))', dir_name)
    if m_dir_term:
        dir_term = m_dir_term.group(1)
    else:
        m_dir_sem = re.search(r'(\d{2})(fall|spring)', dir_name)
        if m_dir_sem:
            y = 2000 + int(m_dir_sem.group(1))
            sem = m_dir_sem.group(2)
            dir_term = f"{y}08" if sem == "fall" else f"{y}01"

    # 3. Determine active term
    ground_truth_term = sorted(list(raw_terms))[0] if raw_terms else dir_term

    plan_term = str(plan.get("term_code", "")).strip()
    plan_sem = str(plan.get("semester", "")).strip().lower()

    if ground_truth_term:
        if plan_term and plan_term != ground_truth_term:
            print(f"\nWARNING: assessment_plan.json term_code '{plan_term}' conflicts with raw data term '{ground_truth_term}'!")
            print(f"Using raw data term '{ground_truth_term}' for accurate report generation.\n")
            active_term = ground_truth_term
        elif plan_sem and (("fall" in plan_sem and ground_truth_term.endswith("01")) or ("spring" in plan_sem and ground_truth_term.endswith("08"))):
            print(f"\nWARNING: assessment_plan.json semester '{plan.get('semester')}' conflicts with raw data term '{ground_truth_term}'!")
            print(f"Using raw data term '{ground_truth_term}' for accurate report generation.\n")
            active_term = ground_truth_term
        else:
            active_term = ground_truth_term
    else:
        active_term = plan_term or "202601"

    # 4. Resolve Academic Year (supports "academic_year": "25-26" or "2025-2026")
    ay_val = plan.get("academic_year")
    fall_term, spring_term = "", ""
    if ay_val:
        m_ay = re.search(r'(\d{2,4})\s*[-/]\s*(\d{2,4})', str(ay_val).strip())
        if m_ay:
            y1 = int(m_ay.group(1)) if len(m_ay.group(1)) == 4 else 2000 + int(m_ay.group(1))
            y2 = int(m_ay.group(2)) if len(m_ay.group(2)) == 4 else 2000 + int(m_ay.group(2))
            fall_term = f"{y1}08"
            spring_term = f"{y2}01"

    if not fall_term or not spring_term:
        m_dir_ay = re.search(r'(\d{2,4})\s*[-/]\s*(\d{2,4})', dir_name)
        if m_dir_ay:
            y1 = int(m_dir_ay.group(1)) if len(m_dir_ay.group(1)) == 4 else 2000 + int(m_dir_ay.group(1))
            y2 = int(m_dir_ay.group(2)) if len(m_dir_ay.group(2)) == 4 else 2000 + int(m_dir_ay.group(2))
            fall_term = f"{y1}08"
            spring_term = f"{y2}01"

    if not fall_term or not spring_term:
        if active_term.endswith("08") or "fall" in plan_sem or (dir_term and dir_term.endswith("08")):
            y = int(active_term[:4])
            fall_term = f"{y}08"
            spring_term = f"{y + 1}01"
        else:
            y = int(active_term[:4])
            fall_term = f"{y - 1}08"
            spring_term = f"{y}01"

    has_fall = any(t.endswith("08") for t in raw_terms)
    has_spring = any(t.endswith("01") for t in raw_terms)
    is_both = has_fall and has_spring

    if is_both:
        is_fall = True
        is_spring = True
        active_term = f"{fall_term} & {spring_term}"
    elif has_fall:
        is_fall = True
        is_spring = False
        active_term = fall_term
    elif has_spring:
        is_fall = False
        is_spring = True
        active_term = spring_term
    else:
        is_fall = bool(active_term.endswith("08") or "fall" in plan_sem or (dir_term and dir_term.endswith("08")))
        is_spring = not is_fall
        has_fall = is_fall
        has_spring = is_spring

    return {
        "course_code": course_code,
        "fall_term": fall_term,
        "spring_term": spring_term,
        "active_term": active_term,
        "is_fall": is_fall,
        "is_spring": is_spring,
        "has_fall": has_fall,
        "has_spring": has_spring,
        "is_both": is_both,
        "academic_year_label": f"{fall_term[:4]}-{spring_term[2:4]}"
    }


def build_sec_cols(sections: list) -> dict:
    """Computes column indices for side-by-side section tables in Excel."""
    sec_cols = {}
    for idx, sec_id in enumerate(sections):
        if idx == 0:
            c_std, c_scr, c_tot = 1, 3, 4
        elif idx == 1:
            c_std, c_scr, c_tot = 7, 8, 9
        else:
            c_std = 7 + (idx - 1) * 3
            c_scr = c_std + 1
            c_tot = c_std + 2
        sec_cols[sec_id] = {
            "std": c_std,
            "score": c_scr,
            "total": c_tot
        }
    return sec_cols


def export_academic_year_excel_report(course_dir: str, plan: dict, report_assessment_results: list, instructor_map: dict = None) -> list:
    """
    Generates standardized Academic Year Assessment Excel Workbooks (e.g. COP4020-CS-ONLY-Individual Data-LO2PI3.xlsx).
    Conforms to the 2-semester academic year format with side-by-side section columns,
    anonymous student IDs (S1, S2, ...), live formulas for counts/percentages, and outcome metadata.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Warning: 'openpyxl' module not found. Skipping .xlsx Excel report generation. Install via 'pip install openpyxl' or run with 'uv run'.")
        return []

    if instructor_map is None:
        instructor_map = load_instructor_mapping(course_dir)

    generated_paths = []
    term_info = resolve_course_and_academic_year(course_dir, plan)
    course_code = term_info["course_code"]
    fall_term = term_info["fall_term"]
    spring_term = term_info["spring_term"]
    is_fall = term_info["is_fall"]
    is_spring = term_info["is_spring"]

    title_text = f"{course_code} Assessment results {fall_term} and {spring_term}"

    font_base = Font(name="Arial", size=10)
    font_bold = Font(name="Arial", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    for res in report_assessment_results:
        cfg = res["cfg"]
        acc_type = cfg.get("accreditation_type", "ABET").upper()
        sheet_title = f"{acc_type} Assessment" if acc_type else "Assessment"
        code = str(res.get("code", "")).strip()

        # Sanitize code for filename: e.g. LO2.PI3 -> LO2PI3, LO.PI 1.1 -> LO1PI1
        filename_code = re.sub(r'[^A-Za-z0-9]', '', code) or "LO"

        # Display code: e.g. LO2-3, LO1-1
        nums = re.findall(r'\d+', code)
        if len(nums) >= 2:
            display_code = f"LO{nums[0]}-{nums[1]}"
        else:
            display_code = code

        # Major scope label: e.g. CS-ONLY, ALL-MAJORS
        target_majors = assessment_majors(plan, cfg)
        if not target_majors:
            major_scope = "ALL-MAJORS"
        elif len(target_majors) == 1:
            major_abbr = get_program_abbr(target_majors[0]).upper()
            major_scope = f"{major_abbr}-ONLY"
        else:
            abbrs = "-".join(get_program_abbr(m).upper() for m in target_majors)
            major_scope = abbrs

        filename = f"{course_code}-{major_scope}-Individual Data-{filename_code}.xlsx"
        output_xlsx_path = os.path.join(course_dir, filename)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        ws.views.sheetView[0].showGridLines = True

        # Group evaluated students by section
        eval_students = res.get("eval_students", [])
        sec_group = defaultdict(list)
        for r in eval_students:
            sec_group[str(r["section_id"])].append(r)

        # Sort sections by CRN
        sections = sorted(list(sec_group.keys()))
        if not sections:
            sections = ["Unknown"]

        # Partition sections into Fall and Spring
        sec_to_term = load_section_term_mapping(course_dir)
        fall_sections = []
        spring_sections = []
        for s in sections:
            t = sec_to_term.get(s, "")
            if t == fall_term or t.endswith("08"):
                fall_sections.append(s)
            elif t == spring_term or t.endswith("01"):
                spring_sections.append(s)
            else:
                s_terms = set(st.get("Term", "") for st in sec_group[s])
                if fall_term in s_terms or any(x.endswith("08") for x in s_terms):
                    fall_sections.append(s)
                elif spring_term in s_terms or any(x.endswith("01") for x in s_terms):
                    spring_sections.append(s)
                elif term_info.get("has_fall") and not term_info.get("has_spring"):
                    fall_sections.append(s)
                elif term_info.get("has_spring") and not term_info.get("has_fall"):
                    spring_sections.append(s)
                elif term_info.get("is_fall"):
                    fall_sections.append(s)
                else:
                    spring_sections.append(s)

        # Build column maps for each semester
        ph_fall = spring_sections[:2] if len(spring_sections) >= 2 else ["Unknown1", "Unknown2"]
        fall_cols = build_sec_cols(fall_sections if fall_sections else ph_fall)

        ph_spring = fall_sections[:2] if len(fall_sections) >= 2 else ["Unknown1", "Unknown2"]
        spring_cols = build_sec_cols(spring_sections if spring_sections else ph_spring)

        max_fall = max(c["total"] for c in fall_cols.values()) if fall_cols else 12
        max_spring = max(c["total"] for c in spring_cols.values()) if spring_cols else 12
        max_used_col = max(max_fall, max_spring, 17)

        # Row 1: Merged Banner title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_used_col)
        cell_a1 = ws.cell(1, 1, value=title_text)
        cell_a1.font = font_bold
        cell_a1.alignment = align_center

        # ==========================================
        # Fall Block (Rows 3 - 25+)
        # ==========================================
        if fall_sections:
            r_fa_inst = 3
            r_fa_crn = 4
            r_fa_lo = 5
            r_fa_std_hdr = 6
            r_fa_data_start = 7
            max_fa_students = max((len(sec_group[s]) for s in fall_sections), default=0)
            num_fa_student_rows = max(max_fa_students, 11)
            r_fa_data_end = r_fa_data_start + num_fa_student_rows - 1
            r_fa_total = r_fa_data_end + 2

            ws.cell(r_fa_lo, 1, value=float(fall_term) if fall_term.isdigit() else fall_term).font = font_bold
            ws.cell(r_fa_total - 1, 1, value="Course total").font = font_base
            ws.cell(r_fa_total - 1, 1).alignment = align_center
            ws.cell(r_fa_total, 1, value=f"Total {fall_term}").font = font_bold
            ws.cell(r_fa_total, 1).alignment = align_center

            for s_id in fall_sections:
                cols = fall_cols[s_id]
                inst_name = instructor_map.get(s_id, "")
                ws.cell(r_fa_inst, cols["score"], value=inst_name).font = font_bold
                ws.cell(r_fa_inst, cols["score"]).alignment = align_center

                crn_val = float(s_id) if s_id.isdigit() else s_id
                ws.cell(r_fa_crn, cols["score"], value=crn_val).font = font_base
                ws.cell(r_fa_crn, cols["score"]).alignment = align_center

                ws.cell(r_fa_lo, cols["score"], value=display_code).font = font_base
                ws.cell(r_fa_lo, cols["score"]).alignment = align_center

                ws.cell(r_fa_std_hdr, cols["std"], value="Student").font = font_base
                ws.cell(r_fa_std_hdr, cols["std"]).alignment = align_center
                if s_id == fall_sections[0]:
                    ws.cell(r_fa_std_hdr, cols["score"], value="Y/N").font = font_base
                    ws.cell(r_fa_std_hdr, cols["score"]).alignment = align_center

                s_list = sec_group[s_id]
                scr_col_letter = get_column_letter(cols["score"])
                for s_idx, st in enumerate(s_list):
                    r_curr = r_fa_data_start + s_idx
                    ws.cell(r_curr, cols["std"], value=st.get("Anonymous ID", f"S{s_idx+1}")).font = font_base
                    ws.cell(r_curr, cols["std"]).alignment = align_center

                    score_letter = "y" if st["is_pass"] else "n"
                    ws.cell(r_curr, cols["score"], value=score_letter).font = font_base
                    ws.cell(r_curr, cols["score"]).alignment = align_center

                ws.cell(r_fa_total, cols["score"], value=f'=countif({scr_col_letter}{r_fa_data_start}:{scr_col_letter}{r_fa_data_end},"=Y")').font = font_base
                ws.cell(r_fa_total, cols["score"]).alignment = align_center
                ws.cell(r_fa_total, cols["total"], value=f'=sum(countif({scr_col_letter}{r_fa_data_start}:{scr_col_letter}{r_fa_data_end},"Y"),countif({scr_col_letter}{r_fa_data_start}:{scr_col_letter}{r_fa_data_end},"N"))').font = font_base
                ws.cell(r_fa_total, cols["total"]).alignment = align_center

            fall_total_row = r_fa_total
        else:
            # Fall placeholder (unpopulated)
            ws.cell(5, 1, value=float(fall_term) if fall_term.isdigit() else fall_term).font = font_bold
            for s_id, cols in fall_cols.items():
                ws.cell(5, cols["score"], value=display_code).font = font_base
                ws.cell(5, cols["score"]).alignment = align_center
                ws.cell(6, cols["std"], value="Student").font = font_base
                ws.cell(6, cols["std"]).alignment = align_center
                if s_id == list(fall_cols.keys())[0]:
                    ws.cell(6, cols["score"], value="Y/N").font = font_base
                    ws.cell(6, cols["score"]).alignment = align_center

            ws.cell(24, 1, value="Course total").font = font_base
            ws.cell(24, 1).alignment = align_center
            ws.cell(25, 1, value=f"Total {fall_term}").font = font_bold
            ws.cell(25, 1).alignment = align_center

            for s_id, cols in fall_cols.items():
                scr_col_letter = get_column_letter(cols["score"])
                ws.cell(25, cols["score"], value=f'=countif({scr_col_letter}7:{scr_col_letter}23,"=Y")').font = font_base
                ws.cell(25, cols["score"]).alignment = align_center
                ws.cell(25, cols["total"], value=f'=sum(countif({scr_col_letter}7:{scr_col_letter}22,"Y"),countif({scr_col_letter}7:{scr_col_letter}22,"N"))').font = font_base
                ws.cell(25, cols["total"]).alignment = align_center

            fall_total_row = 25

        # ==========================================
        # Spring Block (Starts below Fall Total)
        # ==========================================
        r_sp_start = fall_total_row + 2
        if spring_sections:
            r_sp_inst = r_sp_start
            r_sp_crn = r_sp_start + 1
            r_sp_lo = r_sp_start + 2
            r_sp_std_hdr = r_sp_start + 3
            r_sp_data_start = r_sp_start + 4
            max_sp_students = max((len(sec_group[s]) for s in spring_sections), default=0)
            num_sp_student_rows = max(max_sp_students, 11)
            r_sp_data_end = r_sp_data_start + num_sp_student_rows - 1
            r_sp_total = r_sp_data_end + 3

            ws.cell(r_sp_inst, 16, value="Edit if professors or # of sections change").font = font_base
            ws.cell(r_sp_lo, 1, value=float(spring_term) if spring_term.isdigit() else spring_term).font = font_bold
            ws.cell(r_sp_total, 1, value=f"Total {spring_term}").font = font_bold
            ws.cell(r_sp_total, 1).alignment = align_center

            for s_id in spring_sections:
                cols = spring_cols[s_id]
                inst_name = instructor_map.get(s_id, "")
                ws.cell(r_sp_inst, cols["score"], value=inst_name).font = font_bold
                ws.cell(r_sp_inst, cols["score"]).alignment = align_center

                crn_val = float(s_id) if s_id.isdigit() else s_id
                ws.cell(r_sp_crn, cols["score"], value=crn_val).font = font_base
                ws.cell(r_sp_crn, cols["score"]).alignment = align_center

                ws.cell(r_sp_lo, cols["score"], value=display_code).font = font_base
                ws.cell(r_sp_lo, cols["score"]).alignment = align_center

                ws.cell(r_sp_std_hdr, cols["std"], value="Student").font = font_base
                ws.cell(r_sp_std_hdr, cols["std"]).alignment = align_center
                if s_id == spring_sections[0]:
                    ws.cell(r_sp_std_hdr, cols["score"], value="Y/N").font = font_base
                    ws.cell(r_sp_std_hdr, cols["score"]).alignment = align_center

                s_list = sec_group[s_id]
                scr_col_letter = get_column_letter(cols["score"])
                for s_idx, st in enumerate(s_list):
                    r_curr = r_sp_data_start + s_idx
                    ws.cell(r_curr, cols["std"], value=st.get("Anonymous ID", f"S{s_idx+1}")).font = font_base
                    ws.cell(r_curr, cols["std"]).alignment = align_center

                    score_letter = "y" if st["is_pass"] else "n"
                    ws.cell(r_curr, cols["score"], value=score_letter).font = font_base
                    ws.cell(r_curr, cols["score"]).alignment = align_center

                ws.cell(r_sp_total, cols["score"], value=f'=countif({scr_col_letter}{r_sp_data_start}:{scr_col_letter}{r_sp_data_end},"=Y")').font = font_base
                ws.cell(r_sp_total, cols["score"]).alignment = align_center
                ws.cell(r_sp_total, cols["total"], value=f'=sum(countif({scr_col_letter}{r_sp_data_start}:{scr_col_letter}{r_sp_data_end},"Y"),countif({scr_col_letter}{r_sp_data_start}:{scr_col_letter}{r_sp_data_end},"N"))').font = font_base
                ws.cell(r_sp_total, cols["total"]).alignment = align_center

            spring_total_row = r_sp_total
        else:
            # Spring placeholder (unpopulated)
            ws.cell(r_sp_start, 16, value="Edit if professors or # of sections change").font = font_base
            ws.cell(r_sp_start + 2, 1, value=float(spring_term) if spring_term.isdigit() else spring_term).font = font_bold
            spring_total_row = r_sp_start + 15
            ws.cell(spring_total_row, 1, value=f"Total {spring_term}").font = font_bold
            ws.cell(spring_total_row, 1).alignment = align_center

            for s_id, cols in spring_cols.items():
                ws.cell(r_sp_start + 2, cols["score"], value=display_code).font = font_base
                ws.cell(r_sp_start + 2, cols["score"]).alignment = align_center
                ws.cell(r_sp_start + 3, cols["std"], value="Student").font = font_base
                ws.cell(r_sp_start + 3, cols["std"]).alignment = align_center
                if s_id == list(spring_cols.keys())[0]:
                    ws.cell(r_sp_start + 3, cols["score"], value="Y/N").font = font_base
                    ws.cell(r_sp_start + 3, cols["score"]).alignment = align_center

                scr_col_letter = get_column_letter(cols["score"])
                ws.cell(spring_total_row, cols["score"], value=f'=countif({scr_col_letter}{r_sp_start+4}:{scr_col_letter}{spring_total_row-3},"=Y")').font = font_base
                ws.cell(spring_total_row, cols["score"]).alignment = align_center
                ws.cell(spring_total_row, cols["total"], value=f'=sum(countif({scr_col_letter}{r_sp_start+4}:{scr_col_letter}{spring_total_row-3},"Y"),countif({scr_col_letter}{r_sp_start+4}:{scr_col_letter}{spring_total_row-3},"N"))').font = font_base
                ws.cell(spring_total_row, cols["total"]).alignment = align_center

        # ==========================================
        # Summary Block below Spring Total
        # ==========================================
        r_sum_hdr = spring_total_row + 2
        ws.cell(r_sum_hdr, 3, value="Fall").font = font_bold
        ws.cell(r_sum_hdr, 8, value="Spring").font = font_bold

        r_sem_tot = r_sum_hdr + 1
        ws.cell(r_sem_tot, 1, value="Semester Totals").font = font_bold
        ws.cell(r_sem_tot, 1).alignment = align_center

        fall_active_cols = [fall_cols[s] for s in fall_sections] if fall_sections else list(fall_cols.values())
        fall_scr_sum = "+".join(f"{get_column_letter(c['score'])}{fall_total_row}" for c in fall_active_cols)
        fall_tot_sum = "+".join(f"{get_column_letter(c['total'])}{fall_total_row}" for c in fall_active_cols)
        ws.cell(r_sem_tot, 3, value=f"={fall_scr_sum}").font = font_base
        ws.cell(r_sem_tot, 3).alignment = align_center
        ws.cell(r_sem_tot, 4, value=f"={fall_tot_sum}").font = font_base
        ws.cell(r_sem_tot, 4).alignment = align_center

        spring_active_cols = [spring_cols[s] for s in spring_sections] if spring_sections else list(spring_cols.values())
        sp_scr_sum = "+".join(f"{get_column_letter(c['score'])}{spring_total_row}" for c in spring_active_cols)
        sp_tot_sum = "+".join(f"{get_column_letter(c['total'])}{spring_total_row}" for c in spring_active_cols)
        ws.cell(r_sem_tot, 8, value=f"={sp_scr_sum}").font = font_base
        ws.cell(r_sem_tot, 8).alignment = align_center
        ws.cell(r_sem_tot, 9, value=f"={sp_tot_sum}").font = font_base
        ws.cell(r_sem_tot, 9).alignment = align_center

        r_gr_tot = r_sem_tot + 2
        ws.cell(r_gr_tot, 1, value="Grand Totals").font = font_bold
        ws.cell(r_gr_tot, 1).alignment = align_center
        ws.cell(r_gr_tot, 3, value=f"=C{r_sem_tot}+H{r_sem_tot}").font = font_base
        ws.cell(r_gr_tot, 3).alignment = align_center
        ws.cell(r_gr_tot, 4, value=f"=D{r_sem_tot}+I{r_sem_tot}").font = font_base
        ws.cell(r_gr_tot, 4).alignment = align_center

        r_pct = r_gr_tot + 1
        ws.cell(r_pct, 1, value="Percentage").font = font_bold
        ws.cell(r_pct, 1).alignment = align_center
        ws.cell(r_pct, 3, value=f"=IF(D{r_gr_tot}>0,ROUND(C{r_gr_tot}/D{r_gr_tot}*100,0),0)").font = font_base
        ws.cell(r_pct, 3).alignment = align_center

        r_met = r_pct + 2
        ws.cell(r_met, 1, value="Met PI?").font = font_bold
        ws.cell(r_met, 1).alignment = align_center
        target_pct = int(cfg.get("target_pass_rate_percent", 70))
        ws.cell(r_met, 3, value=f'=IF(C{r_pct}>={target_pct}, "Y", "N")').font = font_bold
        ws.cell(r_met, 3).alignment = align_center
        ws.cell(r_met, 7, value=f"Percentages > {target_pct} means the indicator has been met.").font = font_base
        ws.cell(r_met, 7).alignment = align_left

        r_desc = r_met + 2
        ws.cell(r_desc, 1, value=code).font = font_bold
        ws.cell(r_desc, 1).alignment = align_center
        desc_text = cfg.get("notes") or cfg.get("assessment_name") or res.get("full_assessment_name", "")
        ws.merge_cells(start_row=r_desc, start_column=2, end_row=r_desc, end_column=max_used_col)
        ws.cell(r_desc, 2, value=desc_text).font = font_bold
        ws.cell(r_desc, 2).alignment = align_left

        # Set standard column widths matching reference layout
        col_width_defaults = {
            1: 14.75, 2: 9.63, 3: 5.75, 4: 4.38, 5: 10.75, 6: 9.63,
            7: 11.25, 8: 5.88, 9: 5.38, 10: 10.25, 11: 6.25, 12: 5.38,
            13: 8.5, 14: 6.0, 15: 4.5, 16: 9.75, 17: 8.5
        }
        for col_idx in range(1, max_used_col + 1):
            letter = get_column_letter(col_idx)
            w = col_width_defaults.get(col_idx, 8.5)
            ws.column_dimensions[letter].width = w

        wb.save(output_xlsx_path)
        print(f"Generated Academic Year Assessment Excel Report: {output_xlsx_path}")
        generated_paths.append(output_xlsx_path)

    return generated_paths


def export_excel_report(course_dir: str, plan: dict, report_assessment_results: list, per_student_export_rows: list = None, total_enrolled: int = 0, active_count: int = 0, withdrawn_count: int = 0, instructor_map: dict = None) -> str:
    """Wrapper calling export_academic_year_excel_report for compatibility."""
    if instructor_map is None:
        instructor_map = load_instructor_mapping(course_dir)
    paths = export_academic_year_excel_report(course_dir, plan, report_assessment_results, instructor_map)
    return paths[0] if paths else ""


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 -m semester_data_processor.assessments <course_dir>")
        sys.exit(1)

    extract_assessments(sys.argv[1])
