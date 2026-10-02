"""
rosters.py - Merges section rosters in <course_dir>/raw/ into a single combined roster CSV in <course_dir>/.
"""

import sys
import os
import glob
import re
import csv
import json
from .readme import generate_output_readme


def parse_sections_text(filepath: str) -> dict:
    """Parses text file mapping instructors to sections, e.g. 'Fang 16197 17089'."""
    sec_to_inst = {}
    with open(filepath, "r", encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if ":" in line:
                left, right = line.split(":", 1)
                right_tokens = re.split(r'[\s,]+', right.strip())
                if left.strip().isdigit():
                    sec_to_inst[left.strip()] = right.strip()
                else:
                    inst = left.strip()
                    for t in right_tokens:
                        if t.isdigit():
                            sec_to_inst[t] = inst
            else:
                tokens = re.split(r'[\s,]+', line)
                if not tokens:
                    continue
                if not tokens[0].isdigit():
                    inst = tokens[0]
                    for t in tokens[1:]:
                        if t.isdigit():
                            sec_to_inst[t] = inst
                else:
                    sec = tokens[0]
                    inst = " ".join(tokens[1:])
                    sec_to_inst[sec] = inst
    return sec_to_inst


def load_instructor_mapping(course_dir: str) -> dict:
    """Discovers and loads section to instructor mapping from plan or sections text files."""
    sec_to_inst = {}
    plan_paths = [
        os.path.join(course_dir, "assessment_plan.json"),
        os.path.join(course_dir, "raw", "assessment_plan.json")
    ]
    for p in plan_paths:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    cfg = data.get("instructors", {})
                    for k, v in cfg.items():
                        if isinstance(v, list):
                            for s in v:
                                sec_to_inst[str(s)] = str(k)
                        elif isinstance(v, str):
                            if k.isdigit():
                                sec_to_inst[str(k)] = v
                            elif v.isdigit():
                                sec_to_inst[v] = k
            except Exception:
                pass

    search_dirs = [os.path.join(course_dir, "raw"), course_dir]
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        txt_files = sorted(glob.glob(os.path.join(d, "*section*.txt")) + glob.glob(os.path.join(d, "*instructor*.txt")))
        for tf in txt_files:
            parsed = parse_sections_text(tf)
            for s, inst in parsed.items():
                if s not in sec_to_inst:
                    sec_to_inst[s] = inst

    return sec_to_inst


def load_section_term_mapping(course_dir: str) -> dict:
    """
    Extracts section CRN to term code mapping from raw roster filenames and section files.
    e.g. CourseRoster_202508_86053.csv -> {'86053': '202508'}
         202601-sections.txt -> {'16197': '202601', ...}
    """
    sec_to_term = {}
    search_dirs = [os.path.join(course_dir, "raw"), course_dir]
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        try:
            for fname in os.listdir(d):
                m_r = re.search(r'CourseRoster_(20\d{4})_(\d+)', fname, re.IGNORECASE)
                if m_r:
                    sec_to_term[m_r.group(2)] = m_r.group(1)
                m_s = re.search(r'(20\d{4})[-_]sections', fname, re.IGNORECASE)
                if m_s:
                    term = m_s.group(1)
                    parsed = parse_sections_text(os.path.join(d, fname))
                    for s in parsed:
                        sec_to_term[s] = term
                m_g = re.search(r'_(\d{5})_[A-Za-z0-9]+_(20\d{4})\.csv', fname, re.IGNORECASE)
                if m_g:
                    sec_to_term[m_g.group(1)] = m_g.group(2)
        except Exception:
            pass
    return sec_to_term


def process_rosters(course_dir: str) -> str:
    raw_dir = os.path.join(course_dir, "raw")
    if not os.path.exists(raw_dir):
        raise FileNotFoundError(f"Raw directory not found: {raw_dir}")

    instructor_map = load_instructor_mapping(course_dir)

    # Find roster files
    roster_files = sorted(glob.glob(os.path.join(raw_dir, "CourseRoster*.csv")))
    if not roster_files:
        roster_files = sorted(glob.glob(os.path.join(raw_dir, "*[Rr]oster*.csv")))

    if not roster_files:
        raise FileNotFoundError(f"No roster CSV files found in {raw_dir}")

    print(f"Found {len(roster_files)} roster file(s) in {raw_dir}:")

    all_students = []
    seen_keys = set()
    fieldnames_set = set()

    priority_cols = [
        "Anonymous ID", "Instructor", "Term", "UWFID", "Last Name", "First Name", "Preferred Name", "Full Name",
        "Email", "Section Number", "Major", "Level", "Program", "Student Type",
        "Registered", "Withdrawn", "Status", "Cohort"
    ]

    for filepath in roster_files:
        filename = os.path.basename(filepath)
        print(f"  - Reading {filename}")

        # Extract 6-digit term code and 5-digit section CRN from filename
        m_term = re.search(r'CourseRoster_(20\d{4})_', filename) or re.search(r'_(20\d{4})\b', filename)
        term_code = m_term.group(1) if m_term else ""

        section_match = re.search(r'_(\d{5})\b', filename) or re.search(r'_(\d+)', filename)
        section_number = section_match.group(1) if section_match else "Unknown"

        with open(filepath, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames:
                reader.fieldnames = [fn.strip() for fn in reader.fieldnames]

            for row in reader:
                cleaned_row = {k.strip(): (v.strip() if v else "") for k, v in row.items()}
                
                # Explicit Section Number and Term columns
                cleaned_row["Section Number"] = section_number
                if "Section" not in cleaned_row or not cleaned_row["Section"]:
                    cleaned_row["Section"] = section_number
                cleaned_row["Term"] = term_code

                # Map Instructor
                cleaned_row["Instructor"] = instructor_map.get(section_number, "")

                # Normalize Program / Major column
                program_val = cleaned_row.get("Program", "").strip()
                if not program_val or program_val.lower() in ["null", "none"]:
                    cleaned_row["Program"] = "Computer Science"
                    cleaned_row["Major"] = "Computer Science"
                else:
                    cleaned_row["Major"] = program_val

                # Normalize Level
                level_val = cleaned_row.get("Level", "").strip()
                if not level_val or level_val.lower() in ["null", "none"]:
                    cleaned_row["Level"] = "Undergraduate"

                # Compute Full Name
                last_name = cleaned_row.get("Last Name", "")
                first_name = cleaned_row.get("First Name", "")
                if last_name and first_name:
                    cleaned_row["Full Name"] = f"{last_name}, {first_name}"
                else:
                    cleaned_row["Full Name"] = last_name or first_name

                # Compute Status
                withdrawn = cleaned_row.get("Withdrawn", "").upper()
                cleaned_row["Status"] = "Withdrawn" if withdrawn in ['Y', 'YES', 'TRUE', '1'] else "Active"

                fieldnames_set.update(cleaned_row.keys())

                # Deduplicate by UWFID + Section Number + Term
                uwfid = cleaned_row.get("UWFID", "")
                dedup_key = (uwfid, section_number, term_code)
                if dedup_key not in seen_keys:
                    seen_keys.add(dedup_key)
                    all_students.append(cleaned_row)

    # Sort students by Term, Section Number, then Last Name, First Name
    all_students.sort(key=lambda s: (s.get("Term", ""), s.get("Section Number", ""), s.get("Last Name", "").lower(), s.get("First Name", "").lower()))

    # Assign Anonymous ID (S1, S2, ...) per section for active students
    sec_counters = {}
    for s in all_students:
        sec = s.get("Section Number", "Unknown")
        term = s.get("Term", "")
        sec_key = f"{term}_{sec}" if term else sec
        if s.get("Status") == "Active":
            c = sec_counters.get(sec_key, 0) + 1
            sec_counters[sec_key] = c
            s["Anonymous ID"] = f"S{c}"
        else:
            s["Anonymous ID"] = ""
    fieldnames_set.add("Anonymous ID")
    fieldnames_set.add("Instructor")
    fieldnames_set.add("Term")

    # Final overall sort by Last Name, First Name
    all_students.sort(key=lambda s: (s.get("Last Name", "").lower(), s.get("First Name", "").lower()))

    # Build final header order
    final_fieldnames = [c for c in priority_cols if c in fieldnames_set]
    extra_cols = [c for c in sorted(fieldnames_set) if c not in final_fieldnames]
    final_fieldnames.extend(extra_cols)

    output_path = os.path.join(course_dir, "combined_roster.csv")
    with open(output_path, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=final_fieldnames)
        writer.writeheader()
        for s in all_students:
            row_to_write = {col: s.get(col, "") for col in final_fieldnames}
            writer.writerow(row_to_write)

    total_students = len(all_students)
    active_students = sum(1 for s in all_students if s.get("Status") == "Active")
    withdrawn_students = sum(1 for s in all_students if s.get("Status") == "Withdrawn")

    print(f"\nSuccessfully generated: {output_path}")
    print(f"Total Students: {total_students} (Active: {active_students}, Withdrawn: {withdrawn_students})")

    # Section summary
    sections = sorted(list(set(s.get("Section Number", "Unknown") for s in all_students)))
    print("\nBreakdown by Section Number:")
    for sec in sections:
        sec_students = [s for s in all_students if s.get("Section Number") == sec]
        sec_active = sum(1 for s in sec_students if s.get("Status") == "Active")
        sec_withdrawn = sum(1 for s in sec_students if s.get("Status") == "Withdrawn")
        print(f"  - Section {sec}: {len(sec_students)} total ({sec_active} Active, {sec_withdrawn} Withdrawn)")

    generate_output_readme(course_dir, action_name="Process Rosters")

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 -m semester_data_processor.rosters <course_dir>")
        sys.exit(1)

    process_rosters(sys.argv[1])
