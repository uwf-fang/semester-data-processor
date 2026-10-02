"""
gradebook.py - Cleans Canvas gradebook export in <course_dir>/raw/ and merges with combined_roster.csv.
"""

import sys
import os
import glob
import csv
import re
import math
from .rosters import process_rosters
from .readme import generate_output_readme


def discover_gradebooks(raw_dir: str) -> list:
    """Discover all gradebook exports, including mixed filename capitalization."""
    return sorted(path for path in glob.glob(os.path.join(raw_dir, "*"))
                  if os.path.isfile(path) and path.lower().endswith(".csv")
                  and "grade" in os.path.basename(path).lower())


def read_gradebooks(raw_dir: str):
    """Return union headers, unique students, and points keyed by (SIS ID, column).

    Assignment IDs differ between Canvas courses; equal cleaned titles align.
    Ambiguous titles within an export and conflicting overlapping records fail
    before writing outputs, instead of silently selecting a score or denominator.
    """
    files = discover_gradebooks(raw_dir)
    if not files:
        raise FileNotFoundError(f"No gradebook CSV files found in {raw_dir}")
    headers, students, points, sources = [], {}, {}, {}
    for path in files:
        name = os.path.basename(path)
        print(f"Processing gradebook file: {name}")
        with open(path, encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.reader(stream))
        if not rows:
            raise ValueError(f"Empty gradebook: {name}")
        cols = ["Canvas Section" if h.strip() == "Section" else
                re.sub(r"\s*\(\d+\)$", "", h.strip()) for h in rows[0]]
        if len(cols) != len(set(cols)):
            raise ValueError(f"Duplicate assignment titles after removing Canvas IDs in {name}")
        if not {"Student", "SIS User ID"}.issubset(cols):
            raise ValueError(f"Missing Student or SIS User ID column in {name}")
        pts_idx = next((i for i, row in enumerate(rows[1:], 1)
                        if row and row[0].strip() == "Points Possible"), None)
        if pts_idx is None:
            raise ValueError(f"Missing Points Possible row in {name}")
        maxima = {}
        for col, value in zip(cols, rows[pts_idx]):
            try:
                number = float(value.strip())
                maxima[col] = number if math.isfinite(number) else None
            except ValueError:
                maxima[col] = None
        headers.extend(c for c in cols if c not in headers)
        for row in rows[pts_idx + 1:]:
            record = dict(zip(cols, (v.strip() for v in row)))
            sid = record.get("SIS User ID", "")
            if not sid or "Test Student" in record.get("Student", ""):
                continue
            existing = students.setdefault(sid, {})
            for col in cols:
                value, old = record.get(col, ""), existing.get(col, "")
                key = (sid, col)
                if old and value and old != value:
                    try:
                        equal = float(old) == float(value)
                    except ValueError:
                        equal = False
                    if not equal:
                        raise ValueError(f"Conflicting gradebook values for SIS ID {sid}, "
                                         f"column {col!r}: {sources.get(key)} and {name}")
                if value:
                    maximum = maxima.get(col)
                    if old and points.get(key) != maximum:
                        raise ValueError(f"Conflicting Points Possible for SIS ID {sid}, "
                                         f"column {col!r}: {sources.get(key)} and {name}")
                    existing[col] = value
                    points[key] = maximum
                    sources[key] = name
                else:
                    existing.setdefault(col, "")
                    points.setdefault(key, maxima.get(col))
    return headers, list(students.values()), points


def process_gradebook(course_dir: str) -> str:
    raw_dir = os.path.join(course_dir, "raw")
    roster_path = os.path.join(course_dir, "combined_roster.csv")

    if not os.path.exists(roster_path):
        print(f"Combined roster not found at {roster_path}. Running process_rosters first...")
        process_rosters(course_dir)

    # Read combined roster into dictionary keyed by UWFID
    roster_map = {}
    roster_fieldnames = []
    with open(roster_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        roster_fieldnames = reader.fieldnames or []
        for row in reader:
            uwfid = row.get("UWFID", "").strip()
            sec = row.get("Section Number", "").strip()
            if uwfid:
                roster_map[uwfid] = row
                if sec:
                    roster_map[(uwfid, sec)] = row

    gb_headers, gradebook_records, _ = read_gradebooks(raw_dir)
    student_records = []
    seen_uwfids = set()
    for row_dict in gradebook_records:
        sis_id = row_dict["SIS User ID"]

        uwfid = sis_id
        seen_uwfids.add(uwfid)

        # Attempt to match section from Canvas Section (e.g. 18451_COP4020)
        csec = row_dict.get("Canvas Section", "")
        sec_m = re.search(r'\b(\d{5})\b', csec)
        sec_key = sec_m.group(1) if sec_m else ""

        merged_row = {}
        roster_info = roster_map.get((uwfid, sec_key), roster_map.get(uwfid, {}))

        # 1. Canvas identification fields first
        merged_row["Student"] = row_dict.get("Student", "")
        merged_row["ID"] = row_dict.get("ID", "")
        merged_row["SIS User ID"] = sis_id
        merged_row["SIS Login ID"] = row_dict.get("SIS Login ID", "")

        # 2. External roster attributes
        merged_row["Section Number"] = roster_info.get("Section Number", sec_key)
        merged_row["Instructor"] = roster_info.get("Instructor", "")
        merged_row["Anonymous ID"] = roster_info.get("Anonymous ID", "")
        merged_row["Term"] = roster_info.get("Term", "")
        merged_row["Major"] = roster_info.get("Major", roster_info.get("Program", ""))
        merged_row["Status"] = roster_info.get("Status", "Active")
        merged_row["Level"] = roster_info.get("Level", "")
        merged_row["Cohort"] = roster_info.get("Cohort", "")
        merged_row["Canvas Section"] = csec

        # 3. Include ALL gradebook columns from Canvas intact
        skip_initial = {"Student", "ID", "SIS User ID", "SIS Login ID", "Canvas Section"}
        for h in gb_headers:
            if h not in skip_initial:
                merged_row[h] = row_dict.get(h, "")

        for r_field in roster_fieldnames:
            if r_field not in merged_row:
                merged_row[r_field] = roster_info.get(r_field, "")

        student_records.append(merged_row)

    # Check for any roster students not in Canvas gradebook
    for key, roster_info in roster_map.items():
        if isinstance(key, tuple):
            continue
        uwfid = key
        if uwfid not in seen_uwfids:
            merged_row = {
                "Student": roster_info.get("Full Name", ""),
                "ID": "",
                "SIS User ID": uwfid,
                "SIS Login ID": "",
                "Section Number": roster_info.get("Section Number", ""),
                "Instructor": roster_info.get("Instructor", ""),
                "Anonymous ID": roster_info.get("Anonymous ID", ""),
                "Term": roster_info.get("Term", ""),
                "Major": roster_info.get("Major", roster_info.get("Program", "")),
                "Status": roster_info.get("Status", "Active"),
                "Level": roster_info.get("Level", ""),
                "Cohort": roster_info.get("Cohort", ""),
                "Canvas Section": ""
            }
            for r_field in roster_fieldnames:
                if r_field not in merged_row:
                    merged_row[r_field] = roster_info.get(r_field, "")
            student_records.append(merged_row)

    student_records.sort(key=lambda r: r.get("Student", r.get("Last Name", "")).lower())

    id_metadata_cols = ["Student", "ID", "SIS User ID", "SIS Login ID", "Section Number", "Instructor", "Anonymous ID", "Term", "Major", "Status", "Level", "Cohort", "Canvas Section"]
    
    all_keys = set()
    for rec in student_records:
        all_keys.update(rec.keys())

    final_fieldnames = [c for c in id_metadata_cols if c in all_keys or c in gb_headers]
    
    for h in gb_headers:
        if h not in final_fieldnames:
            final_fieldnames.append(h)

    remaining = [c for c in sorted(all_keys) if c not in final_fieldnames]
    final_fieldnames.extend(remaining)

    output_path = os.path.join(course_dir, "merged_gradebook.csv")
    with open(output_path, mode="w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=final_fieldnames)
        writer.writeheader()
        for rec in student_records:
            row_to_write = {col: rec.get(col, "") for col in final_fieldnames}
            writer.writerow(row_to_write)

    print(f"\nSuccessfully generated: {output_path}")
    print(f"Total Merged Student Records: {len(student_records)}")

    generate_output_readme(course_dir, action_name="Process Gradebook")

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 -m semester_data_processor.gradebook <course_dir>")
        sys.exit(1)

    process_gradebook(sys.argv[1])
