"""
gradebook.py - Cleans Canvas gradebook export in <course_dir>/raw/ and merges with combined_roster.csv.
"""

import sys
import os
import glob
import csv
from .rosters import process_rosters


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
            if uwfid:
                roster_map[uwfid] = row

    # Find Canvas gradebook export file
    gradebook_files = sorted(glob.glob(os.path.join(raw_dir, "*Grades*.csv")))
    if not gradebook_files:
        gradebook_files = sorted(glob.glob(os.path.join(raw_dir, "*[Gg]rade*.csv")))

    if not gradebook_files:
        raise FileNotFoundError(f"No gradebook CSV files found in {raw_dir}")

    gradebook_file = gradebook_files[0]
    print(f"Processing gradebook file: {os.path.basename(gradebook_file)}")

    with open(gradebook_file, mode="r", encoding="utf-8-sig") as f:
        reader = list(csv.reader(f))

    if len(reader) < 2:
        raise ValueError("Gradebook CSV file must contain at least Header and Points Possible rows.")

    import re

    raw_headers = [col.strip() for col in reader[0]]

    gb_headers = []
    for h in raw_headers:
        if h == "Section":
            gb_headers.append("Canvas Section")
        else:
            gb_headers.append(re.sub(r'\s*\(\d+\)$', '', h))

    # Dynamically find Points Possible row index and student row start index
    pts_possible_idx = 1
    for idx, r in enumerate(reader[1:], 1):
        r_str = " ".join(r)
        if "Points Possible" in r_str or (r and "Points Possible" in r[0]):
            pts_possible_idx = idx
            break

    student_start_idx = pts_possible_idx + 1

    student_records = []
    seen_uwfids = set()

    for row in reader[student_start_idx:]:
        if not row or not any(row):
            continue

        row_dict = dict(zip(gb_headers, [val.strip() for val in row]))
        student_name = row_dict.get("Student", "")
        sis_id = row_dict.get("SIS User ID", "").strip()

        if "Points Possible" in student_name or "Test Student" in student_name or not sis_id:
            continue

        uwfid = sis_id
        seen_uwfids.add(uwfid)

        merged_row = {}
        roster_info = roster_map.get(uwfid, {})

        # 1. Canvas identification fields first
        merged_row["Student"] = row_dict.get("Student", "")
        merged_row["ID"] = row_dict.get("ID", "")
        merged_row["SIS User ID"] = sis_id
        merged_row["SIS Login ID"] = row_dict.get("SIS Login ID", "")

        # 2. External roster attributes
        merged_row["Section Number"] = roster_info.get("Section Number", "")
        merged_row["Major"] = roster_info.get("Major", roster_info.get("Program", ""))
        merged_row["Status"] = roster_info.get("Status", "Active")
        merged_row["Level"] = roster_info.get("Level", "")
        merged_row["Cohort"] = roster_info.get("Cohort", "")
        merged_row["Canvas Section"] = row_dict.get("Canvas Section", "")

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
    for uwfid, roster_info in roster_map.items():
        if uwfid not in seen_uwfids:
            merged_row = {
                "Student": roster_info.get("Full Name", ""),
                "ID": "",
                "SIS User ID": uwfid,
                "SIS Login ID": "",
                "Section Number": roster_info.get("Section Number", ""),
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

    id_metadata_cols = ["Student", "ID", "SIS User ID", "SIS Login ID", "Section Number", "Major", "Status", "Level", "Cohort", "Canvas Section"]
    
    all_keys = set()
    for rec in student_records:
        all_keys.update(rec.keys())

    final_fieldnames = [c for c in id_metadata_cols if c in all_keys]
    
    for h in gb_headers:
        if h not in final_fieldnames and h in all_keys:
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

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 -m semester_workflow.gradebook <course_dir>")
        sys.exit(1)

    process_gradebook(sys.argv[1])
