"""
rosters.py - Merges section rosters in <course_dir>/raw/ into a single combined roster CSV in <course_dir>/.
"""

import sys
import os
import glob
import re
import csv
from .readme import generate_output_readme


def process_rosters(course_dir: str) -> str:
    raw_dir = os.path.join(course_dir, "raw")
    if not os.path.exists(raw_dir):
        raise FileNotFoundError(f"Raw directory not found: {raw_dir}")

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
        "UWFID", "Last Name", "First Name", "Preferred Name", "Full Name",
        "Email", "Section Number", "Major", "Level", "Program", "Student Type",
        "Registered", "Withdrawn", "Status", "Cohort"
    ]

    for filepath in roster_files:
        filename = os.path.basename(filepath)
        print(f"  - Reading {filename}")

        # Extract 5-digit section CRN from filename (e.g., CourseRoster_202601_18451 (1).csv -> 18451)
        section_match = re.search(r'_(\d{5})\b', filename) or re.search(r'_(\d+)', filename)
        section_number = section_match.group(1) if section_match else "Unknown"

        with open(filepath, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames:
                reader.fieldnames = [fn.strip() for fn in reader.fieldnames]

            for row in reader:
                cleaned_row = {k.strip(): (v.strip() if v else "") for k, v in row.items()}
                
                # Explicit Section Number column
                cleaned_row["Section Number"] = section_number
                if "Section" not in cleaned_row or not cleaned_row["Section"]:
                    cleaned_row["Section"] = section_number

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

                # Deduplicate by UWFID + Section Number
                uwfid = cleaned_row.get("UWFID", "")
                dedup_key = (uwfid, section_number)
                if dedup_key not in seen_keys:
                    seen_keys.add(dedup_key)
                    all_students.append(cleaned_row)

    # Sort students by Last Name, First Name
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
