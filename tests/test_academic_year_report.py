"""
tests/test_academic_year_report.py - Unit tests for Academic Year Assessment Excel Report & Anonymization.
"""

import os
import json
import csv
import shutil
import tempfile
import unittest
from pathlib import Path
import openpyxl

from semester_data_processor import process_rosters, process_gradebook, extract_assessments
from semester_data_processor.rosters import parse_sections_text, load_instructor_mapping


class TestAcademicYearReport(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.raw_dir = os.path.join(self.test_dir, "raw")
        os.makedirs(self.raw_dir, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_parse_sections_text(self):
        """Test parsing sections.txt with instructor and CRN mapping."""
        sec_file = os.path.join(self.raw_dir, "202601-sections.txt")
        with open(sec_file, "w", encoding="utf-8") as f:
            f.write("Fang 16197 17089\n")
            f.write("Zhang 18451 18626\n")

        mapping = parse_sections_text(sec_file)
        self.assertEqual(mapping["16197"], "Fang")
        self.assertEqual(mapping["17089"], "Fang")
        self.assertEqual(mapping["18451"], "Zhang")
        self.assertEqual(mapping["18626"], "Zhang")

        loaded = load_instructor_mapping(self.test_dir)
        self.assertEqual(loaded, mapping)

    def test_academic_year_excel_and_anonymization(self):
        """Test generating Academic Year Assessment Excel report with anonymized student IDs (S1, S2...)."""
        # 1. Create sections text file
        sec_file = os.path.join(self.raw_dir, "202601-sections.txt")
        with open(sec_file, "w", encoding="utf-8") as f:
            f.write("Fang 16197\n")
            f.write("Zhang 18451\n")

        # 2. Create raw roster files
        r1 = os.path.join(self.raw_dir, "CourseRoster_202601_16197.csv")
        with open(r1, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("970111111, Alice, Smith, as1@students.uwf.edu, Undergraduate, Computer Science\n")
            f.write("970222222, Bob, Jones, bj2@students.uwf.edu, Undergraduate, Computer Science\n")

        r2 = os.path.join(self.raw_dir, "CourseRoster_202601_18451.csv")
        with open(r2, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("970333333, Charlie, Brown, cb3@students.uwf.edu, Undergraduate, Computer Science\n")
            f.write("970444444, David, Clark, dc4@students.uwf.edu, Undergraduate, Software Design & Development\n")

        # 3. Create raw gradebook file
        gb = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4020.csv")
        with open(gb, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1270399)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Jones, Bob",80001,970222222,bj2,16197_COP4020,85.00\n')
            f.write('"Smith, Alice",80002,970111111,as1,16197_COP4020,95.00\n')
            f.write('"Brown, Charlie",80003,970333333,cb3,18451_COP4020,60.00\n')
            f.write('"Clark, David",80004,970444444,dc4,18451_COP4020,90.00\n')

        # 4. Create plan for SACSCOC / general accreditation
        plan = {
            "course_name": "COP 4020 - Programming Languages",
            "semester": "Spring 2026",
            "term_code": "202601",
            "target_majors": ["Computer Science"],
            "assessments": [
                {
                    "accreditation_type": "SACSCOC",
                    "code": "LO2.PI3",
                    "assessment_name": "Project 1",
                    "target_majors": ["Computer Science"],
                    "passing_operator": ">=",
                    "passing_threshold_percent": 70.0,
                    "target_pass_rate_percent": 70.0,
                    "notes": "Demonstrate the ability to properly test the solution"
                }
            ]
        }
        with open(os.path.join(self.test_dir, "assessment_plan.json"), "w", encoding="utf-8") as f:
            json.dump(plan, f, indent=2)

        # Run pipeline
        process_rosters(self.test_dir)
        process_gradebook(self.test_dir)
        extract_assessments(self.test_dir)

        # Verify combined_roster.csv
        roster_rows = list(csv.DictReader(open(os.path.join(self.test_dir, "combined_roster.csv"))))
        self.assertIn("Anonymous ID", roster_rows[0])
        self.assertIn("Instructor", roster_rows[0])

        # Verify assessment_results.csv mapping
        results_rows = list(csv.DictReader(open(os.path.join(self.test_dir, "assessment_results.csv"))))
        self.assertIn("Anonymous ID", results_rows[0])
        self.assertIn("Instructor", results_rows[0])
        alice_rec = next(r for r in results_rows if r["First Name"] == "Alice")
        bob_rec = next(r for r in results_rows if r["First Name"] == "Bob")
        charlie_rec = next(r for r in results_rows if r["First Name"] == "Charlie")
        # Alice and Bob in section 16197: Jones (Bob) is S1, Smith (Alice) is S2
        self.assertEqual(bob_rec["Anonymous ID"], "S1")
        self.assertEqual(alice_rec["Anonymous ID"], "S2")
        self.assertEqual(bob_rec["Instructor"], "Fang")
        # Charlie in section 18451: Brown (Charlie) is S1
        self.assertEqual(charlie_rec["Anonymous ID"], "S1")
        self.assertEqual(charlie_rec["Instructor"], "Zhang")

        # Verify generated Excel report
        expected_filename = "COP4020-CS-ONLY-Individual Data-LO2PI3.xlsx"
        xlsx_path = os.path.join(self.test_dir, expected_filename)
        self.assertTrue(os.path.exists(xlsx_path), f"File {expected_filename} not found!")

        wb = openpyxl.load_workbook(xlsx_path, data_only=False)
        self.assertEqual(wb.sheetnames, ["SACSCOC Assessment"])
        ws = wb.active

        # Check Banner Title
        self.assertEqual(ws.cell(1, 1).value, "COP4020 Assessment results 202508 and 202601")

        # Check Spring 202601 row
        self.assertEqual(ws.cell(29, 1).value, 202601.0)
        # Check Section 16197 header (Fang at Col C, 16197 at Col C, LO2-3 at Col C)
        self.assertEqual(ws.cell(27, 3).value, "Fang")
        self.assertEqual(ws.cell(28, 3).value, 16197.0)
        self.assertEqual(ws.cell(29, 3).value, "LO2-3")

        # Check Section 18451 header (Zhang at Col H, 18451 at Col H, LO2-3 at Col H)
        self.assertEqual(ws.cell(27, 8).value, "Zhang")
        self.assertEqual(ws.cell(28, 8).value, 18451.0)

        # Check anonymous student rows
        self.assertEqual(ws.cell(31, 1).value, "S1")
        self.assertEqual(ws.cell(31, 3).value, "y")
        self.assertEqual(ws.cell(32, 1).value, "S2")
        self.assertEqual(ws.cell(32, 3).value, "y")

        # Check Charlie in section 18451 (Col G=S1, Col H=n)
        self.assertEqual(ws.cell(31, 7).value, "S1")
        self.assertEqual(ws.cell(31, 8).value, "n")

        # Verify FERPA compliance: real names and SIS IDs must NOT be anywhere in the sheet
        all_cell_values = [str(cell.value) for row in ws.rows for cell in row if cell.value is not None]
        for sensitive in ["Smith", "Alice", "Jones", "Bob", "Charlie", "Brown", "970111111", "970222222", "970333333"]:
            self.assertNotIn(sensitive, all_cell_values)

        wb.close()

    def test_academic_year_25_26_format(self):
        """Test academic_year: '25-26' format without semester or term_code."""
        from semester_data_processor.assessments import resolve_course_and_academic_year

        plan = {
            "course_name": "COP 4534 - DSA 2",
            "academic_year": "25-26",
            "target_majors": ["Computer Science"],
            "assessments": []
        }

        # Simulate a Fall directory with raw roster
        fall_dir = tempfile.mkdtemp(prefix="cop4534-25-26_")
        try:
            raw_fall = os.path.join(fall_dir, "raw")
            os.makedirs(raw_fall, exist_ok=True)
            with open(os.path.join(raw_fall, "CourseRoster_202508_86053.csv"), "w") as f:
                f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            with open(os.path.join(raw_fall, "202508-sections.txt"), "w") as f:
                f.write("Fang 86053\n")

            info = resolve_course_and_academic_year(fall_dir, plan)
            self.assertEqual(info["course_code"], "COP4534")
            self.assertEqual(info["fall_term"], "202508")
            self.assertEqual(info["spring_term"], "202601")
            self.assertEqual(info["active_term"], "202508")
            self.assertTrue(info["is_fall"])
            self.assertFalse(info["is_spring"])
            self.assertEqual(info["academic_year_label"], "2025-26")
        finally:
            shutil.rmtree(fall_dir)


if __name__ == "__main__":
    unittest.main()
