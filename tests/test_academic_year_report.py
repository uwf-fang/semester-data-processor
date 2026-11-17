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
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        self.test_dir = os.path.join(temp_dir.name, "course-data")
        self.raw_dir = os.path.join(self.test_dir, "raw")
        os.makedirs(self.raw_dir, exist_ok=True)

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
            f.write("100000001, Alice, Smith, fixture1@example.com, Undergraduate, Computer Science\n")
            f.write("100000002, Bob, Jones, fixture2@example.com, Undergraduate, Computer Science\n")

        r2 = os.path.join(self.raw_dir, "CourseRoster_202601_18451.csv")
        with open(r2, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000003, Charlie, Brown, fixture3@example.com, Undergraduate, Computer Science\n")
            f.write("100000004, David, Clark, fixture4@example.com, Undergraduate, Software Design & Development\n")

        # 3. Create raw gradebook file
        gb = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4020.csv")
        with open(gb, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1270399)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Jones, Bob",90001,100000002,fixture2,16197_COP4020,85.00\n')
            f.write('"Smith, Alice",90002,100000001,fixture1,16197_COP4020,95.00\n')
            f.write('"Brown, Charlie",90003,100000003,fixture3,18451_COP4020,60.00\n')
            f.write('"Clark, David",90004,100000004,fixture4,18451_COP4020,90.00\n')

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
        for sensitive in ["Smith", "Alice", "Jones", "Bob", "Charlie", "Brown", "100000001", "100000002", "100000003"]:
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

    def test_dual_semester_academic_year_report(self):
        """Test processing simultaneous Fall & Spring semester data in a single raw directory."""
        # 1. Fall semester setup (202508)
        sec_fa = os.path.join(self.raw_dir, "202508-sections.txt")
        with open(sec_fa, "w", encoding="utf-8") as f:
            f.write("Fang 86053\n")
            f.write("Smith 87575\n")

        r_fa1 = os.path.join(self.raw_dir, "CourseRoster_202508_86053.csv")
        with open(r_fa1, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000001, Alice, Adams, fixture5@example.com, Undergraduate, Computer Science\n")
            f.write("100000002, Bob, Baker, fixture6@example.com, Undergraduate, Computer Science\n")

        r_fa2 = os.path.join(self.raw_dir, "CourseRoster_202508_87575.csv")
        with open(r_fa2, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000003, Charlie, Clark, fixture7@example.com, Undergraduate, Computer Science\n")

        # 2. Spring semester setup (202601)
        sec_sp = os.path.join(self.raw_dir, "202601-sections.txt")
        with open(sec_sp, "w", encoding="utf-8") as f:
            f.write("Zhang 18451\n")
            f.write("Wang 18626\n")

        r_sp1 = os.path.join(self.raw_dir, "CourseRoster_202601_18451.csv")
        with open(r_sp1, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000004, David, Davis, fixture8@example.com, Undergraduate, Computer Science\n")
            f.write("100000005, Eve, Evans, fixture9@example.com, Undergraduate, Computer Science\n")

        r_sp2 = os.path.join(self.raw_dir, "CourseRoster_202601_18626.csv")
        with open(r_sp2, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000006, Frank, Franklin, fixture10@example.com, Undergraduate, Computer Science\n")

        # 3. Gradebook files for both semesters
        gb_fa = os.path.join(self.raw_dir, "2025-12-01_Grades-COP4534.csv")
        with open(gb_fa, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1001)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Adams, Alice",90005,100000001,fixture5,86053_COP4534,95.00\n')
            f.write('"Baker, Bob",90006,100000002,fixture6,86053_COP4534,55.00\n')
            f.write('"Clark, Charlie",90007,100000003,fixture7,87575_COP4534,80.00\n')

        gb_sp = os.path.join(self.raw_dir, "2026-05-01_Grades-COP4534.csv")
        with open(gb_sp, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (2001)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Davis, David",90008,100000004,fixture8,18451_COP4534,90.00\n')
            f.write('"Evans, Eve",90009,100000005,fixture9,18451_COP4534,85.00\n')
            f.write('"Franklin, Frank",90010,100000006,fixture10,18626_COP4534,40.00\n')

        # 4. Assessment Plan with academic_year: 25-26
        plan = {
            "course_name": "COP 4534 - Data Structures and Algorithms II",
            "academic_year": "25-26",
            "target_majors": ["Computer Science"],
            "assessments": [
                {
                    "accreditation_type": "ABET",
                    "code": "LO1.PI2",
                    "assessment_name": "Project 1",
                    "target_majors": ["Computer Science"],
                    "passing_operator": ">=",
                    "passing_threshold_percent": 70.0,
                    "target_pass_rate_percent": 70.0,
                    "notes": "Analyze algorithmic efficiency"
                }
            ]
        }
        with open(os.path.join(self.test_dir, "assessment_plan.json"), "w", encoding="utf-8") as f:
            json.dump(plan, f, indent=2)

        # Run pipeline
        process_rosters(self.test_dir)
        process_gradebook(self.test_dir)
        extract_assessments(self.test_dir)

        # Check combined_roster.csv
        roster_rows = list(csv.DictReader(open(os.path.join(self.test_dir, "combined_roster.csv"))))
        self.assertEqual(len(roster_rows), 6)
        terms = {r["Term"] for r in roster_rows}
        self.assertEqual(terms, {"202508", "202601"})

        # Check assessment report markdown
        report_md_path = os.path.join(self.test_dir, "assessment_report.md")
        self.assertTrue(os.path.exists(report_md_path))
        with open(report_md_path, "r", encoding="utf-8") as f:
            report_content = f.read()
        self.assertIn("Fall & Spring (202508 & 202601)", report_content)

        # Check Excel workbook
        xlsx_path = os.path.join(self.test_dir, "COP4534-CS-ONLY-Individual Data-LO1PI2.xlsx")
        self.assertTrue(os.path.exists(xlsx_path), f"File {xlsx_path} not found!")

        wb = openpyxl.load_workbook(xlsx_path, data_only=False)
        self.assertEqual(wb.sheetnames, ["ABET Assessment"])
        ws = wb.active

        # 1. Merged Banner
        self.assertEqual(ws.cell(1, 1).value, "COP4534 Assessment results 202508 and 202601")

        # 2. Fall Block
        self.assertEqual(ws.cell(5, 1).value, 202508.0)
        # Section 86053 (Fang) at Col C (3)
        self.assertEqual(ws.cell(3, 3).value, "Fang")
        self.assertEqual(ws.cell(4, 3).value, 86053.0)
        self.assertEqual(ws.cell(5, 3).value, "LO1-2")
        # Section 87575 (Smith) at Col H (8)
        self.assertEqual(ws.cell(3, 8).value, "Smith")
        self.assertEqual(ws.cell(4, 8).value, 87575.0)
        self.assertEqual(ws.cell(5, 8).value, "LO1-2")

        # Fall students in section 86053 (Alice=S1, Bob=S2)
        self.assertEqual(ws.cell(7, 1).value, "S1")
        self.assertEqual(ws.cell(7, 3).value, "y")
        self.assertEqual(ws.cell(8, 1).value, "S2")
        self.assertEqual(ws.cell(8, 3).value, "n")

        # Fall student in section 87575 (Charlie=S1 at Col G)
        self.assertEqual(ws.cell(7, 7).value, "S1")
        self.assertEqual(ws.cell(7, 8).value, "y")

        # Fall Total Row (11 rows min data range -> rows 7 to 17, total at row 19)
        fall_total_row = 19
        self.assertEqual(ws.cell(fall_total_row, 1).value, "Total 202508")
        self.assertEqual(ws.cell(fall_total_row, 3).value, '=countif(C7:C17,"=Y")')
        self.assertEqual(ws.cell(fall_total_row, 8).value, '=countif(H7:H17,"=Y")')

        # 3. Spring Block
        # Starts at fall_total_row + 2 = 21
        sp_start = 21
        self.assertEqual(ws.cell(sp_start + 2, 1).value, 202601.0)
        # Section 18451 (Zhang) at Col C
        self.assertEqual(ws.cell(sp_start, 3).value, "Zhang")
        self.assertEqual(ws.cell(sp_start + 1, 3).value, 18451.0)
        self.assertEqual(ws.cell(sp_start + 2, 3).value, "LO1-2")
        # Section 18626 (Wang) at Col H
        self.assertEqual(ws.cell(sp_start, 8).value, "Wang")
        self.assertEqual(ws.cell(sp_start + 1, 8).value, 18626.0)
        self.assertEqual(ws.cell(sp_start + 2, 8).value, "LO1-2")

        # Spring students in section 18451 (David=S1, Eve=S2)
        sp_data_start = sp_start + 4  # row 25
        self.assertEqual(ws.cell(sp_data_start, 1).value, "S1")
        self.assertEqual(ws.cell(sp_data_start, 3).value, "y")
        self.assertEqual(ws.cell(sp_data_start + 1, 1).value, "S2")
        self.assertEqual(ws.cell(sp_data_start + 1, 3).value, "y")

        # Spring student in section 18626 (Frank=S1 at Col G)
        self.assertEqual(ws.cell(sp_data_start, 7).value, "S1")
        self.assertEqual(ws.cell(sp_data_start, 8).value, "n")

        # Spring Total Row (sp_data_start=25, 11 rows -> rows 25 to 35, total at row 38)
        sp_total_row = sp_data_start + 11 - 1 + 3  # row 38
        self.assertEqual(ws.cell(sp_total_row, 1).value, "Total 202601")
        self.assertEqual(ws.cell(sp_total_row, 3).value, f'=countif(C{sp_data_start}:C{sp_data_start+10},"=Y")')
        self.assertEqual(ws.cell(sp_total_row, 8).value, f'=countif(H{sp_data_start}:H{sp_data_start+10},"=Y")')

        # 4. Summary & Totals
        r_sum_hdr = sp_total_row + 2
        self.assertEqual(ws.cell(r_sum_hdr, 3).value, "Fall")
        self.assertEqual(ws.cell(r_sum_hdr, 8).value, "Spring")

        r_sem_tot = r_sum_hdr + 1
        self.assertEqual(ws.cell(r_sem_tot, 1).value, "Semester Totals")
        self.assertEqual(ws.cell(r_sem_tot, 3).value, f"=C{fall_total_row}+H{fall_total_row}")
        self.assertEqual(ws.cell(r_sem_tot, 4).value, f"=D{fall_total_row}+I{fall_total_row}")
        self.assertEqual(ws.cell(r_sem_tot, 8).value, f"=C{sp_total_row}+H{sp_total_row}")
        self.assertEqual(ws.cell(r_sem_tot, 9).value, f"=D{sp_total_row}+I{sp_total_row}")

        r_gr_tot = r_sem_tot + 2
        self.assertEqual(ws.cell(r_gr_tot, 1).value, "Grand Totals")
        self.assertEqual(ws.cell(r_gr_tot, 3).value, f"=C{r_sem_tot}+H{r_sem_tot}")
        self.assertEqual(ws.cell(r_gr_tot, 4).value, f"=D{r_sem_tot}+I{r_sem_tot}")

        r_pct = r_gr_tot + 1
        self.assertEqual(ws.cell(r_pct, 1).value, "Percentage")
        self.assertEqual(ws.cell(r_pct, 3).value, f"=IF(D{r_gr_tot}>0,ROUND(C{r_gr_tot}/D{r_gr_tot}*100,0),0)")

        r_met = r_pct + 2
        self.assertEqual(ws.cell(r_met, 1).value, "Met PI?")
        self.assertEqual(ws.cell(r_met, 3).value, f'=IF(C{r_pct}>=70, "Y", "N")')

        # FERPA check
        all_values = [str(c.value) for row in ws.rows for c in row if c.value is not None]
        for name in ["Alice", "Adams", "Bob", "Baker", "Charlie", "Clark", "David", "Davis", "Eve", "Evans", "Frank", "Franklin"]:
            self.assertNotIn(name, all_values)
        for uwfid in ["100000001", "100000002", "100000003", "100000004", "100000005", "100000006"]:
            self.assertNotIn(uwfid, all_values)

        wb.close()


if __name__ == "__main__":
    unittest.main()

