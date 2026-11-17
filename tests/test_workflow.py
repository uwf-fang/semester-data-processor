#!/usr/bin/env python3
"""
tests/test_workflow.py - Automated Unit Tests for Assessment Data Processing Package.

Tests:
1. Dynamic Gradebook Row Parsing (with and without Manual Posting rows).
2. Section CRN Extraction from raw roster filenames.
3. Roster vs Gradebook Audit Reconciliation (ensuring no student is silently dropped).
4. Multiple Learning Outcomes / Performance Indicators per Assessment Plan.
5. Multi-Stage Fuzzy Assignment Title Matching (handling decorative words, Canvas ID suffixes, zero-padding).

Usage:
    python3 -m unittest discover tests
"""

import sys
import os
import shutil
import tempfile
import json
import unittest

# Add src directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from semester_data_processor import process_rosters, process_gradebook, extract_assessments, generate_output_readme
from semester_data_processor.assessments import find_matching_columns


class TestSemesterWorkflow(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.raw_dir = os.path.join(self.test_dir, "raw")
        os.makedirs(self.raw_dir, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.test_dir)

    def test_fuzzy_assignment_matching(self):
        """Test multi-stage fuzzy assignment matching for decorative words and formatting variations."""
        headers = [
            "Student", "ID", "SIS User ID",
            "Project 1 (1270399)",
            "Project 2 (1282964)",
            "Midterm 1 - Requires Respondus LockDown Browser + Webcam (1305419)",
            "Review-02-Lexical  (1300999)",
            "Final Exam - Requires Respondus LockDown Browser + Webcam (1328299)"
        ]

        # 1. Zero-padded number query
        m1 = find_matching_columns(headers, "project 01")
        self.assertEqual(m1, ["Project 1 (1270399)"])

        # 2. Punctuation query
        m2 = find_matching_columns(headers, "Project-1")
        self.assertEqual(m2, ["Project 1 (1270399)"])

        # 3. Query matching long decorative title
        m3 = find_matching_columns(headers, "Midterm 1")
        self.assertEqual(m3, ["Midterm 1 - Requires Respondus LockDown Browser + Webcam (1305419)"])

        # 4. Query matching fuzzy title
        m4 = find_matching_columns(headers, "Review 2")
        self.assertEqual(m4, ["Review-02-Lexical  (1300999)"])

    def test_roster_section_extraction(self):
        """Test section CRN extraction from roster filenames including suffixes like (1).csv."""
        roster_csv_1 = os.path.join(self.raw_dir, "CourseRoster_202601_18451 (1).csv")
        with open(roster_csv_1, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000003, Avery, Example, fixture1@example.com, Undergraduate, Computer Science\n")

        roster_csv_2 = os.path.join(self.raw_dir, "CourseRoster_202601_18626 (1).csv")
        with open(roster_csv_2, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000001, Blair, Fiction, fixture2@example.com, Undergraduate, Computer Science\n")

        output_path = process_rosters(self.test_dir)
        self.assertTrue(os.path.exists(output_path))

        with open(output_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("18451", content)
        self.assertIn("18626", content)

    def test_gradebook_without_manual_posting_row(self):
        """Test parsing Canvas export WITHOUT Manual Posting row."""
        roster_csv = os.path.join(self.raw_dir, "CourseRoster_202601_18451.csv")
        with open(roster_csv, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000005, Casey, Sample, fixture3@example.com, Undergraduate, Computer Science\n")
            f.write("100000002, Drew, Demo, fixture4@example.com, Undergraduate, Computer Science\n")

        gb_csv = os.path.join(self.raw_dir, "2026-09-02_Grades-COPXXXX.csv")
        with open(gb_csv, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1270399)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Sample, Casey",90002,100000005,fixture3,18451_COPXXXX,100.00\n')
            f.write('"Demo, Drew",90003,100000002,fixture4,18451_COPXXXX,90.00\n')

        process_rosters(self.test_dir)
        output_gb = process_gradebook(self.test_dir)

        self.assertTrue(os.path.exists(output_gb))
        with open(output_gb, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("100000005", content)
        self.assertIn("Sample, Casey", content)
        self.assertIn("100.00", content)

    def test_gradebook_with_manual_posting_row(self):
        """Test parsing Canvas export WITH Manual Posting row (4 header/metadata rows)."""
        roster_csv = os.path.join(self.raw_dir, "CourseRoster_202508_87575.csv")
        with open(roster_csv, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000004, Emery, Placeholder, fixture5@example.com, Undergraduate, Computer Science\n")

        gb_csv = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4534.csv")
        with open(gb_csv, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1200002)\n")
            f.write(",,,,,Manual Posting\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Placeholder, Emery",90001,100000004,fixture5,87575_COP4534,100.00\n')

        process_rosters(self.test_dir)
        output_gb = process_gradebook(self.test_dir)

        self.assertTrue(os.path.exists(output_gb))
        with open(output_gb, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("100000004", content)
        self.assertIn("Placeholder, Emery", content)

    def test_roster_reconciliation(self):
        """Test that active students in roster are never dropped even if missing from Canvas gradebook."""
        roster_csv = os.path.join(self.raw_dir, "CourseRoster_87575.csv")
        with open(roster_csv, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000004, Emery, Placeholder, fixture5@example.com, Undergraduate, Computer Science\n")
            f.write("100000006, Frankie, Testcase, fixture6@example.com, Undergraduate, Computer Science\n")

        gb_csv = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4534.csv")
        with open(gb_csv, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1200002)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Placeholder, Emery",90001,100000004,fixture5,87575_COP4534,100.00\n')

        process_rosters(self.test_dir)
        output_gb = process_gradebook(self.test_dir)

        with open(output_gb, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("100000006", content)
        self.assertIn("Testcase, Frankie", content)

    def test_multiple_learning_outcomes_in_plan(self):
        """Test processing plan containing MULTIPLE learning outcomes / performance indicators."""
        roster_csv = os.path.join(self.raw_dir, "CourseRoster_87575.csv")
        with open(roster_csv, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000004, Emery, Placeholder, fixture5@example.com, Undergraduate, Computer Science\n")

        gb_csv = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4534.csv")
        with open(gb_csv, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1200002),Project 2 (1200003)\n")
            f.write("    Points Possible,,,,,100.00,100.00\n")
            f.write('"Placeholder, Emery",90001,100000004,fixture5,87575_COP4534,100.00,85.00\n')

        plan_data = {
            "course_name": "COP 4534 - Data Structures II",
            "semester": "Fall 2025",
            "target_majors": ["Computer Science"],
            "assessments": [
                {
                    "accreditation_type": "ABET",
                    "assessment_name": "ABET LO.PI 1.1 - Problem Analysis",
                    "assignment_column": "Project 1",
                    "passing_operator": ">=",
                    "passing_threshold_percent": 70.0,
                    "target_pass_rate_percent": 75.0
                },
                {
                    "accreditation_type": "ABET",
                    "assessment_name": "ABET LO.PI 1.2 - Design & Implementation",
                    "assignment_column": "Project 2",
                    "passing_operator": ">=",
                    "passing_threshold_percent": 70.0,
                    "target_pass_rate_percent": 75.0
                }
            ]
        }
        with open(os.path.join(self.test_dir, "assessment_plan.json"), "w", encoding="utf-8") as f:
            json.dump(plan_data, f)

        process_rosters(self.test_dir)
        process_gradebook(self.test_dir)
        report_path = extract_assessments(self.test_dir)

        self.assertTrue(os.path.exists(report_path))
        with open(report_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("ABET LO.PI 1.1 - Problem Analysis", content)
        self.assertIn("ABET LO.PI 1.2 - Design & Implementation", content)

        results_csv = os.path.join(self.test_dir, "assessment_results.csv")
        self.assertTrue(os.path.exists(results_csv))
        with open(results_csv, "r", encoding="utf-8") as f:
            csv_content = f.read()

        self.assertIn("Points Possible", csv_content)
        self.assertIn("Score Percentage (%)", csv_content)
        self.assertIn("Passing Criterion", csv_content)

        summary_csv = os.path.join(self.test_dir, "assessment_summary.csv")
        self.assertTrue(os.path.exists(summary_csv))
        with open(summary_csv, "r", encoding="utf-8") as f:
            summary_content = f.read()

        self.assertIn("Passed Count", summary_content)
        self.assertIn("Total Evaluated", summary_content)
        self.assertIn("Pass Rate (%)", summary_content)
        self.assertNotIn("Fail Count", summary_content)
        self.assertNotIn("Failed Count", summary_content)

    def test_output_readme_and_processing_log(self):
        """Test dataset output README generation and ongoing processing log tracking."""
        roster_csv = os.path.join(self.raw_dir, "CourseRoster_87575.csv")
        with open(roster_csv, "w", encoding="utf-8") as f:
            f.write("UWFID, First Name, Last Name, Email, Level, Program\n")
            f.write("100000004, Emery, Placeholder, fixture5@example.com, Undergraduate, Computer Science\n")

        # 1. Process rosters and check generated README
        process_rosters(self.test_dir)
        readme_path = os.path.join(self.test_dir, "README.md")
        self.assertTrue(os.path.exists(readme_path))

        with open(readme_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("Dataset & Semester Processing Documentation", content)
        self.assertIn("Processing Log & Updates", content)
        self.assertIn("Process Rosters", content)
        self.assertIn("CourseRoster_87575.csv", content)

        # 2. Trigger additional stage and verify log section persistence
        gb_csv = os.path.join(self.raw_dir, "2026-09-02_Grades-COP4534.csv")
        with open(gb_csv, "w", encoding="utf-8") as f:
            f.write("Student,ID,SIS User ID,SIS Login ID,Section,Project 1 (1200002)\n")
            f.write("    Points Possible,,,,,100.00\n")
            f.write('"Placeholder, Emery",90001,100000004,fixture5,87575_COP4534,100.00\n')

        process_gradebook(self.test_dir)

        with open(readme_path, "r", encoding="utf-8") as f:
            updated_content = f.read()

        self.assertIn("Process Gradebook", updated_content)
        self.assertIn("Process Rosters", updated_content)


if __name__ == "__main__":
    unittest.main()
