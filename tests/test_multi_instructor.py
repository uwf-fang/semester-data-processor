"""Synthetic regression fixtures for instructor consolidation and per-item scope."""
import csv
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from semester_data_processor import process_gradebook, extract_assessments
from semester_data_processor.assessments import assessment_majors, load_or_create_plan


class MultiInstructorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.course = Path(self.tmp.name)
        self.raw = self.course / 'raw'
        self.raw.mkdir()
        self.write(self.course / 'combined_roster.csv',
                   ['UWFID', 'Full Name', 'Major', 'Status', 'Section Number'], [
                       ['1', 'Alpha, Ada', 'Computer Science', 'Active', '10001'],
                       ['2', 'Beta, Ben', 'Cybersecurity', 'Active', '10002'],
                       ['3', 'Gamma, Gia', 'Computer Science', 'Active', '10002'],
                       ['4', 'Absent, Alex', 'Computer Science', 'Active', '10001'],
                       ['5', 'Withdrawn, Will', 'Cybersecurity', 'Withdrawn', '10002'],
                   ])
        self.export('A_Grades.csv', ['Project 1 (111)', 'Project 2 (112)'], [50, 20], [
            ['Alpha, Ada', '1', '10001', 40, 10],
        ])
        self.export('B_grade.csv', ['Project 1 (211)', 'Project 2 (212)', 'Final (213)'], [100, 40, 10], [
            ['Beta, Ben', '2', '10002', 90, 32, 9],
            ['Gamma, Gia', '3', '10002', 60, 36, 8],
            ['Withdrawn, Will', '5', '10002', 100, 40, 10],
        ], manual=True)

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, path, headers, rows):
        with path.open('w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)

    def export(self, name, assignments, maxima, rows, manual=False):
        metadata = [['', '', ''] + ['Manual Posting'] * len(assignments)] if manual else []
        self.write(self.raw / name, ['Student', 'SIS User ID', 'Section'] + assignments,
                   metadata + [['Points Possible', '', ''] + maxima] + rows)

    def read(self, name):
        with (self.course / name).open() as f:
            return list(csv.DictReader(f))

    def plan(self, items, **kwargs):
        (self.course / 'assessment_plan.json').write_text(json.dumps({
            'course_name': 'Synthetic course', 'assessments': items, **kwargs
        }))

    def test_union_deduplication_and_roster_only_students(self):
        shutil.copy(self.raw / 'A_Grades.csv', self.raw / 'C_GRADE.csv')
        process_gradebook(str(self.course))
        rows = self.read('merged_gradebook.csv')
        self.assertEqual(len(rows), 5)
        by_id = {r['SIS User ID']: r for r in rows}
        self.assertEqual(by_id['2']['Final'], '9')
        self.assertEqual(by_id['1']['Final'], '')
        self.assertEqual(by_id['4']['Project 1'], '')

    def test_item_scopes_scales_and_excel_agree(self):
        self.plan([
            {'assessment_name': 'Project 1', 'code': 'CS', 'target_majors': ['CS']},
            {'assessment_name': 'Project 2', 'code': 'CY', 'target_majors': ['cybersecurity']},
            {'assessment_name': 'Project 1', 'code': 'ALL', 'target_majors': []},
            {'assessment_name': 'Project 2', 'code': 'OLD'},
            {'assessment_name': 'Project 1', 'code': 'NONE', 'target_majors': ['EE']},
        ], target_majors=['Computer Science'])
        extract_assessments(str(self.course))
        summary = {r['Code']: r for r in self.read('assessment_summary.csv')
                   if r['Section ID'] == 'Total (All Sections)'}
        expected = {'CS': (1, 3), 'CY': (1, 1), 'ALL': (2, 4), 'OLD': (1, 3), 'NONE': (0, 0)}
        details = self.read('assessment_results.csv')
        for code, counts in expected.items():
            self.assertEqual((int(summary[code]['Passed Count']), int(summary[code]['Total Evaluated'])), counts)
            item_rows = [r for r in details if r['Code'] == code]
            self.assertEqual((sum(int(r['Pass Count']) for r in item_rows),
                              sum(int(r['Total Student Count']) for r in item_rows)), counts)
        ada = next(r for r in details if r['Code'] == 'CS' and r['First Name'] == 'Ada')
        ben = next(r for r in details if r['Code'] == 'CY' and r['First Name'] == 'Ben')
        self.assertEqual((ada['Points Possible'], ada['Score Percentage (%)']), ('50', '80.0%'))
        self.assertEqual((ben['Points Possible'], ben['Score Percentage (%)']), ('40', '80.0%'))
        self.assertEqual(summary['NONE']['Outcome Status'], 'NO DATA')
        self.assertFalse(any(r['First Name'] == 'Will' for r in details))
        markdown = (self.course / 'assessment_report.md').read_text()
        self.assertNotIn('CS Majors Only', markdown)
        self.assertIn('cybersecurity', markdown)
        import openpyxl
        wb = openpyxl.load_workbook(self.course / 'assessment_report.xlsx', data_only=True)
        sheet_rows = list(wb['Assessment Summary & Plan'].values)
        for code, counts in expected.items():
            row = next(r for r in sheet_rows if r[0] == code and isinstance(r[2], int))
            self.assertEqual(row[2:4], counts)
            self.assertEqual(row[7], summary[code]['Target Major'])
        result_rows = list(wb['Student Assessment Results'].values)
        self.assertIn('Target Major', result_rows[0])
        self.assertEqual(len(result_rows) - 1, len(details))
        wb.close()

    def test_conflicting_scores_leave_existing_output_untouched(self):
        process_gradebook(str(self.course))
        before = (self.course / 'merged_gradebook.csv').read_bytes()
        self.export('C_Grades.csv', ['Project 1 (333)'], [50], [['Alpha, Ada', '1', '10001', 41]])
        with self.assertRaisesRegex(ValueError, 'Conflicting gradebook values'):
            process_gradebook(str(self.course))
        self.assertEqual((self.course / 'merged_gradebook.csv').read_bytes(), before)

    def test_conflicting_denominator(self):
        self.export('C_Grades.csv', ['Project 1 (333)'], [100], [['Alpha, Ada', '1', '10001', 40]])
        with self.assertRaisesRegex(ValueError, 'Conflicting Points Possible'):
            process_gradebook(str(self.course))

    def test_complementary_duplicates(self):
        self.export('C_Grades.csv', ['Project 1 (333)', 'Extra (334)'], [50, 10], [['Alpha, Ada', '1', '10001', '', 9]])
        process_gradebook(str(self.course))
        ada = next(r for r in self.read('merged_gradebook.csv') if r['SIS User ID'] == '1')
        self.assertEqual((ada['Project 1'], ada['Extra']), ('40', '9'))

    def test_missing_metadata_rejected(self):
        (self.raw / 'C_Grades.csv').write_text('Student,SIS User ID,Project 1\nSomeone,9,50\n')
        with self.assertRaisesRegex(ValueError, 'Missing Points Possible'):
            process_gradebook(str(self.course))

    def test_duplicate_cleaned_headers_rejected(self):
        self.export('C_Grades.csv', ['Quiz (1)', 'Quiz (2)'], [10, 20], [])
        with self.assertRaisesRegex(ValueError, 'Duplicate assignment titles'):
            process_gradebook(str(self.course))

    def test_invalid_denominator_not_assumed_100(self):
        self.export('A_Grades.csv', ['Project 1 (111)'], [''], [['Alpha, Ada', '1', '10001', 40]])
        self.plan([{'assessment_name': 'Project 1'}])
        with self.assertRaisesRegex(ValueError, 'Missing or invalid Points Possible'):
            extract_assessments(str(self.course))

    def test_empty_run_replaces_old_csv_and_withdrawn_stays_excluded(self):
        self.plan([{'assessment_name': 'Project 1'}])
        extract_assessments(str(self.course))
        roster = self.read('combined_roster.csv')
        for row in roster:
            row['Status'] = 'Withdrawn'
        self.write(self.course / 'combined_roster.csv', roster[0].keys(), [list(r.values()) for r in roster])
        extract_assessments(str(self.course))
        self.assertEqual(self.read('assessment_results.csv'), [])
        self.assertEqual(self.read('assessment_summary.csv')[0]['Outcome Status'], 'NO DATA')
        self.plan([])
        extract_assessments(str(self.course))
        self.assertEqual(self.read('assessment_summary.csv'), [])

    def test_invalid_major_configuration(self):
        for invalid in ['CS', None, [''], [123]]:
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                assessment_majors({}, {'target_majors': invalid})
        self.assertEqual(assessment_majors({'target_majors': ['CS']}, {'target_majors': []}), [])

    def test_new_template_uses_item_majors(self):
        template = load_or_create_plan(str(self.course))
        self.assertNotIn('target_majors', template)
        self.assertEqual(template['assessments'][0]['target_majors'], ['Computer Science'])

    def test_second_export_added_after_merge_is_included(self):
        process_gradebook(str(self.course))
        self.export('C_Grades.csv', ['Project 1 (333)'], [10], [['Absent, Alex', '4', '10001', 8]])
        self.plan([{'assessment_name': 'Project 1'}])
        extract_assessments(str(self.course))
        alex = next(r for r in self.read('assessment_results.csv') if r['First Name'] == 'Alex')
        self.assertEqual(alex['Score Percentage (%)'], '80.0%')
