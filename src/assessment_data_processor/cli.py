"""
cli.py - Command Line Interface for assessment_data_processor package.
"""

import sys
import argparse
from .rosters import process_rosters
from .gradebook import process_gradebook
from .assessments import extract_assessments


def main():
    parser = argparse.ArgumentParser(
        prog="assessment-processor",
        description="Automated semester course assessment data processing tool."
    )
    subparsers = parser.add_subparsers(dest="command", help="Available workflow commands")

    # Command: process-rosters
    roster_parser = subparsers.add_parser("process-rosters", help="Merge section rosters in <course_dir>/raw/")
    roster_parser.add_argument("course_dir", help="Path to course semester directory (e.g. cop4020-26spring)")

    # Command: process-gradebook
    gb_parser = subparsers.add_parser("process-gradebook", help="Merge Canvas gradebook with combined roster")
    gb_parser.add_argument("course_dir", help="Path to course semester directory (e.g. cop4020-26spring)")

    # Command: extract-assessments
    ass_parser = subparsers.add_parser("extract-assessments", help="Extract ABET / SACSCOC assessment metrics & report")
    ass_parser.add_argument("course_dir", help="Path to course semester directory (e.g. cop4020-26spring)")

    # Command: run-all
    all_parser = subparsers.add_parser("run-all", help="Execute complete pipeline (rosters -> gradebook -> assessments)")
    all_parser.add_argument("course_dir", help="Path to course semester directory (e.g. cop4020-26spring)")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    course_dir = args.course_dir

    if args.command == "process-rosters":
        process_rosters(course_dir)
    elif args.command == "process-gradebook":
        process_gradebook(course_dir)
    elif args.command == "extract-assessments":
        extract_assessments(course_dir)
    elif args.command == "run-all":
        print(f"=== Running Full Pipeline for {course_dir} ===")
        process_rosters(course_dir)
        process_gradebook(course_dir)
        extract_assessments(course_dir)
        print(f"\n=== Pipeline Completed Successfully for {course_dir} ===")


if __name__ == "__main__":
    main()
