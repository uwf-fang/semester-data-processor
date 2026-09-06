"""
semester_data_processor - Semester Course Data Processing Package.

Automated tools for merging section rosters, consolidating Canvas gradebooks,
generating dataset documentation with processing log records, and extracting ABET / SACSCOC assessment reports.
"""

from .rosters import process_rosters
from .gradebook import process_gradebook
from .assessments import extract_assessments
from .readme import generate_output_readme

__version__ = "0.1.0"
__all__ = ["process_rosters", "process_gradebook", "extract_assessments", "generate_output_readme", "__version__"]
