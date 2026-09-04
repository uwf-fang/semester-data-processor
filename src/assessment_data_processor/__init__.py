"""
assessment_data_processor - Course Assessment Data Processing Package.

Automated tools for merging section rosters, consolidating Canvas gradebooks,
and generating ABET / SACSCOC assessment reports based on configurable outcome plans.
"""

from .rosters import process_rosters
from .gradebook import process_gradebook
from .assessments import extract_assessments

__version__ = "0.1.0"
__all__ = ["process_rosters", "process_gradebook", "extract_assessments", "__version__"]
