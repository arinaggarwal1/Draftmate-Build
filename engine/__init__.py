# engine/__init__.py
"""
DraftMate Engine - Pure Python email generation engine.

This module provides the core functionality for:
- Loading data from CSV files or Google Sheets
- Building preview rows with template assignments
- Generating Outlook email drafts

Public API:
-----------
Data Loading:
    load_csv(path: str) -> Tuple[List[Dict], List[str]]
    load_google_sheet(url: str) -> Tuple[List[Dict], List[str]]

Preview:
    build_preview_rows(rows, headers_lower, templates, ...) -> List[Dict]

Generation:
    generate_emails(rows, headers_lower, templates, ...) -> int

The engine has no UI dependencies and can be used standalone.
"""

from engine.data_sources import load_csv, load_google_sheet
from engine.preview import build_preview_rows, choose_template_for_row
from engine.generator import generate_emails
from engine.resolver import PlaceholderResolver

__all__ = [
    # Data loading
    "load_csv",
    "load_google_sheet",
    # Preview
    "build_preview_rows",
    "choose_template_for_row",
    # Generation
    "generate_emails",
    # Utilities
    "PlaceholderResolver",
]
