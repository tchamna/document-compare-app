"""Core comparison engine for document-compare-app."""

from core.extractors import (
    extract_slide_lines,
    extract_docx_lines,
    extract_txt_lines,
    extract_text_lines,
)
from core.comparators import compute_diffs, compute_diffs_sequential
from core.report import write_word_report
from core.helpers import LineDiff

__all__ = [
    "extract_slide_lines",
    "extract_docx_lines",
    "extract_txt_lines",
    "extract_text_lines",
    "compute_diffs",
    "compute_diffs_sequential",
    "write_word_report",
    "LineDiff",
]
