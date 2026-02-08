"""Generate a Word (.docx) diff report with bold-highlighted differences."""

from __future__ import annotations

from pathlib import Path
from typing import List

from docx import Document

from core.helpers import LineDiff, word_diff_pairs


def write_word_report(
    diffs: List[LineDiff],
    out_path: Path,
    original_name: str,
    corrige_name: str,
) -> None:
    """Write a Word document listing all differences.

    Differing words are **bolded** for easy visual scanning.
    """
    suffix = Path(original_name).suffix.lower()
    section_label = "Slide" if suffix == ".pptx" else "Paragraphe"

    doc = Document()
    doc.add_heading(
        "Differences between the original and corrected files", level=1
    )
    doc.add_paragraph("Files compared:")
    doc.add_paragraph(f"- Original file: {original_name}")
    doc.add_paragraph(f"- Corrected file: {corrige_name}")

    if not diffs:
        doc.add_paragraph("\nNo text differences detected.")
        doc.save(str(out_path))
        return

    current_slide = None
    for d in diffs:
        if d.slide_no != current_slide:
            current_slide = d.slide_no
            if section_label == "Slide":
                doc.add_heading(f"{section_label} {current_slide}", level=2)

        # --- Original ---
        doc.add_paragraph("Original:", style=None)
        p_orig = doc.add_paragraph()
        if d.original:
            for word, is_diff in word_diff_pairs(d.original, d.corrige):
                run = p_orig.add_run(word)
                if is_diff:
                    run.bold = True
                p_orig.add_run(" ")
        else:
            p_orig.add_run("(empty)")

        # --- Corrected ---
        doc.add_paragraph("Corrected:", style=None)
        p_corr = doc.add_paragraph()
        if d.corrige:
            for word, is_diff in word_diff_pairs(d.corrige, d.original):
                run = p_corr.add_run(word)
                if is_diff:
                    run.bold = True
                p_corr.add_run(" ")
        else:
            p_corr.add_run("(empty)")

        doc.add_paragraph("")  # spacer

    doc.save(str(out_path))
