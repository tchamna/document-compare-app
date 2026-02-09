"""
Document Compare – Streamlit Web App

Compare two documents (.pptx, .docx, .txt) side-by-side.
Differences are displayed in-browser and can be downloaded
as a Word report with changed words highlighted in bold.
"""

from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path

import streamlit as st

from core.extractors import extract_slide_lines, extract_text_lines
from core.comparators import compute_diffs, compute_diffs_sequential
from core.report import write_word_report, write_pdf_report
from core.helpers import LineDiff, word_diff_pairs

# ------------------------------------------------------------------ #
#  Page config
# ------------------------------------------------------------------ #
st.set_page_config(
    page_title="Document Compare",
    page_icon="📄",
    layout="wide",
)

SUPPORTED_EXTENSIONS = ["pptx", "docx", "txt"]


# ------------------------------------------------------------------ #
#  Helpers
# ------------------------------------------------------------------ #
def _save_uploaded(uploaded) -> Path:
    """Write an UploadedFile to a temp file and return its Path."""
    suffix = Path(uploaded.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.getvalue())
    tmp.close()
    return Path(tmp.name)


def _highlight_html(text: str, other: str) -> str:
    """Return HTML where differing words are wrapped in <mark>."""
    if not text:
        return '<span style="color:#888;">(empty)</span>'
    pairs = word_diff_pairs(text, other)
    parts = []
    for word, is_diff in pairs:
        if is_diff:
            parts.append(
                f'<mark style="background:#ffd54f;padding:0 2px;'
                f'border-radius:3px;font-weight:600;">{word}</mark>'
            )
        else:
            parts.append(word)
    return " ".join(parts)


# ------------------------------------------------------------------ #
#  UI
# ------------------------------------------------------------------ #
st.title("📄 Document Compare")
st.markdown(
    "Upload two files (`.pptx`, `.docx`, or `.txt`) to compare their text content. "
    "Changed words are **highlighted** for easy review."
)

col1, col2 = st.columns(2)

with col1:
    original_file = st.file_uploader(
        "Original file",
        type=SUPPORTED_EXTENSIONS,
        key="original",
    )

with col2:
    corrected_file = st.file_uploader(
        "Corrected file",
        type=SUPPORTED_EXTENSIONS,
        key="corrected",
    )

if original_file and corrected_file:
    # Validate same extension
    ext_orig = Path(original_file.name).suffix.lower()
    ext_corr = Path(corrected_file.name).suffix.lower()

    if ext_orig != ext_corr:
        st.error(
            f"Both files must be the same type. "
            f"Got **{ext_orig}** and **{ext_corr}**."
        )
        st.stop()

    # Save to temp
    with st.spinner("Reading files…"):
        path_orig = _save_uploaded(original_file)
        path_corr = _save_uploaded(corrected_file)

    # ---- Extract & Compare ---------------------------------------- #
    with st.spinner("Comparing…"):
        try:
            if ext_orig == ".pptx":
                orig_lines = extract_slide_lines(path_orig)
                corr_lines = extract_slide_lines(path_corr)
                diffs = compute_diffs(orig_lines, corr_lines)
            else:
                orig_flat = extract_text_lines(path_orig)
                corr_flat = extract_text_lines(path_corr)
                diffs = compute_diffs_sequential(orig_flat, corr_flat)
        except Exception as e:
            st.error(f"Error during comparison: {e}")
            st.stop()

    # ---- Results -------------------------------------------------- #
    st.divider()

    if not diffs:
        st.success("✅ No text differences found!")
    else:
        st.info(f"Found **{len(diffs)}** difference{'s' if len(diffs) != 1 else ''}.")

        # Build the Word report in memory for download
        report_docx_path = Path(tempfile.mktemp(suffix=".docx"))
        write_word_report(
            diffs=diffs,
            out_path=report_docx_path,
            original_name=original_file.name,
            corrige_name=corrected_file.name,
        )
        report_docx_buf = BytesIO(report_docx_path.read_bytes())
        report_docx_buf.seek(0)

        # Build the PDF report in memory for download
        report_pdf_path = Path(tempfile.mktemp(suffix=".pdf"))
        write_pdf_report(
            diffs=diffs,
            out_path=report_pdf_path,
            original_name=original_file.name,
            corrige_name=corrected_file.name,
        )
        report_pdf_buf = BytesIO(report_pdf_path.read_bytes())
        report_pdf_buf.seek(0)

        # Download buttons
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️  Download Word Report",
                data=report_docx_buf,
                file_name=f"Differences_{Path(original_file.name).stem}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
            )
        with col2:
            st.download_button(
                label="⬇️  Download PDF Report",
                data=report_pdf_buf,
                file_name=f"Differences_{Path(original_file.name).stem}.pdf",
                mime="application/pdf",
            )

        # In-browser preview
        st.subheader("Differences")

        for i, d in enumerate(diffs):
            with st.container():
                # Section header for PPTX slides
                if ext_orig == ".pptx" and d.slide_no:
                    if i == 0 or diffs[i - 1].slide_no != d.slide_no:
                        st.markdown(f"### Slide {d.slide_no}")

                left, right = st.columns(2)
                with left:
                    st.markdown("**Original**", unsafe_allow_html=True)
                    st.markdown(
                        _highlight_html(d.original, d.corrige),
                        unsafe_allow_html=True,
                    )
                with right:
                    st.markdown("**Corrected**", unsafe_allow_html=True)
                    st.markdown(
                        _highlight_html(d.corrige, d.original),
                        unsafe_allow_html=True,
                    )
                st.divider()

    # Clean up temp files
    path_orig.unlink(missing_ok=True)
    path_corr.unlink(missing_ok=True)
    report_docx_path.unlink(missing_ok=True)
    report_pdf_path.unlink(missing_ok=True)

else:
    st.caption("👆 Upload both files to get started.")
