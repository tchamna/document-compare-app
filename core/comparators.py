"""Comparison algorithms: order-independent (PPTX) and sequential (DOCX/TXT)."""

from __future__ import annotations

from difflib import SequenceMatcher
from itertools import zip_longest
from typing import Dict, List

from core.helpers import LineDiff, normalize_whitespace


def compute_diffs(
    original_lines: Dict[int, List[str]],
    corrige_lines: Dict[int, List[str]],
) -> List[LineDiff]:
    """Order-independent per-slide comparison (best for PPTX)."""
    diffs: List[LineDiff] = []
    all_slides = set(original_lines.keys()) | set(corrige_lines.keys())

    for slide_no in sorted(all_slides):
        o = original_lines.get(slide_no, [])
        c = corrige_lines.get(slide_no, [])

        orig_norms = [normalize_whitespace(x) for x in o]
        corr_norms = [normalize_whitespace(x) for x in c]

        # Build map of normalised corrige lines -> available indices
        corr_map: Dict[str, List[int]] = {}
        for idx, val in enumerate(corr_norms):
            corr_map.setdefault(val, []).append(idx)

        matched_corr: set = set()
        matched_orig = [False] * len(orig_norms)

        # First pass: match identical content regardless of position
        for i, val in enumerate(orig_norms):
            if val in corr_map:
                lst = corr_map[val]
                while lst and lst[0] in matched_corr:
                    lst.pop(0)
                if lst:
                    matched_corr.add(lst.pop(0))
                    matched_orig[i] = True

        # Unmatched lines are real diffs
        unmatched_o = [o[i] for i, mflag in enumerate(matched_orig) if not mflag]
        unmatched_c = [c[i] for i in range(len(c)) if i not in matched_corr]

        for a, b in zip_longest(unmatched_o, unmatched_c, fillvalue=""):
            if normalize_whitespace(a) != normalize_whitespace(b):
                diffs.append(LineDiff(slide_no=slide_no, original=a, corrige=b))

    return [d for d in diffs if d.original.strip() or d.corrige.strip()]


def compute_diffs_sequential(
    original: List[str],
    corrige: List[str],
) -> List[LineDiff]:
    """Sequential comparison using SequenceMatcher (best for DOCX/TXT)."""
    orig_norms = [normalize_whitespace(x) for x in original]
    corr_norms = [normalize_whitespace(x) for x in corrige]

    sm = SequenceMatcher(None, orig_norms, corr_norms)
    diffs: List[LineDiff] = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        elif tag == "replace":
            for a, b in zip_longest(
                original[i1:i2], corrige[j1:j2], fillvalue=""
            ):
                diffs.append(LineDiff(slide_no=0, original=a, corrige=b))
        elif tag == "delete":
            for a in original[i1:i2]:
                diffs.append(LineDiff(slide_no=0, original=a, corrige=""))
        elif tag == "insert":
            for b in corrige[j1:j2]:
                diffs.append(LineDiff(slide_no=0, original="", corrige=b))

    return diffs
