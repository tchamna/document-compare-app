"""Shared helpers: normalisation, text splitting, data classes."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import List, Tuple


# ------------------------------------------------------------------ #
#  Normalisation
# ------------------------------------------------------------------ #
def normalize_whitespace(s: str) -> str:
    """NFC-normalise Unicode, collapse whitespace, strip."""
    s = unicodedata.normalize("NFC", s)
    s = s.replace("\u00A0", " ")          # non-breaking space
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s+\n", "\n", s)
    s = re.sub(r"\n\s+", "\n", s)
    return s.strip()


# ------------------------------------------------------------------ #
#  Text splitting
# ------------------------------------------------------------------ #
def split_into_lines(text: str) -> List[str]:
    """Split text on line breaks + sentence-end punctuation."""
    text = normalize_whitespace(text)
    chunks: List[str] = []
    for block in text.split("\n"):
        block = block.strip()
        if not block:
            continue
        parts = re.split(r"(?<=[\.\?\!])\s+", block)
        for p in parts:
            p = p.strip()
            if p:
                chunks.append(p)
    return chunks


def is_digits_only(s: str) -> bool:
    """Return True if the string is only digits/punctuation (page numbers, counters)."""
    s = (s or "").strip()
    if not s:
        return False
    cleaned = re.sub(r"[\s\.,:;()\[\]\-]+", "", s)
    return cleaned.isdigit()


# ------------------------------------------------------------------ #
#  Word-level diff pairs (for bold highlighting)
# ------------------------------------------------------------------ #
def word_diff_pairs(text1: str, text2: str) -> List[Tuple[str, bool]]:
    """
    Compare two sentences word-by-word.
    Returns [(word, is_different), ...] for *text1*.
    """
    words1 = text1.split()
    words2 = text2.split()

    result: List[Tuple[str, bool]] = []
    for w1, w2 in zip(words1, words2):
        result.append((w1, w1 != w2))

    # Extra words in either list are always "different"
    if len(words1) > len(words2):
        for w in words1[len(words2):]:
            result.append((w, True))
    elif len(words2) > len(words1):
        for w in words2[len(words1):]:
            result.append((w, True))

    return result


# ------------------------------------------------------------------ #
#  Data classes
# ------------------------------------------------------------------ #
@dataclass
class LineDiff:
    slide_no: int
    original: str
    corrige: str
