"""Field extraction for scanned checks.

Two changes of substance versus the original single-ROI approach:

1. The amount is read twice - from the courtesy box (numerals, upper right)
   and from the legal line (handwritten words) - and the two are reconciled.
   Every US check carries the amount in both places, so disagreement is a
   reliable "a human should look at this" signal. The original code read only
   the legal line, which is the harder of the two by a wide margin.

2. Region-of-interest boxes live in a JSON file rather than in the source, so
   they can be retuned against a real scanner without a code change.
"""

import json
import os
import re
from dataclasses import dataclass, asdict
from typing import Dict, Optional, Tuple

import numpy as np

DEFAULT_ROI_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "check_rois.json")

# Fractions of (width, height) as (x0, y0, x1, y1) on a standard US personal
# check. Starting points only - retune on real scans before trusting them.
DEFAULT_ROIS: Dict[str, Tuple[float, float, float, float]] = {
    "payer_name": (0.02, 0.02, 0.55, 0.22),
    "courtesy_amount": (0.70, 0.30, 0.99, 0.52),
    "legal_amount": (0.03, 0.48, 0.88, 0.66),
    "memo": (0.03, 0.78, 0.50, 0.93),
}

_UNITS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
    "seventeen": 17, "eighteen": 18, "nineteen": 19,
}
_TENS = {
    "twenty": 20, "thirty": 30, "forty": 40, "fourty": 40, "fifty": 50,
    "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
}
_SCALES = {"hundred": 100, "thousand": 1000, "million": 1000000}
_NUMBER_WORDS = set(_UNITS) | set(_TENS) | set(_SCALES)


def load_rois(path: Optional[str] = None) -> Dict[str, Tuple[float, float, float, float]]:
    """Load ROI fractions, falling back to the built-in defaults."""
    path = path or DEFAULT_ROI_PATH
    rois = dict(DEFAULT_ROIS)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as handle:
            loaded = json.load(handle)
        for key, box in loaded.items():
            if isinstance(box, (list, tuple)) and len(box) == 4:
                rois[key] = tuple(float(v) for v in box)
    return rois


def crop(image: np.ndarray, box: Tuple[float, float, float, float]) -> np.ndarray:
    """Crop a fractional (x0, y0, x1, y1) box out of a BGR image."""
    height, width = image.shape[:2]
    x0, y0, x1, y1 = box
    left = max(0, min(width, int(width * x0)))
    right = max(0, min(width, int(width * x1)))
    top = max(0, min(height, int(height * y0)))
    bottom = max(0, min(height, int(height * y1)))
    if right <= left or bottom <= top:
        raise ValueError(f"Empty crop for box {box} on a {width}x{height} image")
    return image[top:bottom, left:right]


def upscale(image: np.ndarray, min_height: int = 96) -> np.ndarray:
    """Enlarge small crops - recognition models degrade badly below ~64px tall."""
    import cv2

    height = image.shape[0]
    if height >= min_height or height == 0:
        return image
    factor = min_height / float(height)
    return cv2.resize(image, None, fx=factor, fy=factor, interpolation=cv2.INTER_CUBIC)


def _normalise_digits(text: str) -> str:
    """Repair the digit/letter confusions OCR makes inside a numeric field."""
    table = str.maketrans({"O": "0", "o": "0", "D": "0", "l": "1", "I": "1",
                           "|": "1", "S": "5", "s": "5", "B": "8", "Z": "2"})
    return text.translate(table)


def parse_courtesy_amount(text: str) -> Optional[float]:
    """Parse the numeric amount box, e.g. '$1,250.00', '125 00', '125-xx'."""
    if not text:
        return None
    cleaned = _normalise_digits(text.strip())
    cleaned = cleaned.replace("$", " ")
    # Dollars, then an optional cents group written as .00 / -00 / /100 / xx.
    match = re.search(
        r"(\d[\d,]*)\s*(?:[.\-–—/ ]\s*(\d{1,2}|xx|no)\s*(?:/\s*100)?)?",
        cleaned,
        re.IGNORECASE,
    )
    if not match:
        return None
    try:
        dollars = int(match.group(1).replace(",", ""))
    except ValueError:
        return None
    cents = 0
    raw_cents = match.group(2)
    if raw_cents and raw_cents.isdigit():
        cents = int(raw_cents.ljust(2, "0")[:2])
    return dollars + cents / 100.0


def _words_to_number(tokens) -> Optional[int]:
    """Classic accumulator over number words.

    Replaces word2number, which raises on any unexpected token and silently
    mis-parses repeated scales. Unknown tokens are dropped before we get here.
    """
    total = 0
    current = 0
    seen = False
    for token in tokens:
        if token in _UNITS:
            current += _UNITS[token]
            seen = True
        elif token in _TENS:
            current += _TENS[token]
            seen = True
        elif token == "hundred":
            current = (current or 1) * 100
            seen = True
        elif token in _SCALES:
            total += (current or 1) * _SCALES[token]
            current = 0
            seen = True
    if not seen:
        return None
    return total + current


def parse_legal_amount(text: str) -> Optional[float]:
    """Parse the written-words amount line.

    Handles the conventional '<words> and NN/100 Dollars' form, including the
    'no/100' and 'xx/100' spellings for zero cents.
    """
    if not text:
        return None
    lowered = text.lower().replace("-", " ").replace("*", " ")
    lowered = re.sub(r"\bdollars?\b", " ", lowered)
    lowered = re.sub(r"\bonly\b", " ", lowered)

    cents = 0
    fraction = re.search(r"\b(\d{1,2}|no|xx)\s*/\s*100\b", lowered)
    if fraction:
        raw = fraction.group(1)
        if raw.isdigit():
            cents = int(raw.ljust(2, "0")[:2])
        lowered = lowered[: fraction.start()] + " " + lowered[fraction.end():]

    # Drop the 'and' that separates dollars from cents, then keep only tokens
    # that are actually number words.
    tokens = [t for t in re.split(r"[^a-z0-9]+", lowered) if t]
    tokens = [t for t in tokens if t in _NUMBER_WORDS]
    dollars = _words_to_number(tokens)
    if dollars is None:
        return None
    return dollars + cents / 100.0


@dataclass
class AmountReading:
    """Both amount reads plus the reconciliation verdict."""

    value: Optional[float]
    courtesy: Optional[float]
    legal: Optional[float]
    status: str
    needs_review: bool
    courtesy_text: str = ""
    legal_text: str = ""

    def as_dict(self) -> dict:
        return asdict(self)


def reconcile_amount(courtesy: Optional[float], legal: Optional[float],
                     prefer: str = "courtesy", tolerance: float = 0.005) -> Tuple[Optional[float], str, bool]:
    """Combine the two amount reads into one value plus a review verdict.

    Returns (value, status, needs_review). 'courtesy' is preferred by default
    because printed/handwritten numerals OCR far more reliably than cursive
    words; set prefer='legal' to follow banking convention instead.
    """
    if courtesy is not None and legal is not None:
        if abs(courtesy - legal) <= tolerance:
            return courtesy, "agree", False
        chosen = courtesy if prefer == "courtesy" else legal
        return chosen, "mismatch", True
    if courtesy is not None:
        return courtesy, "courtesy_only", True
    if legal is not None:
        return legal, "legal_only", True
    return None, "unreadable", True
