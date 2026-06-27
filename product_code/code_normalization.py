"""Product-code normalization helpers shared by backend workflows."""

import re


def normalize_code_decimal_separators(text):
    return re.sub(r"(?<=\d),(?=\d)", ".", str(text or ""))


def sanitize_product_code(value, accent_normalizer=None, trim=None):
    """Remove FAST-forbidden product-code characters after a code is generated.

    `accent_normalizer` and `trim` are injected by app.py so the extracted helper
    keeps using the existing Vietmax accent and max-length rules.
    """
    text = str(value or "")
    if accent_normalizer:
        text = accent_normalizer(text)
    text = normalize_code_decimal_separators(text)
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^A-Za-z0-9.]+", "", text)
    return trim(text) if trim else text
