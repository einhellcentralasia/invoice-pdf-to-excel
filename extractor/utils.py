import re

def normalize_number(text):
    """Normalize numbers like '1 109,20' → '1109.20'"""
    if not text:
        return None
    text = str(text).replace(" ", "").replace("\u00a0", "")
    text = text.replace(",", ".")
    return text

def safe_float(x):
    try:
        return float(x)
    except Exception:
        return None
