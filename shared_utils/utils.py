import re


def parse_amount(text: str) -> float:
    val = re.sub(r"[^\d\.,]", "", text).replace(",", ".")
    try:
        return float(val)
    except Exception:
        return 0.0


def extract_deal_id(text: str) -> str:
    m = re.search(r"(?:№)?(\d{4})", text)
    return m.group(1) if m else ""
