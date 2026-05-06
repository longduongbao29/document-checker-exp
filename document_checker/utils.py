from __future__ import annotations

import re


def normalize_text(text: str) -> str:
    if not text:
        return ""
    cleaned = text.replace("\u00a0", " ")
    cleaned = re.sub(r"-\s+", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip().casefold()
    cleaned = re.sub(r"[^\w\s]+", "", cleaned, flags=re.UNICODE)
    return cleaned


def text_head(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[:max_len]


def text_tail(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[-max_len:]


def truncate_text(text: str, max_len: int = 120) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."
