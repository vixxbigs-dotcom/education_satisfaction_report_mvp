from __future__ import annotations

import re


_LEADING_QUESTION_NO_RE = re.compile(r"^\s*\d+\s*[.)]\s*")


def strip_leading_question_numbers(text: str) -> str:
    """Remove one or more leading question-number prefixes such as '1.' or '2)'.

    Some survey files restart numbering for each module. The report assigns a new
    continuous number, so embedded source numbers must be removed first to avoid
    output such as '2. 1. [과정명] ...'.
    """
    cleaned = str(text or "").strip()
    while True:
        updated = _LEADING_QUESTION_NO_RE.sub("", cleaned, count=1).strip()
        if updated == cleaned:
            return cleaned
        cleaned = updated
