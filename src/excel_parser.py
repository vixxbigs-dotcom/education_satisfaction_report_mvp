from __future__ import annotations

import io
import math
import re
from collections import Counter
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook

from .models import ObjectiveQuestion, ReportData, SubjectiveQuestion
from .text_utils import strip_leading_question_numbers


DEFAULT_SCALE = ["매우 그렇다", "그렇다", "보통이다", "그렇지 않다", "매우 그렇지 않다"]

TEXT_SCORE_MAP = {
    "매우그렇다": 5,
    "그렇다": 4,
    "보통이다": 3,
    "보통": 3,
    "그렇지않다": 2,
    "매우그렇지않다": 1,
    "매우만족": 5,
    "만족": 4,
    "불만족": 2,
    "매우불만족": 1,
    "매우우수": 5,
    "우수": 4,
    "미흡": 2,
    "매우미흡": 1,
    "매우도움": 5,
    "도움": 4,
    "도움안됨": 2,
    "전혀도움안됨": 1,
}

QUESTION_HINTS = (
    "만족", "추천", "적절", "난이도", "강의", "교육", "좋았", "아쉬", "개선",
    "운영", "도움", "평가", "의견", "작성", "전반적", "전문", "과정"
)

TITLE_SUFFIXES = (
    "교수", "강사", "코치", "대표", "노무사", "변호사", "박사", "컨설턴트",
    "연구위원", "위원", "전문가", "선생님"
)

NO_OPINION_CORES = {
    "", "없", "없음", "없습니다", "없어요", "없다", "없었음", "없었습니다",
    "해당없음", "해당사항없음", "특별히없음", "특별히없었음", "딱히없음",
    "무", "무응답", "na", "n/a", "-", ".", "x"
}


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _norm(value: Any) -> str:
    return re.sub(r"[\s\u00a0]+", "", str(value or "")).strip().lower()


def _is_question(text: Any) -> bool:
    t = _clean_text(text)
    if len(t) < 5:
        return False
    if "?" in t or "습니까" in t or "작성해" in t:
        return True
    return any(k in t for k in QUESTION_HINTS)


def _to_score(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if math.isnan(value) if isinstance(value, float) else False:
            return None
        iv = int(round(float(value)))
        if 1 <= iv <= 7:
            return iv
    t = _norm(value)
    if not t:
        return None
    m = re.fullmatch(r"([1-7])(?:점)?", t)
    if m:
        return int(m.group(1))
    return TEXT_SCORE_MAP.get(t)


def _looks_like_likert(values: Sequence[Any]) -> bool:
    nonempty = [v for v in values if _clean_text(v)]
    if not nonempty:
        return False
    mapped = [_to_score(v) for v in nonempty]
    hits = sum(v is not None for v in mapped)
    return hits >= max(2, math.ceil(len(nonempty) * 0.6))


def _question_number(text: str, fallback: int) -> int:
    m = re.match(r"\s*(\d+)\s*[.)]?", text)
    return int(m.group(1)) if m else fallback


def _extract_module(text: str) -> str:
    m = re.search(r"\[\s*([^\]]+?)\s*\]", text)
    return _clean_text(m.group(1)) if m else ""


def _extract_instructor(text: str) -> str:
    suffix = "|".join(map(re.escape, TITLE_SUFFIXES))
    matches = re.findall(rf"([가-힣A-Za-z][가-힣A-Za-z·.\- ]{{1,20}}?\s*(?:{suffix}))(?=의|는|은|께서|\s|$)", text)
    if not matches:
        return ""
    candidate = _clean_text(matches[-1])
    candidate = re.sub(r"^.*?\]\s*", "", candidate)
    candidate = re.sub(r"^\d+\s*[.)]?\s*", "", candidate)
    words = candidate.split()
    if len(words) > 3:
        candidate = " ".join(words[-2:])
    return candidate


def _classify_question(text: str, is_objective: bool) -> Tuple[str, str]:
    t = _clean_text(text)
    if not is_objective:
        if any(k in t for k in ("좋았", "장점", "인상", "유익", "만족스러운")):
            return "좋았던 점", "subjective_good"
        if any(k in t for k in ("아쉬", "개선", "보완", "불편", "단점", "건의", "바라는")):
            return "아쉬웠던 점", "subjective_bad"
        return "주관식", "subjective_other"

    if "추천" in t:
        return "추천도", "recommendation"
    if "일정" in t or "시간" in t:
        return "교육 일정", "schedule"
    if "난이도" in t:
        return "교육 난이도", "difficulty"
    if any(k in t for k in ("장소", "사전 안내", "과정 운영", "교육 운영", "운영")):
        return "교육 운영", "operation"
    if _extract_instructor(t):
        return "강사 만족도", "instructor"
    if "전반" in t and "만족" in t:
        return "전반적 만족도", "overall"
    if _extract_module(t) or "과정에 만족" in t:
        return "과정 만족도", "course"
    return "만족도", "other"


def _build_objective(question: str, values: Sequence[Any], fallback_no: int, qid: str) -> ObjectiveQuestion:
    source_question = _clean_text(question)
    cleaned_question = strip_leading_question_numbers(source_question)
    scores = [s for s in (_to_score(v) for v in values) if s is not None]
    max_point = max(scores) if scores else 5
    max_point = 5 if max_point <= 5 else 7
    labels = DEFAULT_SCALE if max_point == 5 else [f"{i}점" for i in range(max_point, 0, -1)]
    counts = [scores.count(score) for score in range(max_point, 0, -1)]
    avg = round(sum(scores) / len(scores), 2) if scores else 0.0
    section, category = _classify_question(cleaned_question, True)
    return ObjectiveQuestion(
        id=qid,
        number=_question_number(source_question, fallback_no),
        section_label=section,
        question=cleaned_question,
        counts=counts,
        scale_labels=list(labels),
        average=avg,
        valid_responses=len(scores),
        category=category,
        instructor_name=_extract_instructor(cleaned_question),
        course_module=_extract_module(cleaned_question),
    )


def _build_subjective(question: str, values: Sequence[Any], fallback_no: int, qid: str) -> SubjectiveQuestion:
    source_question = _clean_text(question)
    cleaned_question = strip_leading_question_numbers(source_question)
    answers = []
    for v in values:
        t = _clean_text(v)
        if t:
            answers.append(t)
    section, category = _classify_question(cleaned_question, False)
    return SubjectiveQuestion(
        id=qid,
        number=_question_number(source_question, fallback_no),
        section_label=section,
        question=cleaned_question,
        answers=answers,
        category=category,
    )


def _sheet_matrix(ws) -> List[List[Any]]:
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    matrix: List[List[Any]] = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True):
        matrix.append(list(row))
    return matrix


def _orientation_score(matrix: List[List[Any]]) -> Tuple[int, int]:
    if not matrix:
        return 0, 0
    first_row = matrix[0]
    first_col = [row[0] if row else None for row in matrix]
    row_score = sum(_is_question(v) for v in first_row)
    col_score = sum(_is_question(v) for v in first_col)
    return row_score, col_score


def _parse_column_oriented(matrix: List[List[Any]], sheet_name: str) -> Tuple[List[ObjectiveQuestion], List[SubjectiveQuestion], int]:
    headers = matrix[0] if matrix else []
    rows = matrix[1:]
    objective: List[ObjectiveQuestion] = []
    subjective: List[SubjectiveQuestion] = []
    qidx = 1
    response_count = sum(1 for row in rows if any(_clean_text(v) for v in row))
    for ci, header in enumerate(headers):
        question = _clean_text(header)
        if not question or "타임" in question.lower() or question.lower() in {"timestamp", "time"}:
            continue
        values = [row[ci] if ci < len(row) else None for row in rows]
        qid = f"{sheet_name}-c{ci+1}"
        if _looks_like_likert(values):
            objective.append(_build_objective(question, values, qidx, qid))
        elif _is_question(question):
            subjective.append(_build_subjective(question, values, qidx, qid))
        qidx += 1
    return objective, subjective, response_count


def _parse_row_oriented(matrix: List[List[Any]], sheet_name: str) -> Tuple[List[ObjectiveQuestion], List[SubjectiveQuestion], int]:
    objective: List[ObjectiveQuestion] = []
    subjective: List[SubjectiveQuestion] = []
    qidx = 1
    response_count = 0
    for ri, row in enumerate(matrix):
        if not row:
            continue
        question = _clean_text(row[0])
        if not question or "타임" in question.lower() or question.lower() in {"timestamp", "time"}:
            if question and "타임" in question:
                response_count = max(response_count, sum(1 for v in row[1:] if _clean_text(v)))
            continue
        if not _is_question(question):
            continue
        values = list(row[1:])
        response_count = max(response_count, sum(1 for v in values if _clean_text(v)))
        qid = f"{sheet_name}-r{ri+1}"
        if _looks_like_likert(values):
            objective.append(_build_objective(question, values, qidx, qid))
        else:
            subjective.append(_build_subjective(question, values, qidx, qid))
        qidx += 1
    return objective, subjective, response_count


def _sheet_candidate(matrix: List[List[Any]], sheet_name: str) -> Optional[Dict[str, Any]]:
    row_score, col_score = _orientation_score(matrix)
    if max(row_score, col_score) == 0:
        return None
    if col_score > row_score:
        objective, subjective, n = _parse_row_oriented(matrix, sheet_name)
        orientation = "문항이 행 / 응답자가 열"
    else:
        objective, subjective, n = _parse_column_oriented(matrix, sheet_name)
        orientation = "문항이 열 / 응답자가 행"
    quality = len(objective) * 5 + len(subjective) * 2 + min(n, 30)
    return {
        "sheet_name": sheet_name,
        "orientation": orientation,
        "objective": objective,
        "subjective": subjective,
        "response_count": n,
        "quality": quality,
    }


def _all_cell_texts(workbook) -> List[str]:
    texts: List[str] = []
    for ws in workbook.worksheets:
        for row in ws.iter_rows(values_only=True):
            for value in row:
                t = _clean_text(value)
                if t:
                    texts.append(t)
    return texts


def _extract_metadata(texts: Iterable[str], filename: str) -> Dict[str, str]:
    joined = "\n".join(texts)
    metadata: Dict[str, str] = {
        "company_name": "",
        "course_name": "",
        "course_round": "",
        "schedule": "",
    }

    course_patterns = [
        r"교육과정\s*[:：]\s*([^\n]+)",
        r"과정명\s*[:：]\s*([^\n]+)",
    ]
    for pattern in course_patterns:
        m = re.search(pattern, joined)
        if m:
            raw = _clean_text(m.group(1))
            raw = re.sub(r"^[■●\-]\s*", "", raw)
            metadata["course_name"] = raw
            break

    date_patterns = [
        r"교육일시\s*[:：]\s*([^\n]+)",
        r"교육일정\s*[:：]\s*([^\n]+)",
        r"(20\d{2}[./-]\d{1,2}[./-]\d{1,2}\.?\s*[~～-]\s*20?\d{0,2}[./-]?\d{1,2}[./-]\d{1,2}\.?)",
    ]
    for pattern in date_patterns:
        m = re.search(pattern, joined)
        if m:
            metadata["schedule"] = _clean_text(m.group(1))
            break

    round_match = re.search(r"\((\d+)차(?:수)?\)|\b(\d+)차수\b|과정\s*\((\d+)차\)", metadata["course_name"] or filename)
    if round_match:
        val = next((g for g in round_match.groups() if g), "")
        metadata["course_round"] = f"{val}차수" if val else ""

    clean_filename = Path(filename).stem
    clean_filename = re.sub(r"\([^)]*로우데이터[^)]*\)", "", clean_filename, flags=re.I)
    clean_filename = re.sub(r"_?결과본$", "", clean_filename)
    clean_filename = re.sub(r"교육 만족도 (설문|조사).*$", "", clean_filename)
    clean_filename = _clean_text(clean_filename)

    if not metadata["course_name"]:
        metadata["course_name"] = clean_filename

    company_candidates = [
        r"^\(([^)]+)\)",
        r"^([A-Za-z][A-Za-z0-9 &.-]{1,25})\s+",
        r"^(한국[가-힣A-Za-z0-9]+)",
    ]
    for pattern in company_candidates:
        m = re.search(pattern, Path(filename).stem)
        if m:
            metadata["company_name"] = _clean_text(m.group(1))
            break

    if metadata["course_name"]:
        metadata["course_name"] = re.sub(r"^\(([^)]+)\)\s*", "", metadata["course_name"])
        metadata["course_name"] = re.sub(r"\s*과정\s*\(\d+차\)\s*$", "", metadata["course_name"])

    return metadata


def _dedupe_instructors(questions: Sequence[ObjectiveQuestion]) -> List[str]:
    found = []
    seen = set()
    for q in questions:
        name = q.instructor_name.strip()
        if name and name not in seen:
            seen.add(name)
            found.append(name)
    return found


def is_no_opinion(text: str) -> bool:
    core = re.sub(r"[\s.·,!?~]+", "", str(text or "")).lower()
    return core in NO_OPINION_CORES


def parse_excel(file_bytes: bytes, filename: str) -> ReportData:
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=False)
    candidates: List[Dict[str, Any]] = []
    diagnostics: List[str] = []

    for ws in workbook.worksheets:
        matrix = _sheet_matrix(ws)
        candidate = _sheet_candidate(matrix, ws.title)
        if candidate:
            candidates.append(candidate)
            diagnostics.append(
                f"'{ws.title}' 시트: {candidate['orientation']}, 객관식 {len(candidate['objective'])}개, "
                f"주관식 {len(candidate['subjective'])}개, 응답자 추정 {candidate['response_count']}명"
            )

    if not candidates:
        raise ValueError("설문 문항과 응답 구조를 찾지 못했습니다. 문항이 첫 행 또는 첫 열에 있는지 확인해 주세요.")

    best = max(candidates, key=lambda item: item["quality"])
    metadata = _extract_metadata(_all_cell_texts(workbook), filename)
    objective_questions: List[ObjectiveQuestion] = best["objective"]
    subjective_questions: List[SubjectiveQuestion] = best["subjective"]
    instructors = _dedupe_instructors(objective_questions)

    course_name = metadata["course_name"]
    report_title = f"{course_name} 결과보고서" if course_name else "교육만족도 결과보고서"

    diagnostics.insert(0, f"분석에 사용한 시트: '{best['sheet_name']}'")
    if not metadata["schedule"]:
        diagnostics.append("교육일정은 엑셀에서 찾지 못해 공란으로 두었습니다.")
    if not instructors:
        diagnostics.append("강사명은 문항에서 찾지 못해 공란으로 두었습니다.")

    return ReportData(
        source_filename=filename,
        company_name=metadata["company_name"],
        report_title=report_title,
        course_name=course_name,
        course_round=metadata["course_round"],
        schedule=metadata["schedule"],
        delivery_method="",
        target_text="",
        total_participants=None,
        response_count=best["response_count"],
        objective="",
        instructors=instructors,
        objective_questions=objective_questions,
        subjective_questions=subjective_questions,
        diagnostics=diagnostics,
    )
