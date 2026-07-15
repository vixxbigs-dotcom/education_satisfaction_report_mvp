from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List

from .models import ReportData


@dataclass
class SlideItem:
    id: str
    kind: str
    title: str
    payload: Dict[str, Any]


def build_slide_plan(
    report: ReportData,
    subjective_per_slide: int = 5,
    photo_count: int = 0,
    photos_per_slide: int = 6,
    survey_questions_per_slide: int = 8,
    summary_questions_per_slide: int = 7,
) -> List[SlideItem]:
    """Create the exact slide order used by both HTML preview and PPT export.

    Dynamic pages are planned here so every exported page is also available in
    the web preview. This prevents the preview/export mismatch that occurred
    when the renderer independently split long survey or summary sections.
    """
    slides: List[SlideItem] = [
        SlideItem("cover", "cover", "표지", {}),
        SlideItem("toc", "toc", "목차", {}),
        SlideItem(
            "section-overview",
            "section",
            "1. 교육 개요",
            {"section_no": "01", "items": ["교육 개요", "커리큘럼"]},
        ),
        SlideItem("overview", "overview", "교육 개요", {}),
        SlideItem("curriculum", "curriculum", "커리큘럼", {}),
        SlideItem(
            "section-survey",
            "section",
            "2. 만족도 통계",
            {"section_no": "02", "items": ["설문 구성", "객관식 설문 결과", "주관식 설문 결과"]},
        ),
    ]

    all_questions = list(report.objective_questions) + list(report.subjective_questions)
    survey_pages = max(1, (len(all_questions) + survey_questions_per_slide - 1) // survey_questions_per_slide)
    for page_index in range(survey_pages):
        start = page_index * survey_questions_per_slide
        end = min(start + survey_questions_per_slide, len(all_questions))
        suffix = f" ({page_index + 1}/{survey_pages})" if survey_pages > 1 else ""
        slides.append(
            SlideItem(
                id=f"survey-structure-{page_index}",
                kind="survey_structure",
                title=f"설문 구성{suffix}",
                payload={"page_index": page_index, "start": start, "end": end},
            )
        )

    summary_pages = max(
        1,
        (len(report.objective_questions) + summary_questions_per_slide - 1) // summary_questions_per_slide,
    )
    for page_index in range(summary_pages):
        start = page_index * summary_questions_per_slide
        end = min(start + summary_questions_per_slide, len(report.objective_questions))
        suffix = f" ({page_index + 1}/{summary_pages})" if summary_pages > 1 else ""
        slides.append(
            SlideItem(
                id=f"summary-{page_index}",
                kind="summary",
                title=f"객관식 설문 결과 요약{suffix}",
                payload={"page_index": page_index, "start": start, "end": end},
            )
        )

    for idx, question in enumerate(report.objective_questions, start=1):
        slides.append(
            SlideItem(
                id=f"objective-{question.id}",
                kind="objective",
                title=f"{question.section_label} · {idx}",
                payload={"question_index": idx - 1},
            )
        )

    for q_index, question in enumerate(report.subjective_questions):
        chunks = [
            question.answers[i : i + subjective_per_slide]
            for i in range(0, len(question.answers), subjective_per_slide)
        ] or [[]]
        for chunk_index, chunk in enumerate(chunks):
            suffix = f" ({chunk_index + 1}/{len(chunks)})" if len(chunks) > 1 else ""
            slides.append(
                SlideItem(
                    id=f"subjective-{question.id}-{chunk_index}",
                    kind="subjective",
                    title=f"{question.section_label}{suffix}",
                    payload={
                        "question_index": q_index,
                        "chunk_index": chunk_index,
                        "answers": chunk,
                        "page_suffix": suffix,
                    },
                )
            )

    photo_pages = max(1, (photo_count + photos_per_slide - 1) // photos_per_slide) if photo_count else 1
    for photo_page in range(photo_pages):
        suffix = f" ({photo_page + 1}/{photo_pages})" if photo_pages > 1 else ""
        slides.append(
            SlideItem(
                id=f"photos-{photo_page}",
                kind="photos",
                title=f"현장 사진{suffix}",
                payload={"page_index": photo_page, "page_suffix": suffix},
            )
        )

    slides.append(SlideItem("thanks", "thanks", "THANK YOU", {}))
    return slides
