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


def _chunks(values: List[int], size: int) -> List[List[int]]:
    return [values[i : i + size] for i in range(0, len(values), size)] or [[]]


def build_slide_plan(
    report: ReportData,
    subjective_per_slide: int = 5,
    photo_count: int = 0,
    photos_per_slide: int = 6,
    survey_questions_per_slide: int = 8,
    summary_questions_per_slide: int = 7,
) -> List[SlideItem]:
    """Create the shared HTML/PPT slide order for V2."""
    slides: List[SlideItem] = [
        SlideItem("cover", "cover", "표지", {}),
        SlideItem("toc", "toc", "목차", {}),
        SlideItem("section-overview", "section", "1. 교육 개요", {"section_no": "01", "items": ["교육 개요", "커리큘럼"]}),
        SlideItem("overview", "overview", "교육 개요", {}),
        SlideItem("curriculum", "curriculum", "커리큘럼", {}),
        SlideItem(
            "section-survey",
            "section",
            "2. 만족도 통계",
            {
                "section_no": "02",
                "items": ["설문 구성", "만족도 요약", "공통문항", "강사 비교", "선택형", "주관식", "종합 시사점"],
            },
        ),
    ]

    all_questions = list(report.objective_questions) + list(report.choice_questions) + list(report.subjective_questions)
    survey_pages = max(1, (len(all_questions) + survey_questions_per_slide - 1) // survey_questions_per_slide)
    for page_index in range(survey_pages):
        start = page_index * survey_questions_per_slide
        end = min(start + survey_questions_per_slide, len(all_questions))
        suffix = f" ({page_index + 1}/{survey_pages})" if survey_pages > 1 else ""
        slides.append(
            SlideItem(
                f"survey-structure-{page_index}",
                "survey_structure",
                f"설문 구성{suffix}",
                {"page_index": page_index, "start": start, "end": end},
            )
        )

    # Summary charts remain separated internally by scale range so axes stay
    # accurate, but scale labels are not exposed in the slide title. Wide
    # numeric scales such as 0~10 use one question per page and show the full
    # response distribution instead of only an average bar.
    scale_groups: Dict[str, List[int]] = {}
    for index, question in enumerate(report.objective_questions):
        scale_groups.setdefault(question.scale_key, []).append(index)

    summary_specs: List[Dict[str, Any]] = []
    for scale_key, indices in scale_groups.items():
        first = report.objective_questions[indices[0]] if indices else None
        is_wide_scale = bool(first and len(first.scale_values) >= 8)
        page_size = 1 if is_wide_scale else summary_questions_per_slide
        for page_indices in _chunks(indices, page_size):
            summary_specs.append(
                {
                    "scale_key": scale_key,
                    "question_indices": page_indices,
                    "summary_mode": "distribution" if is_wide_scale else "average",
                }
            )

    summary_page_count = len(summary_specs)
    for page_index, spec in enumerate(summary_specs):
        suffix = f" ({page_index + 1}/{summary_page_count})" if summary_page_count > 1 else ""
        slides.append(
            SlideItem(
                f"summary-{spec['scale_key']}-{page_index}",
                "summary",
                f"만족도 요약{suffix}",
                {
                    **spec,
                    "page_index": page_index,
                    "page_count": summary_page_count,
                },
            )
        )

    common_indices = [idx for idx, question in enumerate(report.objective_questions) if not question.instructor_name]
    for order, index in enumerate(common_indices, start=1):
        question = report.objective_questions[index]
        slides.append(
            SlideItem(
                f"objective-{question.id}",
                "objective",
                f"{question.section_label} · {order}",
                {"question_index": index, "display_order": order, "group": "common"},
            )
        )

    if len(report.instructors) >= 2:
        lecturer_pages = _chunks(list(range(len(report.instructors))), 5)
        for page_index, lecturer_indices in enumerate(lecturer_pages):
            suffix = f" ({page_index + 1}/{len(lecturer_pages)})" if len(lecturer_pages) > 1 else ""
            slides.append(
                SlideItem(
                    f"lecturer-comparison-{page_index}",
                    "lecturer_comparison",
                    f"강사 비교{suffix}",
                    {"lecturer_indices": lecturer_indices, "page_index": page_index},
                )
            )

    # Instructor details remain grouped by lecturer instead of raw column order.
    for lecturer_index, lecturer_name in enumerate(report.instructors):
        lecturer_question_indices = [
            idx for idx, question in enumerate(report.objective_questions) if question.instructor_name == lecturer_name
        ]
        for local_order, question_index in enumerate(lecturer_question_indices, start=1):
            question = report.objective_questions[question_index]
            slides.append(
                SlideItem(
                    f"lecturer-{lecturer_index}-{question.id}",
                    "objective",
                    f"{lecturer_name} · {question.instructor_metric or local_order}",
                    {
                        "question_index": question_index,
                        "display_order": local_order,
                        "group": "lecturer",
                        "lecturer_name": lecturer_name,
                    },
                )
            )

    for choice_index, question in enumerate(report.choice_questions):
        slides.append(
            SlideItem(
                f"choice-{question.id}",
                "choice",
                f"{question.section_label} · {choice_index + 1}",
                {"question_index": choice_index},
            )
        )

    for q_index, question in enumerate(report.subjective_questions):
        chunks = [question.answers[i : i + subjective_per_slide] for i in range(0, len(question.answers), subjective_per_slide)] or [[]]
        for chunk_index, chunk in enumerate(chunks):
            suffix = f" ({chunk_index + 1}/{len(chunks)})" if len(chunks) > 1 else ""
            slides.append(
                SlideItem(
                    f"subjective-{question.id}-{chunk_index}",
                    "subjective",
                    f"{question.section_label}{suffix}",
                    {
                        "question_index": q_index,
                        "chunk_index": chunk_index,
                        "answers": chunk,
                        "page_suffix": suffix,
                    },
                )
            )

    slides.append(SlideItem("insights", "insights", "종합 시사점", {}))

    slides.append(SlideItem("section-photos", "section", "3. 현장 사진", {"section_no": "03", "items": ["교육 현장 스케치"]}))
    photo_pages = max(1, (photo_count + photos_per_slide - 1) // photos_per_slide) if photo_count else 1
    for photo_page in range(photo_pages):
        suffix = f" ({photo_page + 1}/{photo_pages})" if photo_pages > 1 else ""
        slides.append(
            SlideItem(
                f"photos-{photo_page}",
                "photos",
                f"현장 사진{suffix}",
                {"page_index": photo_page, "page_suffix": suffix},
            )
        )

    slides.append(SlideItem("thanks", "thanks", "THANK YOU", {}))
    return slides
