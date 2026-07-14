from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional


@dataclass
class ObjectiveQuestion:
    id: str
    number: int
    section_label: str
    question: str
    counts: List[int]
    scale_labels: List[str]
    average: float
    valid_responses: int
    category: str = "기타"
    instructor_name: str = ""
    course_module: str = ""

    def recalculate(self) -> None:
        point_count = len(self.scale_labels)
        total = sum(max(0, int(v)) for v in self.counts)
        self.valid_responses = total
        if total == 0:
            self.average = 0.0
            return
        score_sum = sum((point_count - idx) * max(0, int(count)) for idx, count in enumerate(self.counts))
        self.average = round(score_sum / total, 2)


@dataclass
class SubjectiveQuestion:
    id: str
    number: int
    section_label: str
    question: str
    answers: List[str] = field(default_factory=list)
    category: str = "주관식"


@dataclass
class CurriculumRow:
    day: str = ""
    time: str = ""
    content: str = ""
    instructor: str = ""


@dataclass
class ReportData:
    source_filename: str
    company_name: str = ""
    report_title: str = ""
    course_name: str = ""
    course_round: str = ""
    schedule: str = ""
    delivery_method: str = ""
    target_text: str = ""
    total_participants: Optional[int] = None
    response_count: int = 0
    objective: str = ""
    instructors: List[str] = field(default_factory=list)
    curriculum: List[CurriculumRow] = field(default_factory=list)
    objective_questions: List[ObjectiveQuestion] = field(default_factory=list)
    subjective_questions: List[SubjectiveQuestion] = field(default_factory=list)
    diagnostics: List[str] = field(default_factory=list)

    @property
    def response_rate(self) -> Optional[int]:
        if not self.total_participants:
            return None
        if self.total_participants <= 0:
            return None
        return round(self.response_count / self.total_participants * 100)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def objective_from_dict(data: Dict[str, Any]) -> ObjectiveQuestion:
    q = ObjectiveQuestion(**data)
    q.recalculate()
    return q


def subjective_from_dict(data: Dict[str, Any]) -> SubjectiveQuestion:
    return SubjectiveQuestion(**data)


def curriculum_from_dict(data: Dict[str, Any]) -> CurriculumRow:
    return CurriculumRow(**data)


def report_from_dict(data: Dict[str, Any]) -> ReportData:
    copied = dict(data)
    copied["objective_questions"] = [objective_from_dict(v) for v in data.get("objective_questions", [])]
    copied["subjective_questions"] = [subjective_from_dict(v) for v in data.get("subjective_questions", [])]
    copied["curriculum"] = [curriculum_from_dict(v) for v in data.get("curriculum", [])]
    return ReportData(**copied)
