from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class ObjectiveQuestion:
    """Numeric/Likert scale question.

    ``counts`` and ``scale_labels`` are always stored from high score to low
    score for compatibility with the V1 horizontal distribution chart.
    ``scale_values`` mirrors that order and removes the old assumption that
    the number of labels is the score itself (important for 0~10 scales).
    """

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
    scale_min: int = 1
    scale_max: int = 5
    scale_values: List[int] = field(default_factory=list)
    scale_type: str = "likert"
    instructor_metric: str = ""

    def __post_init__(self) -> None:
        if not self.scale_values:
            self.scale_values = list(range(int(self.scale_max), int(self.scale_min) - 1, -1))
        if len(self.scale_values) != len(self.counts):
            self.scale_values = list(range(int(self.scale_max), int(self.scale_min) - 1, -1))[: len(self.counts)]

    def recalculate(self) -> None:
        total = sum(max(0, int(v)) for v in self.counts)
        self.valid_responses = total
        if total == 0:
            self.average = 0.0
            return
        values = self.scale_values or list(range(self.scale_max, self.scale_min - 1, -1))
        score_sum = sum(float(score) * max(0, int(count)) for score, count in zip(values, self.counts))
        self.average = round(score_sum / total, 2)

    @property
    def scale_key(self) -> str:
        return f"{self.scale_min}-{self.scale_max}"


@dataclass
class ChoiceQuestion:
    id: str
    number: int
    section_label: str
    question: str
    options: List[str]
    counts: List[int]
    percentages: List[float]
    selection_type: str = "single"  # single | multiple
    valid_responses: int = 0
    category: str = "choice"

    def recalculate(self) -> None:
        self.counts = [max(0, int(v)) for v in self.counts]
        denominator = max(0, int(self.valid_responses))
        if denominator <= 0:
            denominator = sum(self.counts) if self.selection_type == "single" else 0
            self.valid_responses = denominator
        self.percentages = [round(count / denominator * 100, 1) if denominator else 0.0 for count in self.counts]


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
    choice_questions: List[ChoiceQuestion] = field(default_factory=list)
    subjective_questions: List[SubjectiveQuestion] = field(default_factory=list)
    insights: List[str] = field(default_factory=list)
    diagnostics: List[str] = field(default_factory=list)

    @property
    def response_rate(self) -> Optional[int]:
        if not self.total_participants or self.total_participants <= 0:
            return None
        return round(self.response_count / self.total_participants * 100)

    @property
    def common_objective_questions(self) -> List[ObjectiveQuestion]:
        return [q for q in self.objective_questions if not q.instructor_name]

    @property
    def instructor_objective_questions(self) -> List[ObjectiveQuestion]:
        return [q for q in self.objective_questions if q.instructor_name]

    @property
    def lecturers(self) -> Dict[str, List[ObjectiveQuestion]]:
        grouped: Dict[str, List[ObjectiveQuestion]] = {}
        for question in self.instructor_objective_questions:
            grouped.setdefault(question.instructor_name, []).append(question)
        return grouped

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def objective_from_dict(data: Dict[str, Any]) -> ObjectiveQuestion:
    q = ObjectiveQuestion(**data)
    q.recalculate()
    return q


def choice_from_dict(data: Dict[str, Any]) -> ChoiceQuestion:
    q = ChoiceQuestion(**data)
    q.recalculate()
    return q


def subjective_from_dict(data: Dict[str, Any]) -> SubjectiveQuestion:
    return SubjectiveQuestion(**data)


def curriculum_from_dict(data: Dict[str, Any]) -> CurriculumRow:
    return CurriculumRow(**data)


def report_from_dict(data: Dict[str, Any]) -> ReportData:
    copied = dict(data)
    copied["objective_questions"] = [objective_from_dict(v) for v in data.get("objective_questions", [])]
    copied["choice_questions"] = [choice_from_dict(v) for v in data.get("choice_questions", [])]
    copied["subjective_questions"] = [subjective_from_dict(v) for v in data.get("subjective_questions", [])]
    copied["curriculum"] = [curriculum_from_dict(v) for v in data.get("curriculum", [])]
    copied.setdefault("insights", [])
    return ReportData(**copied)
