from __future__ import annotations

import base64
import html
import math
from pathlib import Path
from typing import List, Sequence

from .font_manager import get_preview_font_css
from .models import ChoiceQuestion, ObjectiveQuestion, ReportData
from .slide_plan import SlideItem
from .text_utils import strip_leading_question_numbers

ORANGE = "#ED702C"
ORANGE_DARK = "#D85C1E"
DARK = "#444444"
MUTED = "#6F6F6F"
PURPLE = "#A02B98"
BLUE = "#1B9FD0"
LIGHT_GRID = "#E5E5E5"
PALETTE = [ORANGE, PURPLE, BLUE, "#F39A66", "#7C8DB5", "#58A98C", "#C9A13B", "#9B9B9B", "#D87B9A", "#7B6FB0"]
ASSET_DIR = Path(__file__).resolve().parent.parent / "assets"
SUBJECTIVE_PER_SLIDE = 5
FONT_FACE_CSS = get_preview_font_css()


def _data_uri(filename: str) -> str:
    path = ASSET_DIR / filename
    if not path.exists():
        return ""
    mime = "image/png"
    payload = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{payload}"


SECTION_BG_URI = _data_uri("multicampus_section_background.png")
LOGO_URI = _data_uri("multicampus_logo.png")


def _e(value) -> str:
    return html.escape(str(value or ""))


def _base(content: str, class_name: str = "") -> str:
    return f'<div class="ppt-stage"><div class="ppt-slide {class_name}">{content}</div></div>'


def _content_header(main_title: str, section_no: str, section_title: str) -> str:
    return f"""
    <div class="content-main-title">{_e(main_title)}</div>
    <div class="content-rule"></div>
    <div class="content-section-tag"><span>{_e(section_no)}</span><b>{_e(section_title)}</b></div>
    """


def _content_footer() -> str:
    if LOGO_URI:
        return f'<img class="content-logo" src="{LOGO_URI}" alt="multicampus" />'
    return '<div class="content-logo-text">multicampus</div>'


def _section_background_content(title_html: str, extra_html: str = "") -> str:
    return _base(f'<div class="section-right-title">{title_html}</div>{extra_html}', "section-background")


def _cover(report: ReportData) -> str:
    title1 = report.cover_title1 or report.course_name or "과정 이름"
    title2 = report.cover_title2 or "결과보고서"
    top_label = report.cover_top_label or "결과보고서"
    company = f'<div class="section-company">{_e(report.company_name)}</div>' if report.company_name else ""
    schedule = f'<div class="section-schedule">{_e(report.schedule)}</div>' if report.schedule else ""
    return _section_background_content(
        f'<div class="cover-top-label">{_e(top_label)}</div>'
        f'<div class="cover-course">{_e(title1)}</div><div class="cover-report">{_e(title2)}</div>',
        company + schedule,
    )


def _toc() -> str:
    return _section_background_content(
        '<div class="section-heading">목차</div>',
        """
        <div class="toc-right-list">
          <div><span>01</span><b>교육 개요</b><em>교육 개요 · 커리큘럼</em></div>
          <div><span>02</span><b>만족도 통계</b><em>요약 · 공통문항 · 강사 · 선택형 · 주관식 · 시사점</em></div>
          <div><span>03</span><b>현장 사진</b><em>교육 현장 스케치</em></div>
        </div>
        """,
    )


def _section(slide: SlideItem) -> str:
    title = slide.title.replace("1. ", "").replace("2. ", "").replace("3. ", "")
    items = " · ".join(_e(v) for v in slide.payload.get("items", []))
    return _section_background_content(
        f'<div class="section-heading"><small>{_e(slide.payload.get("section_no"))}</small>{_e(title)}</div>',
        f'<div class="section-item-line">{items}</div>',
    )


def _overview(report: ReportData) -> str:
    target = report.target_text
    if report.total_participants:
        target = f"{target} {report.total_participants}명".strip()
    rows = [
        ("과정명", report.course_name), ("교육일정", report.schedule), ("교육방식", report.delivery_method),
        ("교육 대상", target), ("교육 목표", report.objective),
    ]
    body = "".join(f'<div class="overview-label">{_e(k)}</div><div class="overview-value">{_e(v)}</div>' for k, v in rows)
    return _base(f'{_content_header("1) 교육 개요", "01", "교육 개요")}<div class="overview-grid">{body}</div>{_content_footer()}')


def _curriculum(report: ReportData) -> str:
    body = "".join(
        f"<tr><td>{_e(r.day)}</td><td>{_e(r.time)}</td><td>{_e(r.content)}</td><td>{_e(r.instructor)}</td></tr>"
        for r in report.curriculum[:8]
    ) or '<tr><td colspan="4" class="empty-cell"></td></tr>'
    return _base(
        f"{_content_header('2) 커리큘럼', '01', '교육 개요')}"
        f'<table class="curriculum-table"><thead><tr><th>Day</th><th>시간</th><th>교육 내용</th><th>강사/비고</th></tr></thead><tbody>{body}</tbody></table>'
        f"{_content_footer()}"
    )


def _all_questions(report: ReportData):
    return list(report.objective_questions) + list(report.choice_questions) + list(report.subjective_questions)


def _survey_structure(report: ReportData, slide: SlideItem) -> str:
    questions = _all_questions(report)
    start = int(slide.payload.get("start", 0)); end = int(slide.payload.get("end", len(questions)))
    chunk = questions[start:end]
    rows = "".join(
        f"<tr><td>{_e(q.section_label)}</td><td>{start + idx}. {_e(strip_leading_question_numbers(q.question))}</td></tr>"
        for idx, q in enumerate(chunk, start=1)
    ) or '<tr><td colspan="2" class="empty-cell"></td></tr>'
    suffix = f" ({slide.payload.get('page_index', 0) + 1})" if len(questions) > 8 else ""
    return _base(
        f"{_content_header(f'1) 설문 구성{suffix}', '02', '만족도 통계')}"
        f'<table class="survey-table"><thead><tr><th>항목</th><th>문항</th></tr></thead><tbody>{rows}</tbody></table>'
        f"{_content_footer()}"
    )


def _summary_label(q: ObjectiveQuestion) -> str:
    # Compatibility with reports parsed before ``summary_label`` was added
    # to ObjectiveQuestion. Old session objects may not have the attribute.
    custom_label = getattr(q, "summary_label", "")
    if custom_label:
        return _e(custom_label).replace("\n", "<br/>")
    instructor_name = getattr(q, "instructor_name", "")
    if instructor_name:
        metric = getattr(q, "instructor_metric", "") or "만족도"
        return f"{_e(instructor_name)}<br/>{_e(metric)}"
    return _e(getattr(q, "section_label", "만족도"))


def _count_axis(max_count: int) -> tuple[float, List[float]]:
    """Return a clean response-count axis with four equal intervals."""
    raw = max(1, int(max_count))
    if raw <= 4:
        axis_max = 4
    elif raw <= 8:
        axis_max = 8
    elif raw <= 10:
        axis_max = 10
    else:
        axis_max = int(math.ceil(raw / 5.0) * 5)
    ticks = [axis_max * ratio / 4 for ratio in range(4, -1, -1)]
    return float(axis_max), ticks


def _format_tick(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else f"{value:.1f}"


def _summary_distribution(report: ReportData, slide: SlideItem, q: ObjectiveQuestion, suffix: str) -> str:
    max_count = max(max(q.counts) if q.counts else 0, 1)
    axis_max, ticks = _count_axis(max_count)
    valid = max(0, int(q.valid_responses))
    if report.total_participants:
        count_line = (
            f"수강인원 총 {report.total_participants}명 중 {report.response_count}명 응답"
            f" ｜ 해당 문항 유효응답 {valid}명"
        )
    elif report.response_count != valid:
        count_line = f"총 {report.response_count}명 응답 ｜ 해당 문항 유효응답 {valid}명"
    else:
        count_line = f"유효 응답 {valid}명"

    columns = []
    for value, count in reversed(list(zip(q.scale_values, q.counts))):
        percentage = (count / valid * 100.0) if valid else 0.0
        height = max(0.0, min(100.0, count / axis_max * 100.0))
        bar = f'<div class="dist-bar" style="height:{height:.1f}%"></div>' if count > 0 else ""
        columns.append(
            f'<div class="dist-col">'
            f'<div class="dist-value" style="bottom:calc({height:.1f}% + .35em)"><b>{count}명</b><small>({percentage:.0f}%)</small></div>'
            f'<div class="dist-bar-area">{bar}</div><div class="dist-score">{value}</div></div>'
        )
    tick_html = "".join(f"<span>{_format_tick(value)}</span>" for value in ticks)
    question = strip_leading_question_numbers(q.question)
    return _base(
        f"{_content_header(f'2) 만족도 요약{suffix}', '02', '만족도 통계')}"
        f'<div class="summary-note"><b>■ 만족도 요약</b><span>- {_e(count_line)}</span></div>'
        f'<div class="summary-dist-question">■ {_e(q.section_label)} &nbsp; {_e(question)}</div>'
        f'<div class="summary-dist-average">평균 <b>{q.average:.1f}점</b></div>'
        f'<div class="summary-dist-chart"><div class="summary-dist-y-labels">{tick_html}</div>'
        f'<div class="summary-dist-plot"><div class="dist-columns">{"".join(columns)}</div></div>'
        f'<div class="summary-dist-x-title">점수</div></div>{_content_footer()}'
    )


def _summary(report: ReportData, slide: SlideItem) -> str:
    indices = slide.payload.get("question_indices", [])
    questions = [report.objective_questions[i] for i in indices]
    suffix = f" ({slide.payload.get('page_index', 0) + 1}/{slide.payload.get('page_count', 1)})" if slide.payload.get("page_count", 1) > 1 else ""
    if slide.payload.get("summary_mode") == "distribution" and questions:
        return _summary_distribution(report, slide, questions[0], suffix)

    scale_min = questions[0].scale_min if questions else 0
    scale_max = questions[0].scale_max if questions else 5
    span = max(1, scale_max - scale_min)
    count_line = (
        f"수강인원 총 {report.total_participants}명 중 {report.response_count}명 응답 ｜ 응답률 {report.response_rate}%"
        if report.total_participants else f"총 {report.response_count}명 응답"
    )
    bars = []
    for q in questions:
        height = max(0.0, min(100.0, ((q.average - scale_min) / span) * 100.0))
        bars.append(
            f'<div class="summary-bar-column"><div class="summary-value">{q.average:.1f}</div>'
            f'<div class="summary-bar" style="height:{height:.1f}%"></div>'
            f'<div class="summary-x-label">{_summary_label(q)}</div></div>'
        )
    ticks = list(range(scale_max, scale_min - 1, -1))
    tick_html = "".join(f"<span>{value:.1f}</span>" for value in ticks)
    return _base(
        f"{_content_header(f'2) 만족도 요약{suffix}', '02', '만족도 통계')}"
        f'<div class="summary-note"><b>■ 만족도 요약</b><span>- {_e(count_line)}</span></div>'
        f'<div class="summary-chart-wrap"><div class="summary-y-labels">{tick_html}</div>'
        f'<div class="summary-plot"><div class="summary-bars">{"".join(bars)}</div></div></div>{_content_footer()}'
    )


def _objective_horizontal(q: ObjectiveQuestion) -> str:
    max_count = max(q.valid_responses, max(q.counts) if q.counts else 0, 1)
    rows = []
    for idx, (label, count) in enumerate(zip(q.scale_labels, q.counts)):
        width = count / max_count * 100 if max_count else 0
        fill = "" if count <= 0 else (
            f'<div class="objective-fill" style="width:{width:.1f}%;background:{PALETTE[idx % len(PALETTE)]}">'
            f'<span class="objective-count">{count}</span></div>'
        )
        rows.append(f'<div class="objective-row"><div class="objective-label">{_e(label)}</div><div class="objective-track">{fill}</div></div>')
    ticks = "".join(f"<span>{i}</span>" for i in range(max_count + 1))
    return f'<div class="objective-chart">{"".join(rows)}<div class="objective-ticks">{ticks}</div></div>'


def _objective_histogram(q: ObjectiveQuestion) -> str:
    max_count = max(max(q.counts) if q.counts else 0, 1)
    columns = []
    # Histogram reads naturally from low to high, while model storage is high-to-low.
    for value, count in reversed(list(zip(q.scale_values, q.counts))):
        height = count / max_count * 100 if max_count else 0
        columns.append(
            f'<div class="hist-col"><div class="hist-count">{count}</div>'
            f'<div class="hist-bar-wrap"><div class="hist-bar" style="height:{height:.1f}%"></div></div>'
            f'<div class="hist-label">{value}</div></div>'
        )
    return f'<div class="histogram"><div class="hist-y-title">응답 수</div><div class="hist-columns">{"".join(columns)}</div><div class="hist-x-title">점수</div></div>'


def _objective(report: ReportData, slide: SlideItem) -> str:
    index = int(slide.payload["question_index"]); q = report.objective_questions[index]
    group = slide.payload.get("group", "common")
    title = "3) 강사별 상세" if group == "lecturer" else "3) 공통문항"
    chart = _objective_histogram(q) if len(q.scale_values) >= 8 else _objective_horizontal(q)
    question_no = int(slide.payload.get("display_order", index + 1))
    meta = f'<div class="scale-average">평균 <b>{q.average:.1f}</b> / {q.scale_max}점</div>'
    return _base(
        f"{_content_header(title, '02', '만족도 통계')}"
        f'<div class="question-title">■ {_e(q.section_label)} &nbsp; {question_no}. {_e(strip_leading_question_numbers(q.question))}</div>'
        f"{meta}{chart}{_content_footer()}"
    )


def _lecturer_comparison(report: ReportData, slide: SlideItem) -> str:
    names = [report.instructors[i] for i in slide.payload.get("lecturer_indices", [])]
    metrics = []
    for name in names:
        for question in report.lecturers.get(name, []):
            metric = question.instructor_metric or "기타"
            if metric not in metrics:
                metrics.append(metric)
    headers = "".join(f"<th>{_e(metric)}</th>" for metric in metrics)
    rows = []
    for name in names:
        metric_values = {}
        for question in report.lecturers.get(name, []):
            metric_values.setdefault(question.instructor_metric or "기타", []).append(question.average)
        by_metric = {metric: sum(values) / len(values) for metric, values in metric_values.items()}
        cells = "".join(f'<td>{by_metric.get(metric, "-") if metric not in by_metric else f"{by_metric[metric]:.1f}"}</td>' for metric in metrics)
        average = sum(q.average for q in report.lecturers.get(name, [])) / len(report.lecturers.get(name, [])) if report.lecturers.get(name, []) else 0
        rows.append(f'<tr><td>{_e(name)}</td>{cells}<td class="lecturer-total">{average:.1f}</td></tr>')
    return _base(
        f"{_content_header('3) 강사 비교', '02', '만족도 통계')}"
        f'<div class="comparison-note">■ 강사별 문항 평균 비교</div>'
        f'<table class="comparison-table"><thead><tr><th>강사</th>{headers}<th>종합 평균</th></tr></thead><tbody>{"".join(rows)}</tbody></table>'
        f"{_content_footer()}"
    )


def _pie_style(question: ChoiceQuestion) -> str:
    total = sum(question.counts) or 1
    cursor = 0.0; segments = []
    for idx, count in enumerate(question.counts):
        start = cursor; cursor += count / total * 100
        segments.append(f"{PALETTE[idx % len(PALETTE)]} {start:.2f}% {cursor:.2f}%")
    return "conic-gradient(" + ",".join(segments) + ")"


def _choice(report: ReportData, slide: SlideItem) -> str:
    index = int(slide.payload["question_index"]); q = report.choice_questions[index]
    question_title = f"■ {index + 1}. {_e(strip_leading_question_numbers(q.question))}"
    if q.selection_type == "single":
        legend = "".join(
            f'<div><i style="background:{PALETTE[idx % len(PALETTE)]}"></i><span>{_e(option)}</span><b>{count}명 · {pct:.1f}%</b></div>'
            for idx, (option, count, pct) in enumerate(zip(q.options, q.counts, q.percentages))
        )
        body = f'<div class="choice-pie" style="background:{_pie_style(q)}"></div><div class="choice-legend">{legend}</div>'
        subtitle = "4) 단일선택 문항"
    else:
        max_count = max(q.counts) if q.counts else 1
        rows = "".join(
            f'<div class="multi-row"><div class="multi-label">{_e(option)}</div><div class="multi-track"><div class="multi-fill" style="width:{count / max_count * 100 if max_count else 0:.1f}%"></div></div><div class="multi-value">{count}명 · {pct:.1f}%</div></div>'
            for option, count, pct in zip(q.options, q.counts, q.percentages)
        )
        body = f'<div class="multi-chart">{rows}<div class="multi-foot">※ 복수선택 문항은 비율 합계가 100%를 초과할 수 있습니다.</div></div>'
        subtitle = "4) 복수선택 문항"
    return _base(f'{_content_header(subtitle, "02", "만족도 통계")}<div class="choice-question">{question_title}</div>{body}{_content_footer()}')


def _subjective(report: ReportData, slide: SlideItem) -> str:
    index = int(slide.payload["question_index"]); q = report.subjective_questions[index]
    start = int(slide.payload.get("chunk_index", 0)) * SUBJECTIVE_PER_SLIDE
    answers: Sequence[str] = q.answers[start : start + SUBJECTIVE_PER_SLIDE]
    items = "".join(f'<li>{_e(answer)}</li>' for answer in answers) or '<li class="empty-opinion">별도 의견 없음</li>'
    suffix = _e(slide.payload.get("page_suffix", ""))
    header = _content_header(f"5) 주관식 설문 결과{suffix}", "02", "만족도 통계")
    return _base(
        f'{header}<div class="subjective-question">■ {_e(q.section_label)} &nbsp; {index + 1}. {_e(strip_leading_question_numbers(q.question))}</div>'
        f'<ul class="opinion-list">{items}</ul>{_content_footer()}'
    )


def _insights(report: ReportData) -> str:
    items = "".join(f'<li><span>{idx + 1}</span><p>{_e(text)}</p></li>' for idx, text in enumerate(report.insights))
    if not items:
        items = '<li><span>1</span><p>자동 분석된 시사점이 없습니다. 입력부에서 내용을 추가해 주세요.</p></li>'
    return _base(f'{_content_header("6) 종합 시사점", "02", "만족도 통계")}<ul class="insight-list">{items}</ul>{_content_footer()}')


def _photos(slide: SlideItem, photo_names: List[str]) -> str:
    start = int(slide.payload.get("page_index", 0)) * 6; names = photo_names[start : start + 6]
    if names:
        cells = "".join(f'<div class="photo-cell">{_e(name)}</div>' for name in names)
        cells += "".join('<div class="photo-cell empty"></div>' for _ in range(6 - len(names)))
        body = f'<div class="photo-grid">{cells}</div>'
    else:
        body = '<div class="photo-empty-canvas"></div>'
    suffix = _e(slide.payload.get("page_suffix", ""))
    return _base(f'{_content_header(f"1) 현장 사진{suffix}", "03", "현장 사진")}{body}{_content_footer()}')


def _thanks() -> str:
    return _section_background_content('<div class="thanks-main">THANK YOU</div><div class="thanks-sub">감사합니다</div>')


def render_slide_html(report: ReportData, slide: SlideItem, photo_names: List[str] | None = None) -> str:
    photo_names = photo_names or []
    renderers = {
        "cover": lambda: _cover(report), "toc": _toc, "section": lambda: _section(slide),
        "overview": lambda: _overview(report), "curriculum": lambda: _curriculum(report),
        "survey_structure": lambda: _survey_structure(report, slide), "summary": lambda: _summary(report, slide),
        "objective": lambda: _objective(report, slide), "lecturer_comparison": lambda: _lecturer_comparison(report, slide),
        "choice": lambda: _choice(report, slide), "subjective": lambda: _subjective(report, slide),
        "insights": lambda: _insights(report), "photos": lambda: _photos(slide, photo_names), "thanks": _thanks,
    }
    return renderers[slide.kind]()


PREVIEW_CSS = f"""
<style>
{FONT_FACE_CSS}
:root{{--orange:{ORANGE};--orange-dark:{ORANGE_DARK};--dark:{DARK};--muted:{MUTED};--grid:{LIGHT_GRID};}}
.ppt-stage{{width:100%;display:flex;justify-content:center;padding:10px;background:#e8e8e8;border-radius:12px;box-sizing:border-box;}}
.ppt-slide{{position:relative;width:100%;aspect-ratio:16/9;background:#fff;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.16);font-family:'ReportAssetFont','Malgun Gothic','Noto Sans KR',sans-serif;color:var(--dark);box-sizing:border-box;}}
.section-background{{background-image:url('{SECTION_BG_URI}');background-size:100% 100%;background-repeat:no-repeat;}}
.section-right-title{{position:absolute;left:51.5%;top:35%;width:39%;color:var(--orange);font-weight:800;line-height:1.18;}}
.cover-course{{font-size:clamp(21px,3vw,48px);word-break:keep-all;}}.cover-report{{font-size:clamp(24px,3.5vw,56px);}}
.section-company{{position:absolute;left:51.7%;top:28%;font-size:clamp(11px,1.15vw,19px);color:#777;}}.section-schedule{{position:absolute;left:51.7%;top:60%;font-size:clamp(11px,1.15vw,19px);color:#777;}}
.section-heading{{font-size:clamp(25px,3.6vw,58px);}}.section-heading small{{display:block;font-size:.48em;margin-bottom:.2em;}}
.section-item-line{{position:absolute;left:51.7%;top:57%;width:39%;font-size:clamp(11px,1.05vw,18px);color:#777;line-height:1.6;}}
.toc-right-list{{position:absolute;left:51.7%;top:44%;width:42%;display:flex;flex-direction:column;gap:1.35vh;}}.toc-right-list div{{display:grid;grid-template-columns:12% 29% 59%;align-items:start;}}
.toc-right-list span{{font-weight:800;color:var(--orange);font-size:clamp(11px,1.2vw,20px);}}.toc-right-list b{{font-size:clamp(10px,1.15vw,19px);}}.toc-right-list em{{font-style:normal;color:#777;font-size:clamp(8px,.78vw,13px);line-height:1.5;}}
.thanks-main{{font-size:clamp(27px,4vw,64px);}}.thanks-sub{{font-size:clamp(13px,1.5vw,24px);margin-top:.35em;}}
.content-main-title{{position:absolute;left:2.4%;top:4.1%;z-index:2;background:#fff;padding-right:1%;font-size:clamp(20px,3vw,48px);font-weight:800;color:var(--orange);letter-spacing:-.04em;}}
.content-rule{{position:absolute;left:24%;right:3.4%;top:9.7%;height:.55%;background:var(--orange);}}.content-section-tag{{position:absolute;right:2.7%;top:3.8%;display:flex;align-items:center;gap:.8em;background:#fff;padding-left:.8em;color:var(--orange);font-size:clamp(13px,1.55vw,25px);font-weight:800;white-space:nowrap;}}
.content-section-tag span{{display:inline-flex;align-items:center;justify-content:center;background:var(--orange);color:#fff;padding:.18em .6em;font-weight:500;}}
.content-logo{{position:absolute;right:3.2%;bottom:3.2%;width:11.2%;height:auto;object-fit:contain;}}.content-logo-text{{position:absolute;right:3.6%;bottom:3.4%;font-size:clamp(9px,1vw,16px);font-weight:800;color:#111;}}
.overview-grid{{position:absolute;left:9%;top:21%;width:80%;height:61%;display:grid;grid-template-columns:23% 77%;grid-template-rows:repeat(4,1fr) 1.65fr;border-top:1px solid #cfcfcf;border-left:1px solid #cfcfcf;}}
.overview-label,.overview-value{{display:flex;align-items:center;padding:0 4%;border-right:1px solid #cfcfcf;border-bottom:1px solid #cfcfcf;font-size:clamp(12px,1.25vw,21px);}}.overview-label{{font-weight:800;background:var(--orange);color:#fff;justify-content:center;font-size:clamp(13px,1.35vw,22px);}}.overview-value{{font-weight:600;}}
.curriculum-table,.survey-table,.comparison-table{{position:absolute;left:4.8%;top:19%;width:90.4%;border-collapse:collapse;table-layout:fixed;font-size:clamp(9px,.98vw,16.5px);}}
.curriculum-table th,.survey-table th,.comparison-table th{{background:var(--orange);color:white;padding:.8%;border:1px solid #fff;}}.curriculum-table td,.survey-table td,.comparison-table td{{padding:.72%;border:1px solid #d4d4d4;vertical-align:middle;line-height:1.3;}}
.curriculum-table th:nth-child(1){{width:12%;}}.curriculum-table th:nth-child(2){{width:19%;}}.curriculum-table th:nth-child(3){{width:51%;}}.curriculum-table th:nth-child(4){{width:18%;}}.survey-table th:first-child{{width:23%;}}.survey-table td:first-child{{font-weight:700;background:#f7f7f7;}}.empty-cell{{height:4em;}}
.summary-note{{position:absolute;left:7%;top:14%;font-size:clamp(13px,1.32vw,22px);line-height:1.8;}}.summary-note b{{display:block;}}.summary-note span{{display:block;margin-left:1.2em;color:#222;}}
.summary-chart-wrap{{position:absolute;left:7%;top:24%;width:88%;height:58%;display:grid;grid-template-columns:5% 95%;}}.summary-y-labels{{display:flex;flex-direction:column;justify-content:space-between;text-align:right;padding-right:14%;font-size:clamp(9px,1vw,17px);color:#555;}}.summary-plot{{position:relative;border-left:1px solid #2a94bf;border-bottom:1px solid #ddd;background:repeating-linear-gradient(to bottom,var(--grid) 0,var(--grid) 1px,transparent 1px,transparent 16.666%);}}
.summary-bars{{position:absolute;left:2%;right:1%;top:0;bottom:0;display:flex;align-items:flex-end;justify-content:space-around;gap:1.2%;}}.summary-bar-column{{position:relative;height:100%;flex:1;display:flex;align-items:center;justify-content:flex-end;flex-direction:column;min-width:0;}}.summary-bar{{width:28%;min-width:16px;background:var(--orange);}}.summary-value{{font-size:clamp(11px,1.15vw,19px);margin-bottom:.4em;color:#111;}}.summary-x-label{{position:absolute;top:103%;width:135%;text-align:center;font-size:clamp(8px,.9vw,15px);line-height:1.25;color:#555;word-break:keep-all;}}
.summary-dist-question{{position:absolute;left:7%;top:22.2%;width:72%;font-size:clamp(10px,1.02vw,17px);font-weight:800;line-height:1.35;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}.summary-dist-average{{position:absolute;right:6%;top:21.3%;font-size:clamp(11px,1.12vw,18px);color:#555;}}.summary-dist-average b{{font-size:1.45em;color:var(--orange);}}
.summary-dist-chart{{position:absolute;left:7%;top:31%;width:88%;height:48%;display:grid;grid-template-columns:5% 95%;}}.summary-dist-y-labels{{display:flex;flex-direction:column;justify-content:space-between;text-align:right;padding-right:14%;font-size:clamp(8px,.88vw,14px);color:#555;}}.summary-dist-plot{{position:relative;border-left:1px solid #999;border-bottom:1px solid #999;background:repeating-linear-gradient(to bottom,var(--grid) 0,var(--grid) 1px,transparent 1px,transparent 25%);}}.dist-columns{{position:absolute;left:1%;right:1%;top:0;bottom:0;display:flex;align-items:stretch;gap:.7%;}}.dist-col{{position:relative;flex:1;height:100%;min-width:0;}}.dist-bar-area{{position:absolute;inset:0;display:flex;align-items:flex-end;justify-content:center;}}.dist-bar{{width:58%;min-width:7px;background:var(--orange);}}.dist-value{{position:absolute;left:50%;transform:translateX(-50%);z-index:2;text-align:center;white-space:nowrap;font-size:clamp(7px,.7vw,12px);line-height:1.15;color:#333;}}.dist-value b{{display:block;font-weight:800;}}.dist-value small{{display:block;font-size:.86em;}}.dist-score{{position:absolute;top:102%;left:0;right:0;text-align:center;font-size:clamp(8px,.82vw,14px);color:#333;}}.summary-dist-x-title{{position:absolute;left:53%;top:108%;font-size:clamp(8px,.8vw,13px);color:#777;}}
.question-title,.choice-question{{position:absolute;left:5.2%;top:15.2%;width:82%;max-height:12%;overflow:hidden;font-size:clamp(11px,1.15vw,19px);font-weight:800;line-height:1.32;word-break:keep-all;}}.scale-average{{position:absolute;right:5%;top:15%;font-size:clamp(10px,1.05vw,17px);color:#555;}}.scale-average b{{font-size:1.35em;color:var(--orange);}}
.objective-chart{{position:absolute;left:10.4%;top:25%;width:84%;height:57%;display:flex;flex-direction:column;justify-content:space-between;}}.objective-row{{display:grid;grid-template-columns:0 100%;grid-template-rows:auto 1fr;position:relative;flex:1;min-height:0;}}.objective-label{{grid-column:2;align-self:end;margin-bottom:.35em;font-size:clamp(11px,1.15vw,19px);}}.objective-track{{grid-column:2;position:relative;height:36%;min-height:14px;background:repeating-linear-gradient(to right,#e8e8e8 0,#e8e8e8 1px,transparent 1px,transparent calc(100% / 9));}}.objective-fill{{height:100%;display:flex;align-items:center;justify-content:flex-end;padding-right:.7em;min-width:0;overflow:hidden;}}.objective-count{{color:#fff;font-weight:800;font-size:clamp(10px,1.02vw,17px);}}.objective-ticks{{display:flex;justify-content:space-between;font-size:clamp(10px,.92vw,15px);color:#555;margin-top:.3em;}}
.histogram{{position:absolute;left:9%;top:26%;width:86%;height:55%;border-left:1px solid #999;border-bottom:1px solid #999;}}.hist-columns{{position:absolute;inset:3% 2% 0 4%;display:flex;align-items:flex-end;gap:2%;}}.hist-col{{height:100%;flex:1;display:grid;grid-template-rows:9% 80% 11%;align-items:end;text-align:center;min-width:0;}}.hist-count{{font-size:clamp(8px,.83vw,14px);color:#444;}}.hist-bar-wrap{{height:100%;display:flex;align-items:flex-end;justify-content:center;background:repeating-linear-gradient(to top,#eee 0,#eee 1px,transparent 1px,transparent 20%);}}.hist-bar{{width:58%;min-width:8px;background:var(--orange);}}.hist-label{{font-size:clamp(9px,.92vw,15px);padding-top:.25em;}}.hist-y-title{{position:absolute;left:-5%;top:42%;transform:rotate(-90deg);font-size:clamp(8px,.8vw,13px);color:#777;}}.hist-x-title{{position:absolute;bottom:-8%;left:47%;font-size:clamp(8px,.8vw,13px);color:#777;}}
.comparison-note{{position:absolute;left:6%;top:15%;font-size:clamp(12px,1.2vw,20px);font-weight:800;}}.comparison-table{{top:23%;left:7%;width:86%;}}.comparison-table td{{text-align:center;}}.comparison-table td:first-child{{font-weight:800;background:#f7f7f7;}}.comparison-table .lecturer-total{{font-weight:800;color:var(--orange);background:#fff7f2;}}
.choice-pie{{position:absolute;left:12%;top:27%;width:28%;aspect-ratio:1;border-radius:50%;box-shadow:inset 0 0 0 1px rgba(0,0,0,.08);}}.choice-legend{{position:absolute;left:49%;top:25%;width:43%;height:55%;display:flex;flex-direction:column;justify-content:center;gap:4%;}}.choice-legend div{{display:grid;grid-template-columns:5% 55% 40%;align-items:center;font-size:clamp(9px,1vw,17px);}}.choice-legend i{{width:.8em;height:.8em;border-radius:2px;}}.choice-legend b{{text-align:right;color:#333;}}
.multi-chart{{position:absolute;left:8%;top:24%;width:86%;height:58%;display:flex;flex-direction:column;justify-content:center;gap:4%;}}.multi-row{{display:grid;grid-template-columns:25% 55% 20%;align-items:center;gap:1%;font-size:clamp(9px,1vw,17px);}}.multi-label{{text-align:right;padding-right:4%;}}.multi-track{{height:1.45em;background:#eee;}}.multi-fill{{height:100%;background:var(--orange);}}.multi-value{{font-weight:800;padding-left:4%;}}.multi-foot{{font-size:clamp(8px,.8vw,13px);color:#777;text-align:right;margin-top:1%;}}
.subjective-question{{position:absolute;left:5.2%;top:16%;width:89%;font-size:clamp(12px,1.28vw,21px);font-weight:800;line-height:1.4;}}.opinion-list{{position:absolute;left:7.2%;top:24%;width:87%;height:60%;margin:0;padding-left:1.6em;display:flex;flex-direction:column;gap:1.4%;font-size:clamp(11px,1.1vw,18px);line-height:1.48;}}.opinion-list li{{padding-left:.35em;}}.opinion-list li::marker{{content:'-  ';color:#444;}}.empty-opinion{{color:#999;}}
.insight-list{{position:absolute;left:8%;top:22%;width:84%;height:60%;list-style:none;padding:0;margin:0;display:flex;flex-direction:column;justify-content:center;gap:6%;}}.insight-list li{{display:grid;grid-template-columns:7% 93%;align-items:center;background:#fafafa;border-left:.55em solid var(--orange);padding:2.2% 3%;box-shadow:0 2px 8px rgba(0,0,0,.06);}}.insight-list span{{font-size:clamp(17px,2vw,32px);font-weight:800;color:var(--orange);}}.insight-list p{{margin:0;font-size:clamp(11px,1.15vw,19px);line-height:1.45;}}
.photo-grid{{position:absolute;left:4.7%;top:17%;width:90.6%;height:68%;display:grid;grid-template-columns:repeat(3,1fr);grid-template-rows:repeat(2,1fr);gap:1.4%;}}.photo-cell{{display:flex;align-items:center;justify-content:center;background:#f1f1f1;border:1px solid #ddd;color:#777;font-size:clamp(8px,.8vw,14px);padding:4%;text-align:center;overflow:hidden;}}.photo-cell.empty{{background:#fff;border-color:#eee;color:transparent;}}.photo-empty-canvas{{position:absolute;left:4.7%;top:17%;width:90.6%;height:68%;background:#fff;}}
</style>
"""
