from __future__ import annotations

import base64
import html
from pathlib import Path
from typing import List, Sequence

from .models import ObjectiveQuestion, ReportData
from .slide_plan import SlideItem
from .text_utils import strip_leading_question_numbers


ORANGE = "#ED702C"
ORANGE_DARK = "#D85C1E"
DARK = "#444444"
MUTED = "#6F6F6F"
PURPLE = "#A02B98"
BLUE = "#1B9FD0"
LIGHT_GRID = "#E5E5E5"

ASSET_DIR = Path(__file__).resolve().parent.parent / "assets"


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
    class_attr = f"ppt-slide {class_name}".strip()
    return f'<div class="ppt-stage"><div class="{class_attr}">{content}</div></div>'


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
    return _base(
        f"""
        <div class="section-right-title">{title_html}</div>
        {extra_html}
        """,
        "section-background",
    )


def _cover(report: ReportData) -> str:
    course = report.course_name or "[과정 이름]"
    company = f'<div class="section-company">{_e(report.company_name)}</div>' if report.company_name else ""
    schedule = f'<div class="section-schedule">{_e(report.schedule)}</div>' if report.schedule else ""
    return _section_background_content(
        f'<div class="cover-course">[{_e(course)}]</div><div class="cover-report">결과보고서</div>',
        company + schedule,
    )


def _toc() -> str:
    return _section_background_content(
        '<div class="section-heading">목차</div>',
        """
        <div class="toc-right-list">
          <div><span>01</span><b>교육 개요</b><em>교육 개요 · 커리큘럼</em></div>
          <div><span>02</span><b>만족도 통계</b><em>설문 구성 · 객관식 설문 결과 · 주관식 설문 결과</em></div>
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
        ("과정명", report.course_name),
        ("교육일정", report.schedule),
        ("교육방식", report.delivery_method),
        ("교육 대상", target),
        ("교육 목표", report.objective),
    ]
    row_html = "".join(
        f'<div class="overview-label">{_e(k)}</div><div class="overview-value">{_e(v)}</div>'
        for k, v in rows
    )
    return _base(
        f"""
        {_content_header('1) 교육 개요', '01', '교육 개요')}
        <div class="overview-grid">{row_html}</div>
        {_content_footer()}
        """
    )


def _curriculum(report: ReportData) -> str:
    rows = report.curriculum or []
    body = "".join(
        f"<tr><td>{_e(r.day)}</td><td>{_e(r.time)}</td><td>{_e(r.content)}</td><td>{_e(r.instructor)}</td></tr>"
        for r in rows[:8]
    )
    if not body:
        body = '<tr><td colspan="4" class="empty-cell"></td></tr>'
    return _base(
        f"""
        {_content_header('2) 커리큘럼', '01', '교육 개요')}
        <table class="curriculum-table"><thead><tr><th>Day</th><th>시간</th><th>교육 내용</th><th>강사/비고</th></tr></thead><tbody>{body}</tbody></table>
        {_content_footer()}
        """
    )


def _survey_structure(report: ReportData, slide: SlideItem) -> str:
    questions = list(report.objective_questions) + list(report.subjective_questions)
    start = int(slide.payload.get("start", 0))
    end = int(slide.payload.get("end", min(start + 10, len(questions))))
    chunk = questions[start:end]
    rows = "".join(
        f"<tr><td>{_e(q.section_label)}</td><td>{start + idx}. {_e(strip_leading_question_numbers(q.question))}</td></tr>"
        for idx, q in enumerate(chunk, start=1)
    )
    if not rows:
        rows = '<tr><td colspan="2" class="empty-cell"></td></tr>'
    suffix = f" ({slide.payload.get('page_index', 0) + 1})" if len(questions) > 10 else ""
    return _base(
        f"""
        {_content_header(f'1) 설문 구성{suffix}', '02', '만족도 통계')}
        <table class="survey-table"><thead><tr><th>항목</th><th>문항</th></tr></thead><tbody>{rows}</tbody></table>
        {_content_footer()}
        """
    )


def _summary_label(q: ObjectiveQuestion) -> str:
    if q.instructor_name:
        return f"{q.instructor_name}\n만족도"
    if q.course_module and q.section_label in {"과정 만족도", "강사 만족도"}:
        return q.course_module
    return q.section_label


def _summary(report: ReportData, slide: SlideItem) -> str:
    start = int(slide.payload.get("start", 0))
    end = int(slide.payload.get("end", min(start + 7, len(report.objective_questions))))
    qs = report.objective_questions[start:end]
    if report.total_participants:
        count_line = (
            f"수강인원 총 {report.total_participants}명 중 {report.response_count}명 응답"
            f" ｜ 응답률 {report.response_rate}%"
        )
    else:
        count_line = f"총 {report.response_count}명 응답"

    bars = []
    for q in qs:
        height = max(0.0, min(100.0, (q.average / 6.0) * 100.0))
        label = _summary_label(q).replace("\n", "<br/>")
        bars.append(
            f"""
            <div class="summary-bar-column">
              <div class="summary-value">{q.average:.1f}</div>
              <div class="summary-bar" style="height:{height:.1f}%"></div>
              <div class="summary-x-label">{label}</div>
            </div>
            """
        )
    page_suffix = f" ({slide.payload.get('page_index', 0) + 1})" if len(report.objective_questions) > 7 else ""
    return _base(
        f"""
        {_content_header(f'2) 객관식 설문 결과{page_suffix}', '02', '만족도 통계')}
        <div class="summary-note"><b>■ 만족도 요약</b><span>- {_e(count_line)}</span></div>
        <div class="summary-chart-wrap">
          <div class="summary-y-labels"><span>6.0</span><span>5.0</span><span>4.0</span><span>3.0</span><span>2.0</span><span>1.0</span><span>0.0</span></div>
          <div class="summary-plot"><div class="summary-bars">{''.join(bars)}</div></div>
        </div>
        {_content_footer()}
        """
    )


def _objective(report: ReportData, slide: SlideItem) -> str:
    q_index = int(slide.payload["question_index"])
    q = report.objective_questions[q_index]
    max_count = max(report.response_count, q.valid_responses, max(q.counts) if q.counts else 0, 1)
    rows = []
    colors = [PURPLE, BLUE, ORANGE, "#F39A66", "#C9C9C9", "#B4B4B4", "#9B9B9B"]
    for idx, (label, count) in enumerate(zip(q.scale_labels, q.counts)):
        width = max(0.0, min(100.0, (count / max_count) * 100.0))
        value = f'<span class="objective-count">{count}</span>' if count > 0 else ""
        rows.append(
            f"""
            <div class="objective-row">
              <div class="objective-label">{_e(label)}</div>
              <div class="objective-track"><div class="objective-fill" style="width:{width:.1f}%;background:{colors[idx % len(colors)]}">{value}</div></div>
            </div>
            """
        )
    ticks = "".join(f"<span>{i}</span>" for i in range(max_count + 1))
    return _base(
        f"""
        {_content_header('2) 객관식 설문 결과', '02', '만족도 통계')}
        <div class="question-title">■ {_e(q.section_label)} &nbsp; {q_index + 1}. {_e(strip_leading_question_numbers(q.question))}</div>
        <div class="objective-chart">{''.join(rows)}<div class="objective-ticks">{ticks}</div></div>
        {_content_footer()}
        """
    )


def _subjective(report: ReportData, slide: SlideItem) -> str:
    q_index = int(slide.payload["question_index"])
    q = report.subjective_questions[q_index]
    question_no = len(report.objective_questions) + q_index + 1
    answers: Sequence[str] = slide.payload.get("answers", [])
    items = "".join(f'<li>{_e(a)}</li>' for a in answers)
    if not items:
        items = '<li class="empty-opinion">별도 의견 없음</li>'
    suffix = slide.payload.get("page_suffix", "")
    return _base(
        f"""
        {_content_header(f'3) 주관식 설문 결과{_e(suffix)}', '02', '만족도 통계')}
        <div class="subjective-question">■ {_e(q.section_label)} &nbsp; {question_no}. {_e(strip_leading_question_numbers(q.question))}</div>
        <ul class="opinion-list">{items}</ul>
        {_content_footer()}
        """
    )


def _photos(slide: SlideItem, photo_names: List[str]) -> str:
    start = int(slide.payload.get("page_index", 0)) * 6
    names = photo_names[start : start + 6]
    cells = "".join(f'<div class="photo-cell">{_e(name)}</div>' for name in names)
    cells += "".join('<div class="photo-cell empty">사진</div>' for _ in range(6 - len(names)))
    return _base(
        f"""
        {_content_header(f"3) 현장 사진{_e(slide.payload.get('page_suffix', ''))}", '03', '현장 사진')}
        <div class="photo-grid">{cells}</div>
        {_content_footer()}
        """
    )


def _thanks() -> str:
    return _section_background_content(
        '<div class="thanks-main">THANK YOU</div><div class="thanks-sub">감사합니다</div>'
    )


def render_slide_html(report: ReportData, slide: SlideItem, photo_names: List[str] | None = None) -> str:
    photo_names = photo_names or []
    renderers = {
        "cover": lambda: _cover(report),
        "toc": _toc,
        "section": lambda: _section(slide),
        "overview": lambda: _overview(report),
        "curriculum": lambda: _curriculum(report),
        "survey_structure": lambda: _survey_structure(report, slide),
        "summary": lambda: _summary(report, slide),
        "objective": lambda: _objective(report, slide),
        "subjective": lambda: _subjective(report, slide),
        "photos": lambda: _photos(slide, photo_names),
        "thanks": _thanks,
    }
    return renderers[slide.kind]()


PREVIEW_CSS = f"""
<style>
:root {{--orange:{ORANGE};--orange-dark:{ORANGE_DARK};--dark:{DARK};--muted:{MUTED};--grid:{LIGHT_GRID};}}
.ppt-stage {{width:100%;display:flex;justify-content:center;padding:10px;background:#e8e8e8;border-radius:12px;box-sizing:border-box;}}
.ppt-slide {{position:relative;width:100%;aspect-ratio:16/9;background:#fff;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.16);font-family:'Malgun Gothic','Noto Sans KR',sans-serif;color:var(--dark);box-sizing:border-box;}}
.section-background {{background-image:url('{SECTION_BG_URI}');background-size:100% 100%;background-repeat:no-repeat;}}
.section-right-title {{position:absolute;left:51.5%;top:35%;width:39%;color:var(--orange);font-weight:800;line-height:1.18;}}
.cover-course {{font-size:clamp(21px,3vw,48px);word-break:keep-all;}}
.cover-report {{font-size:clamp(24px,3.5vw,56px);}}
.section-company {{position:absolute;left:51.7%;top:28%;font-size:clamp(9px,1vw,17px);color:#777;}}
.section-schedule {{position:absolute;left:51.7%;top:60%;font-size:clamp(9px,1vw,17px);color:#777;}}
.section-heading {{font-size:clamp(25px,3.6vw,58px);}}
.section-heading small {{display:block;font-size:.48em;margin-bottom:.2em;}}
.section-item-line {{position:absolute;left:51.7%;top:57%;width:39%;font-size:clamp(9px,1.05vw,17px);color:#777;line-height:1.6;}}
.toc-right-list {{position:absolute;left:51.7%;top:44%;width:40%;display:flex;flex-direction:column;gap:1.4vh;}}
.toc-right-list div {{display:grid;grid-template-columns:12% 31% 57%;align-items:start;}}
.toc-right-list span {{font-weight:800;color:var(--orange);font-size:clamp(11px,1.2vw,20px);}}
.toc-right-list b {{font-size:clamp(10px,1.15vw,19px);}}
.toc-right-list em {{font-style:normal;color:#777;font-size:clamp(8px,.85vw,14px);line-height:1.5;}}
.thanks-main {{font-size:clamp(27px,4vw,64px);}}
.thanks-sub {{font-size:clamp(13px,1.5vw,24px);margin-top:.35em;}}
.content-main-title {{position:absolute;left:2.4%;top:4.1%;z-index:2;background:#fff;padding-right:1%;font-size:clamp(20px,3vw,48px);font-weight:800;color:var(--orange);letter-spacing:-.04em;}}
.content-rule {{position:absolute;left:24%;right:3.4%;top:9.7%;height:.55%;background:var(--orange);}}
.content-section-tag {{position:absolute;right:2.7%;top:3.8%;display:flex;align-items:center;gap:.8em;background:#fff;padding-left:.8em;color:var(--orange);font-size:clamp(12px,1.6vw,26px);font-weight:800;}}
.content-section-tag span {{display:inline-flex;align-items:center;justify-content:center;background:var(--orange);color:#fff;padding:.18em .6em;font-weight:500;}}
.content-logo {{position:absolute;right:3.6%;bottom:3.4%;width:10.8%;height:auto;object-fit:contain;}}
.content-logo-text {{position:absolute;right:3.6%;bottom:3.4%;font-size:clamp(9px,1vw,16px);font-weight:800;color:#111;}}
.overview-grid {{position:absolute;left:9%;top:21%;width:80%;height:61%;display:grid;grid-template-columns:23% 77%;grid-template-rows:repeat(4,1fr) 1.65fr;border-top:1px solid #cfcfcf;border-left:1px solid #cfcfcf;}}
.overview-label,.overview-value {{display:flex;align-items:center;padding:0 4%;border-right:1px solid #cfcfcf;border-bottom:1px solid #cfcfcf;font-size:clamp(9px,1.05vw,17px);}}
.overview-label {{font-weight:700;background:#f5f5f5;}}
.curriculum-table,.survey-table {{position:absolute;left:4.8%;top:19%;width:90.4%;border-collapse:collapse;table-layout:fixed;font-size:clamp(7px,.78vw,13px);}}
.curriculum-table th,.survey-table th {{background:var(--orange);color:white;padding:.8%;border:1px solid #fff;}}
.curriculum-table td,.survey-table td {{padding:.75%;border:1px solid #d4d4d4;vertical-align:middle;line-height:1.35;}}
.curriculum-table th:nth-child(1){{width:12%;}}.curriculum-table th:nth-child(2){{width:19%;}}.curriculum-table th:nth-child(3){{width:51%;}}.curriculum-table th:nth-child(4){{width:18%;}}
.survey-table th:first-child{{width:23%;}}.survey-table td:first-child{{font-weight:700;background:#f7f7f7;}}
.empty-cell{{height:4em;}}
.summary-note {{position:absolute;left:7%;top:14%;font-size:clamp(10px,1.1vw,18px);line-height:1.8;}}
.summary-note b{{display:block;}}.summary-note span{{display:block;margin-left:1.2em;color:#222;}}
.summary-chart-wrap {{position:absolute;left:7%;top:24%;width:88%;height:58%;display:grid;grid-template-columns:5% 95%;}}
.summary-y-labels {{display:flex;flex-direction:column;justify-content:space-between;text-align:right;padding-right:14%;font-size:clamp(8px,.9vw,15px);color:#555;}}
.summary-plot {{position:relative;border-left:1px solid #2a94bf;border-bottom:1px solid #ddd;background:repeating-linear-gradient(to bottom,var(--grid) 0,var(--grid) 1px,transparent 1px,transparent 16.666%);}}
.summary-bars {{position:absolute;left:2%;right:1%;top:0;bottom:0;display:flex;align-items:flex-end;justify-content:space-around;gap:1.2%;}}
.summary-bar-column {{position:relative;height:100%;flex:1;display:flex;align-items:center;justify-content:flex-end;flex-direction:column;min-width:0;}}
.summary-bar {{width:28%;min-width:18px;background:var(--orange);}}
.summary-value {{font-size:clamp(9px,1vw,17px);margin-bottom:.4em;color:#111;}}
.summary-x-label {{position:absolute;top:103%;width:130%;text-align:center;font-size:clamp(7px,.78vw,13px);line-height:1.25;color:#555;word-break:keep-all;}}
.question-title,.subjective-question {{position:absolute;left:5.2%;top:16%;width:89%;font-size:clamp(11px,1.35vw,22px);font-weight:800;line-height:1.45;}}
.objective-chart {{position:absolute;left:10.4%;top:25%;width:84%;height:57%;display:flex;flex-direction:column;justify-content:space-between;}}
.objective-row {{display:grid;grid-template-columns:0 100%;grid-template-rows:auto 1fr;position:relative;flex:1;min-height:0;}}
.objective-label {{grid-column:2;align-self:end;margin-bottom:.35em;font-size:clamp(9px,1vw,17px);}}
.objective-track {{grid-column:2;position:relative;height:36%;min-height:14px;background:repeating-linear-gradient(to right,#e8e8e8 0,#e8e8e8 1px,transparent 1px,transparent calc(100% / 9));}}
.objective-fill {{height:100%;display:flex;align-items:center;justify-content:flex-end;padding-right:.7em;min-width:0;}}
.objective-count {{color:#fff;font-weight:800;font-size:clamp(8px,.9vw,15px);}}
.objective-ticks {{display:flex;justify-content:space-between;font-size:clamp(7px,.75vw,12px);color:#555;margin-top:.3em;}}
.subjective-question {{top:16%;}}
.opinion-list {{position:absolute;left:7.2%;top:24%;width:87%;height:60%;margin:0;padding-left:1.6em;display:flex;flex-direction:column;gap:1.4%;font-size:clamp(8px,.9vw,15px);line-height:1.48;}}
.opinion-list li {{padding-left:.35em;}}
.opinion-list li::marker {{content:'-  ';color:#444;}}
.empty-opinion{{color:#999;}}
.photo-grid {{position:absolute;left:4.7%;top:17%;width:90.6%;height:68%;display:grid;grid-template-columns:repeat(3,1fr);grid-template-rows:repeat(2,1fr);gap:1.4%;}}
.photo-cell {{display:flex;align-items:center;justify-content:center;background:#f1f1f1;border:1px solid #ddd;color:#777;font-size:clamp(8px,.8vw,14px);padding:4%;text-align:center;overflow:hidden;}}
.photo-cell.empty{{color:#aaa;}}
</style>
"""
