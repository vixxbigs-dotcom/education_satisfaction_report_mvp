from __future__ import annotations

import io
import math
from pathlib import Path
from typing import List, Sequence, Tuple

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .models import ObjectiveQuestion, ReportData
from .slide_plan import SlideItem, build_slide_plan
from .text_utils import strip_leading_question_numbers


SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
ORANGE = RGBColor(237, 112, 44)
ORANGE_DARK = RGBColor(216, 92, 30)
DARK = RGBColor(68, 68, 68)
MUTED = RGBColor(111, 111, 111)
LIGHT = RGBColor(247, 247, 247)
LINE = RGBColor(215, 215, 215)
GRID = RGBColor(229, 229, 229)
WHITE = RGBColor(255, 255, 255)
BLACK = RGBColor(20, 20, 20)
PURPLE = RGBColor(160, 43, 152)
BLUE = RGBColor(27, 159, 208)
FONT = "맑은 고딕"

ASSET_DIR = Path(__file__).resolve().parent.parent / "assets"
SECTION_BACKGROUND = ASSET_DIR / "multicampus_section_background.png"
LOGO_PATH = ASSET_DIR / "multicampus_logo.png"
SUBJECTIVE_PER_SLIDE = 5


def _set_bg(slide, color=WHITE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=16,
    bold=False,
    color=DARK,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.MIDDLE,
    margin=0.03,
    font_name=FONT,
):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text or "")
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _add_full_background(slide) -> None:
    if SECTION_BACKGROUND.exists():
        slide.shapes.add_picture(str(SECTION_BACKGROUND), 0, 0, SLIDE_W, SLIDE_H)
    else:
        _set_bg(slide, RGBColor(242, 242, 242))
        panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.25), 0, Inches(7.083), SLIDE_H)
        panel.fill.solid()
        panel.fill.fore_color.rgb = WHITE
        panel.line.fill.background()


def _add_logo(slide) -> None:
    if LOGO_PATH.exists():
        slide.shapes.add_picture(str(LOGO_PATH), Inches(11.45), Inches(6.98), Inches(1.55), Inches(0.247))
    else:
        _add_text(slide, "multicampus", Inches(11.25), Inches(6.95), Inches(1.6), Inches(0.3), 10, True, BLACK, PP_ALIGN.RIGHT)


def _add_content_header(slide, main_title: str, section_no: str, section_title: str) -> None:
    _add_text(slide, main_title, Inches(0.32), Inches(0.27), Inches(5.0), Inches(0.56), 31, True, ORANGE)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.6), Inches(0.73), Inches(7.75), Inches(0.035))
    line.fill.solid()
    line.fill.fore_color.rgb = ORANGE
    line.line.fill.background()

    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(11.4), Inches(0.31), Inches(0.54), Inches(0.38))
    box.fill.solid()
    box.fill.fore_color.rgb = ORANGE
    box.line.fill.background()
    _add_text(slide, section_no, Inches(11.41), Inches(0.32), Inches(0.52), Inches(0.35), 16.5, False, WHITE, PP_ALIGN.CENTER)
    _add_text(slide, section_title, Inches(11.98), Inches(0.30), Inches(1.34), Inches(0.40), 15.5, True, ORANGE, PP_ALIGN.LEFT)
    _add_logo(slide)


def _add_cover(prs: Presentation, report: ReportData) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_full_background(slide)
    if report.company_name:
        _add_text(slide, report.company_name, Inches(6.9), Inches(2.18), Inches(5.25), Inches(0.36), 12, False, MUTED)
    _add_text(slide, report.course_name or '과정 이름', Inches(6.88), Inches(3.08), Inches(5.4), Inches(0.55), 29, True, ORANGE, valign=MSO_ANCHOR.BOTTOM)
    _add_text(slide, "결과보고서", Inches(6.88), Inches(3.62), Inches(5.4), Inches(0.58), 31, True, ORANGE, valign=MSO_ANCHOR.TOP)
    if report.schedule:
        _add_text(slide, report.schedule, Inches(6.9), Inches(4.55), Inches(4.8), Inches(0.32), 11, False, MUTED)


def _add_toc(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_full_background(slide)
    _add_text(slide, "목차", Inches(6.88), Inches(2.48), Inches(4.8), Inches(0.55), 31, True, ORANGE)
    entries = [
        ("01", "교육 개요", "교육 개요 · 커리큘럼"),
        ("02", "만족도 통계", "설문 구성 · 객관식 설문 결과 · 주관식 설문 결과"),
        ("03", "현장 사진", "교육 현장 스케치"),
    ]
    for idx, (no, title, detail) in enumerate(entries):
        y = 3.32 + idx * 0.72
        _add_text(slide, no, Inches(6.9), Inches(y), Inches(0.55), Inches(0.28), 14, True, ORANGE)
        _add_text(slide, title, Inches(7.45), Inches(y), Inches(1.65), Inches(0.28), 13, True, DARK)
        _add_text(slide, detail, Inches(9.05), Inches(y), Inches(3.65), Inches(0.38), 9, False, MUTED, valign=MSO_ANCHOR.TOP)


def _add_section(prs: Presentation, slide_item: SlideItem) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_full_background(slide)
    title = slide_item.title.replace("1. ", "").replace("2. ", "").replace("3. ", "")
    _add_text(slide, slide_item.payload.get("section_no", ""), Inches(6.9), Inches(2.7), Inches(1.2), Inches(0.38), 17, True, ORANGE)
    _add_text(slide, title, Inches(6.88), Inches(3.12), Inches(5.25), Inches(0.65), 31, True, ORANGE)
    items = " · ".join(slide_item.payload.get("items", []))
    _add_text(slide, items, Inches(6.9), Inches(4.03), Inches(5.1), Inches(0.42), 11, False, MUTED)


def _style_table(
    table,
    header=True,
    first_col_gray=False,
    first_col_orange=False,
    body_font_size=11.0,
    header_font_size=11.5,
) -> None:
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.035)
            cell.margin_bottom = Inches(0.035)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            is_header = r_idx == 0 and header
            is_orange_col = first_col_orange and c_idx == 0 and not is_header
            is_gray_col = first_col_gray and c_idx == 0 and not is_header
            cell.fill.solid()
            if is_header or is_orange_col:
                cell.fill.fore_color.rgb = ORANGE
            elif is_gray_col:
                cell.fill.fore_color.rgb = LIGHT
            else:
                cell.fill.fore_color.rgb = WHITE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if is_orange_col else PP_ALIGN.LEFT
                for run in p.runs:
                    run.font.name = FONT
                    run.font.size = Pt(header_font_size if is_header else body_font_size)
                    run.font.color.rgb = WHITE if is_header or is_orange_col else DARK
                    run.font.bold = is_header or is_orange_col or is_gray_col




def _apply_table_row_heights(table, heights_in: Sequence[float]) -> None:
    """Set explicit row heights so PPT tables do not stretch vertically.

    python-pptx distributes rows across the height passed to add_table().
    When a slide has only a few rows, a large fixed table height makes each
    row expand unnaturally. Calling this after text insertion normalizes rows
    to the intended content height.
    """
    for idx, height in enumerate(heights_in):
        if idx < len(table.rows):
            table.rows[idx].height = Inches(height)

def _add_overview(prs: Presentation, report: ReportData) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    _add_content_header(slide, "1) 교육 개요", "01", "교육 개요")
    target = report.target_text
    if report.total_participants:
        target = f"{target} {report.total_participants}명".strip()
    data = [
        ("과정명", report.course_name),
        ("교육일정", report.schedule),
        ("교육방식", report.delivery_method),
        ("교육 대상", target),
        ("교육 목표", report.objective),
    ]
    table_shape = slide.shapes.add_table(5, 2, Inches(0.46), Inches(1.32), Inches(12.08), Inches(5.55))
    table = table_shape.table
    table.columns[0].width = Inches(2.20)
    table.columns[1].width = Inches(9.88)
    for i, (label, value) in enumerate(data):
        table.cell(i, 0).text = label
        table.cell(i, 1).text = str(value or "")
        table.rows[i].height = Inches(1.06 if i < 4 else 1.30)
    _style_table(table, header=False, first_col_orange=True, body_font_size=13.0)


def _add_curriculum(prs: Presentation, report: ReportData) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    _add_content_header(slide, "2) 커리큘럼", "01", "교육 개요")
    rows = report.curriculum[:8]
    row_count = max(2, len(rows) + 1)
    row_heights = [0.48] + [0.50] * (row_count - 1)
    table_h = sum(row_heights)
    shape = slide.shapes.add_table(row_count, 4, Inches(0.62), Inches(1.45), Inches(12.08), Inches(table_h))
    table = shape.table
    _apply_table_row_heights(table, row_heights)
    widths = [1.15, 2.1, 6.75, 2.08]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)
    headers = ["Day", "시간", "교육 내용", "강사/비고"]
    for c, header in enumerate(headers):
        table.cell(0, c).text = header
    if rows:
        for r_idx, row in enumerate(rows, start=1):
            for c_idx, value in enumerate([row.day, row.time, row.content, row.instructor]):
                table.cell(r_idx, c_idx).text = str(value or "")
    else:
        table.cell(1, 0).merge(table.cell(1, 3))
        table.cell(1, 0).text = ""
    _style_table(table, body_font_size=11.4, header_font_size=12.0)


def _add_survey_structure(prs: Presentation, report: ReportData, slide_item: SlideItem) -> None:
    questions = list(report.objective_questions) + list(report.subjective_questions)
    start = int(slide_item.payload.get("start", 0))
    end = int(slide_item.payload.get("end", min(start + 8, len(questions))))
    chunk = questions[start:end]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    title = "1) 설문 구성"
    if len(questions) > 8:
        title += f" ({int(slide_item.payload.get('page_index', 0)) + 1})"
    _add_content_header(slide, title, "02", "만족도 통계")
    row_count = max(2, len(chunk) + 1)
    row_heights = [0.48] + [0.55] * (row_count - 1)
    table_h = sum(row_heights)
    shape = slide.shapes.add_table(row_count, 2, Inches(0.62), Inches(1.45), Inches(12.08), Inches(table_h))
    table = shape.table
    _apply_table_row_heights(table, row_heights)
    table.columns[0].width = Inches(2.45)
    table.columns[1].width = Inches(9.63)
    table.cell(0, 0).text = "항목"
    table.cell(0, 1).text = "문항"
    for r_idx, q in enumerate(chunk, start=1):
        table.cell(r_idx, 0).text = q.section_label
        table.cell(r_idx, 1).text = f"{start + r_idx}. {strip_leading_question_numbers(q.question)}"
    if not chunk:
        table.cell(1, 0).merge(table.cell(1, 1))
        table.cell(1, 0).text = ""
    _style_table(table, first_col_gray=True, body_font_size=11.4, header_font_size=12.0)


def _summary_label(q: ObjectiveQuestion) -> str:
    if q.instructor_name:
        return f"{q.instructor_name}\n만족도"
    if q.course_module and q.section_label in {"과정 만족도", "강사 만족도"}:
        return q.course_module
    return q.section_label


def _add_summary(prs: Presentation, report: ReportData, slide_item: SlideItem) -> None:
    start = int(slide_item.payload.get("start", 0))
    end = int(slide_item.payload.get("end", min(start + 7, len(report.objective_questions))))
    questions = report.objective_questions[start:end]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    title = "2) 객관식 설문 결과"
    if len(report.objective_questions) > 7:
        title += f" ({int(slide_item.payload.get('page_index', 0)) + 1})"
    _add_content_header(slide, title, "02", "만족도 통계")

    _add_text(slide, "■ 만족도 요약", Inches(0.95), Inches(1.05), Inches(3.5), Inches(0.38), 15, True, DARK)
    if report.total_participants:
        count_line = (
            f"- 수강인원 총 {report.total_participants}명 중 {report.response_count}명 응답"
            f" ｜ 응답률 {report.response_rate}%"
        )
    else:
        count_line = f"- 총 {report.response_count}명 응답"
    _add_text(slide, count_line, Inches(1.18), Inches(1.39), Inches(8.8), Inches(0.34), 14, False, BLACK)

    chart_x, chart_y, chart_w, chart_h = 1.4, 1.82, 11.35, 4.3
    for value in range(7):
        y = chart_y + chart_h - (value / 6.0) * chart_h
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(chart_x), Inches(y), Inches(chart_w), Inches(0.012))
        line.fill.solid()
        line.fill.fore_color.rgb = GRID
        line.line.fill.background()
        _add_text(slide, f"{value:.1f}", Inches(0.78), Inches(y - 0.14), Inches(0.5), Inches(0.28), 11, False, MUTED, PP_ALIGN.RIGHT)

    y_axis = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(chart_x), Inches(chart_y), Inches(0.012), Inches(chart_h))
    y_axis.fill.solid()
    y_axis.fill.fore_color.rgb = RGBColor(42, 148, 191)
    y_axis.line.fill.background()

    n = max(1, len(questions))
    slot = chart_w / n
    bar_w = min(0.46, slot * 0.34)
    for idx, q in enumerate(questions):
        center_x = chart_x + slot * idx + slot / 2
        bar_h = max(0.0, min(chart_h, (q.average / 6.0) * chart_h))
        bar_y = chart_y + chart_h - bar_h
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(center_x - bar_w / 2),
            Inches(bar_y),
            Inches(bar_w),
            Inches(bar_h),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = ORANGE
        bar.line.fill.background()
        _add_text(slide, f"{q.average:.1f}", Inches(center_x - 0.35), Inches(bar_y - 0.30), Inches(0.7), Inches(0.28), 12.5, False, BLACK, PP_ALIGN.CENTER)
        label = _summary_label(q)
        _add_text(slide, label, Inches(center_x - slot * 0.48), Inches(chart_y + chart_h + 0.1), Inches(slot * 0.96), Inches(0.66), 11.2, False, MUTED, PP_ALIGN.CENTER, MSO_ANCHOR.TOP)


def _axis_max(max_count: int) -> Tuple[int, int]:
    if max_count <= 9:
        return max(1, max_count), 1
    step = max(1, math.ceil(max_count / 9))
    return int(math.ceil(max_count / step) * step), step


def _objective_question_font_size(text: str) -> float:
    length = len(str(text or ""))
    if length <= 58:
        return 14.5
    if length <= 90:
        return 13.2
    if length <= 125:
        return 12.2
    return 11.2


def _add_objective(prs: Presentation, report: ReportData, q_index: int) -> None:
    q = report.objective_questions[q_index]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    _add_content_header(slide, "2) 객관식 설문 결과", "02", "만족도 통계")
    question_text = f"■ {q.section_label} {q_index + 1}. {strip_leading_question_numbers(q.question)}"
    _add_text(
        slide,
        question_text,
        Inches(0.88),
        Inches(1.18),
        Inches(11.55),
        Inches(0.78),
        _objective_question_font_size(question_text),
        True,
        DARK,
        valign=MSO_ANCHOR.TOP,
    )

    raw_max = max(report.response_count, q.valid_responses, max(q.counts) if q.counts else 0, 1)
    axis_max, step = _axis_max(raw_max)
    chart_x, chart_y, chart_w, chart_h = 1.4, 2.02, 11.25, 4.13
    tick_values = list(range(0, axis_max + 1, step))
    for tick in tick_values:
        x = chart_x + chart_w * tick / axis_max
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(chart_y), Inches(0.012), Inches(chart_h))
        line.fill.solid()
        line.fill.fore_color.rgb = GRID
        line.line.fill.background()
        _add_text(slide, str(tick), Inches(x - 0.18), Inches(chart_y + chart_h + 0.06), Inches(0.36), Inches(0.26), 9.5, False, MUTED, PP_ALIGN.CENTER)

    count_rows = max(1, len(q.scale_labels))
    row_h = chart_h / count_rows
    colors = [PURPLE, BLUE, ORANGE, RGBColor(243, 154, 102), RGBColor(201, 201, 201), RGBColor(180, 180, 180), RGBColor(155, 155, 155)]
    for idx, (label, count) in enumerate(zip(q.scale_labels, q.counts)):
        row_y = chart_y + idx * row_h
        _add_text(slide, label, Inches(chart_x + 0.08), Inches(row_y), Inches(3.2), Inches(0.32), 13.6, False, DARK, valign=MSO_ANCHOR.BOTTOM)
        bar_y = row_y + 0.31
        bar_h = min(0.34, row_h * 0.42)
        bar_w = chart_w * count / axis_max if axis_max else 0
        if bar_w > 0:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(chart_x), Inches(bar_y), Inches(bar_w), Inches(bar_h))
            bar.fill.solid()
            bar.fill.fore_color.rgb = colors[idx % len(colors)]
            bar.line.fill.background()
            if bar_w >= 0.35:
                _add_text(slide, str(count), Inches(chart_x + max(0, bar_w - 0.42)), Inches(bar_y), Inches(0.35), Inches(bar_h), 11, True, WHITE, PP_ALIGN.RIGHT)
            else:
                _add_text(slide, str(count), Inches(chart_x + bar_w + 0.04), Inches(bar_y), Inches(0.35), Inches(bar_h), 10.5, True, DARK)


def _subjective_font_size(answers: Sequence[str]) -> float:
    total_chars = sum(len(str(a)) for a in answers)
    if len(answers) <= 4 and total_chars <= 420:
        return 13.0
    if total_chars <= 650:
        return 12.0
    return 10.7


def _add_subjective(prs: Presentation, report: ReportData, slide_item: SlideItem) -> None:
    q_index = int(slide_item.payload["question_index"])
    q = report.subjective_questions[q_index]
    chunk_index = int(slide_item.payload.get("chunk_index", 0))
    start = chunk_index * SUBJECTIVE_PER_SLIDE
    answers: Sequence[str] = q.answers[start : start + SUBJECTIVE_PER_SLIDE] or ["별도 의견 없음"]
    question_no = len(report.objective_questions) + q_index + 1
    suffix = slide_item.payload.get("page_suffix", "")

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    _add_content_header(slide, f"3) 주관식 설문 결과{suffix}", "02", "만족도 통계")
    _add_text(
        slide,
        f"■ {q.section_label} {question_no}. {strip_leading_question_numbers(q.question)}",
        Inches(0.68),
        Inches(1.2),
        Inches(12.0),
        Inches(0.68),
        16,
        True,
        DARK,
        valign=MSO_ANCHOR.TOP,
    )

    box = slide.shapes.add_textbox(Inches(0.88), Inches(1.82), Inches(11.7), Inches(4.95))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    font_size = _subjective_font_size(answers)
    for idx, answer in enumerate(answers):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        p.line_spacing = 1.18
        run = p.add_run()
        run.text = f"-  {answer}"
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.color.rgb = DARK


def _fit_image(image_bytes: bytes, box_w: float, box_h: float) -> Tuple[float, float]:
    with Image.open(io.BytesIO(image_bytes)) as image:
        w, h = image.size
    ratio = min(box_w / w, box_h / h)
    return w * ratio, h * ratio


def _add_photos(prs: Presentation, photos: Sequence[Tuple[str, bytes]], slide_item: SlideItem) -> None:
    page_index = int(slide_item.payload.get("page_index", 0))
    suffix = slide_item.payload.get("page_suffix", "")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_bg(slide)
    _add_content_header(slide, f"3) 현장 사진{suffix}", "03", "현장 사진")
    page_photos = photos[page_index * 6 : (page_index + 1) * 6]
    if not page_photos:
        return
    left, top, gap_x, gap_y = 0.62, 1.4, 0.14, 0.14
    box_w = (12.08 - gap_x * 2) / 3
    box_h = (5.25 - gap_y) / 2
    for idx, (name, data) in enumerate(page_photos[:6]):
        row, col = divmod(idx, 3)
        x = left + col * (box_w + gap_x)
        y = top + row * (box_h + gap_y)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(box_w), Inches(box_h))
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(241, 241, 241)
        bg.line.color.rgb = LINE
        try:
            pw, ph = _fit_image(data, box_w, box_h)
            px = x + (box_w - pw) / 2
            py = y + (box_h - ph) / 2
            slide.shapes.add_picture(io.BytesIO(data), Inches(px), Inches(py), Inches(pw), Inches(ph))
        except Exception:
            _add_text(slide, name, Inches(x + 0.1), Inches(y + 0.1), Inches(box_w - 0.2), Inches(box_h - 0.2), 10.5, False, MUTED, PP_ALIGN.CENTER)


def _add_thanks(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_full_background(slide)
    _add_text(slide, "THANK YOU", Inches(6.88), Inches(3.08), Inches(5.1), Inches(0.65), 34, True, ORANGE)
    _add_text(slide, "감사합니다", Inches(6.9), Inches(3.78), Inches(4.8), Inches(0.38), 16, True, ORANGE)


def generate_pptx(report: ReportData, photos: Sequence[Tuple[str, bytes]] | None = None) -> bytes:
    photos = photos or []
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    while len(prs.slides):
        xml_slides = prs.slides._sldIdLst
        xml_slides.remove(xml_slides[0])

    plan = build_slide_plan(report, photo_count=len(photos))
    for slide_item in plan:
        if slide_item.kind == "cover":
            _add_cover(prs, report)
        elif slide_item.kind == "toc":
            _add_toc(prs)
        elif slide_item.kind == "section":
            _add_section(prs, slide_item)
        elif slide_item.kind == "overview":
            _add_overview(prs, report)
        elif slide_item.kind == "curriculum":
            _add_curriculum(prs, report)
        elif slide_item.kind == "survey_structure":
            _add_survey_structure(prs, report, slide_item)
        elif slide_item.kind == "summary":
            _add_summary(prs, report, slide_item)
        elif slide_item.kind == "objective":
            _add_objective(prs, report, int(slide_item.payload["question_index"]))
        elif slide_item.kind == "subjective":
            _add_subjective(prs, report, slide_item)
        elif slide_item.kind == "photos":
            _add_photos(prs, photos, slide_item)
        elif slide_item.kind == "thanks":
            _add_thanks(prs)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()
