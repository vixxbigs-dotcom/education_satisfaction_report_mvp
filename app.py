from __future__ import annotations

import re
from typing import List, Tuple

import streamlit as st
import streamlit.components.v1 as components

from src.excel_parser import is_no_opinion, parse_excel
from src.live_input import live_text
from src.models import CurriculumRow, ReportData
from src.ppt_renderer import generate_pptx
from src.preview_renderer import PREVIEW_CSS, render_slide_html
from src.slide_plan import build_slide_plan


st.set_page_config(
    page_title="교육만족도 결과보고서 생성기",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1800px;}
      .app-title {font-size:2rem; font-weight:800; margin-bottom:.2rem; color:#ED702C;}
      .app-sub {color:#6b7280; margin-bottom:1.1rem;}
      .panel-title {font-size:1.25rem; font-weight:800; margin-bottom:.6rem;}
      div[data-testid="stFileUploader"] {border:1px solid #d9e1e8; border-radius:12px; padding:.4rem .8rem;}
      .status-card {background:#fff7f2; border:1px solid #f3d3c1; border-radius:10px; padding:.8rem 1rem; margin:.4rem 0 1rem;}
      .small-muted {font-size:.85rem; color:#6b7280;}
      div[data-testid="stVerticalBlockBorderWrapper"] {border-radius:14px;}
    </style>
    """,
    unsafe_allow_html=True,
)


def _clear_editor_state() -> None:
    prefixes = ("field_", "obj_", "subj_", "live_", "curriculum_editor", "survey_structure_editor", "photo_uploader")
    for key in list(st.session_state.keys()):
        if key.startswith(prefixes):
            st.session_state.pop(key, None)


def _reset_report() -> None:
    for key in ["report", "upload_signature", "slide_index", "slide_id", "slide_jump_id", "ppt_bytes", "photos"]:
        st.session_state.pop(key, None)
    _clear_editor_state()


def _safe_filename(text: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", text or "결과보고서")
    return re.sub(r"\s+", "_", name).strip("_") + ".pptx"


def _photo_payload() -> List[Tuple[str, bytes]]:
    return st.session_state.get("photos", [])


def _render_basic_editor(report: ReportData) -> None:
    report.company_name = live_text("회사 이름", report.company_name, key="field_company_name")
    report.course_name = live_text("교육 과정명", report.course_name, key="field_course_name")
    report.report_title = live_text("보고서 이름", report.report_title, key="field_report_title")
    report.schedule = live_text("교육 일정", report.schedule, key="field_schedule")
    instructor_text = live_text(
        "강사명 · 직책 (줄바꿈으로 구분)",
        "\n".join(report.instructors),
        key="field_instructors",
        multiline=True,
        height=110,
    )
    report.instructors = [v.strip() for v in instructor_text.splitlines() if v.strip()]

def _render_overview_editor(report: ReportData) -> None:
    report.course_name = live_text("과정명", report.course_name, key="field_course_name")
    report.schedule = live_text("교육일정", report.schedule, key="field_schedule")
    report.delivery_method = st.selectbox(
        "교육방식",
        ["", "대면 교육", "비대면 교육", "혼합 교육"],
        index=["", "대면 교육", "비대면 교육", "혼합 교육"].index(report.delivery_method)
        if report.delivery_method in ["", "대면 교육", "비대면 교육", "혼합 교육"] else 0,
        key="field_delivery_method",
    )
    report.target_text = live_text("교육 대상", report.target_text, key="field_target_text")
    total = st.number_input(
        "총 수강인원 (모르면 0)",
        min_value=0,
        value=int(report.total_participants or 0),
        step=1,
        key="field_total_participants",
    )
    report.total_participants = int(total) if total > 0 else None
    report.objective = live_text(
        "교육 목표",
        report.objective,
        key="field_objective",
        multiline=True,
        height=140,
    )

def _render_curriculum_editor(report: ReportData) -> None:
    rows = [vars(row).copy() for row in report.curriculum]
    if not rows:
        rows = [{"day": "", "time": "", "content": "", "instructor": ""}]
    edited = st.data_editor(
        rows,
        column_config={
            "day": st.column_config.TextColumn("Day", width="small"),
            "time": st.column_config.TextColumn("시간", width="medium"),
            "content": st.column_config.TextColumn("교육 내용", width="large"),
            "instructor": st.column_config.TextColumn("강사/비고", width="medium"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="curriculum_editor",
    )
    report.curriculum = [CurriculumRow(**row) for row in edited if any(str(v or "").strip() for v in row.values())]


def _render_survey_structure_editor(report: ReportData) -> None:
    rows = []
    for q in report.objective_questions:
        rows.append({"type": "객관식", "section": q.section_label, "question": q.question, "id": q.id})
    for q in report.subjective_questions:
        rows.append({"type": "주관식", "section": q.section_label, "question": q.question, "id": q.id})
    edited = st.data_editor(
        rows,
        column_config={
            "type": st.column_config.TextColumn("유형", disabled=True, width="small"),
            "section": st.column_config.TextColumn("항목", width="medium"),
            "question": st.column_config.TextColumn("문항", width="large"),
            "id": None,
        },
        use_container_width=True,
        hide_index=True,
        key="survey_structure_editor",
    )
    objective_map = {q.id: q for q in report.objective_questions}
    subjective_map = {q.id: q for q in report.subjective_questions}
    for row in edited:
        target = objective_map.get(row["id"]) or subjective_map.get(row["id"])
        if target:
            target.section_label = str(row["section"] or "").strip()
            target.question = str(row["question"] or "").strip()


def _render_summary_editor(report: ReportData) -> None:
    left, right = st.columns(2)
    with left:
        total = st.number_input("총 수강인원 (0은 미입력)", min_value=0, value=int(report.total_participants or 0), step=1, key="field_total_participants")
        report.total_participants = int(total) if total else None
    with right:
        report.response_count = st.number_input("응답자 수", min_value=0, value=int(report.response_count), step=1, key="field_response_count")
    st.caption("평균 점수는 각 객관식 장표에서 응답 수를 수정하면 자동으로 다시 계산됩니다.")
    st.dataframe(
        [{"항목": q.section_label, "문항": q.question, "평균": q.average, "유효 응답": q.valid_responses} for q in report.objective_questions],
        use_container_width=True,
        hide_index=True,
    )


def _render_objective_editor(report: ReportData, q_index: int) -> None:
    q = report.objective_questions[q_index]
    q.section_label = live_text("항목명", q.section_label, key=f"obj_section_{q.id}")
    q.question = live_text(
        "문항 내용",
        q.question,
        key=f"obj_question_{q.id}",
        multiline=True,
        height=100,
    )
    st.markdown("**응답 분포**")
    columns = st.columns(len(q.scale_labels))
    new_counts = []
    for idx, (column, label, count) in enumerate(zip(columns, q.scale_labels, q.counts)):
        with column:
            new_counts.append(
                st.number_input(label, min_value=0, value=int(count), step=1, key=f"obj_count_{q.id}_{idx}")
            )
    q.counts = [int(v) for v in new_counts]
    q.recalculate()
    st.info(f"재계산 평균: **{q.average:.2f}점** · 유효 응답: **{q.valid_responses}명**")

def _render_subjective_editor(report: ReportData, q_index: int) -> None:
    q = report.subjective_questions[q_index]
    q.section_label = live_text("항목명", q.section_label, key=f"subj_section_{q.id}")
    q.question = live_text(
        "문항 내용",
        q.question,
        key=f"subj_question_{q.id}",
        multiline=True,
        height=90,
    )
    remove_none = st.checkbox("'없음'류 응답 제외", value=q.category == "subjective_bad", key=f"subj_filter_{q.id}")
    answer_text = live_text(
        "응답 내용 (한 줄에 한 응답)",
        "\n".join(q.answers),
        key=f"subj_answers_{q.id}",
        multiline=True,
        height=330,
        debounce=280,
    )
    answers = [line.strip() for line in answer_text.splitlines() if line.strip()]
    if remove_none:
        answers = [answer for answer in answers if not is_no_opinion(answer)]
    q.answers = answers
    st.caption(f"현재 유효 응답 {len(q.answers)}건 · 한 장당 최대 5건씩 자동 분할됩니다.")

def _render_photo_editor() -> None:
    uploaded = st.file_uploader(
        "현장 사진 업로드",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="photo_uploader",
    )
    if uploaded:
        st.session_state.photos = [(file.name, file.getvalue()) for file in uploaded]
    photos = _photo_payload()
    st.caption(f"현재 {len(photos)}장 · 슬라이드당 최대 6장으로 배치됩니다.")
    if photos:
        st.write([name for name, _ in photos])


def _render_editor(report: ReportData, slide) -> None:
    st.markdown(f'<div class="panel-title">입력 설정 · {slide.title}</div>', unsafe_allow_html=True)
    if slide.kind == "cover":
        _render_basic_editor(report)
    elif slide.kind == "overview":
        _render_overview_editor(report)
    elif slide.kind == "curriculum":
        _render_curriculum_editor(report)
    elif slide.kind == "survey_structure":
        _render_survey_structure_editor(report)
    elif slide.kind == "summary":
        _render_summary_editor(report)
    elif slide.kind == "objective":
        _render_objective_editor(report, slide.payload["question_index"])
    elif slide.kind == "subjective":
        _render_subjective_editor(report, slide.payload["question_index"])
    elif slide.kind == "photos":
        _render_photo_editor()
    else:
        st.info("이 장표는 표준 고정 장표입니다. 입력값을 수정할 필요가 없습니다.")


def _sync_slide_jump() -> None:
    selected_id = st.session_state.get("slide_jump_id")
    if selected_id:
        st.session_state.slide_id = selected_id


st.markdown('<div class="app-title">교육만족도 결과보고서 생성기</div>', unsafe_allow_html=True)
st.markdown('<div class="app-sub">엑셀을 업로드하면 분석 결과가 자동 입력되고, 장표별로 수정하면서 PPT를 생성할 수 있습니다.</div>', unsafe_allow_html=True)

upload_col, reset_col = st.columns([8, 1])
with upload_col:
    uploaded_excel = st.file_uploader("설문 엑셀 업로드", type=["xlsx", "xlsm"], key="excel_uploader")
with reset_col:
    st.write("")
    st.write("")
    if st.button("전체 초기화", use_container_width=True):
        _reset_report()
        st.rerun()

if uploaded_excel:
    signature = f"{uploaded_excel.name}:{uploaded_excel.size}"
    if st.session_state.get("upload_signature") != signature:
        try:
            with st.spinner("엑셀 구조와 설문 결과를 분석하고 있습니다..."):
                _clear_editor_state()
                st.session_state.report = parse_excel(uploaded_excel.getvalue(), uploaded_excel.name)
                st.session_state.upload_signature = signature
                st.session_state.slide_index = 0
                st.session_state.slide_id = "cover"
                st.session_state.slide_jump_id = "cover"
                st.session_state.ppt_bytes = None
                st.session_state.photos = []
        except Exception as exc:
            st.error(f"엑셀 분석에 실패했습니다: {exc}")
            st.stop()

if "report" not in st.session_state:
    st.info("먼저 설문 엑셀 파일을 업로드해 주세요. 엑셀에서 찾지 못한 정보는 공란으로 생성됩니다.")
    st.stop()

report: ReportData = st.session_state.report
photo_names = [name for name, _ in _photo_payload()]
plan = build_slide_plan(report, photo_count=len(photo_names))
slide_ids = [item.id for item in plan]

requested_id = st.session_state.get("slide_id")
if requested_id not in slide_ids:
    # If a dynamic continuation page disappears while editing, stay on the
    # closest page for the same question instead of jumping to the cover.
    replacement_id = None
    if isinstance(requested_id, str) and requested_id.startswith("subjective-"):
        base_id = requested_id.rsplit("-", 1)[0]
        replacement_id = next((sid for sid in slide_ids if sid.startswith(base_id)), None)
    elif isinstance(requested_id, str) and requested_id.startswith("survey-structure"):
        replacement_id = next((sid for sid in slide_ids if sid.startswith("survey-structure")), None)
    elif isinstance(requested_id, str) and requested_id.startswith("summary"):
        replacement_id = next((sid for sid in slide_ids if sid.startswith("summary")), None)
    if replacement_id is None:
        fallback_index = min(max(int(st.session_state.get("slide_index", 0)), 0), len(plan) - 1)
        replacement_id = slide_ids[fallback_index]
    requested_id = replacement_id

current_index = slide_ids.index(requested_id)
st.session_state.slide_id = requested_id
st.session_state.slide_index = current_index
if st.session_state.get("slide_jump_id") not in slide_ids:
    st.session_state.slide_jump_id = requested_id

with st.expander("자동 분석 결과 및 확인 사항", expanded=False):
    st.markdown(
        f"""
        <div class="status-card">
          <b>파일:</b> {report.source_filename}<br/>
          <b>응답자:</b> {report.response_count}명<br/>
          <b>객관식:</b> {len(report.objective_questions)}문항 · <b>주관식:</b> {len(report.subjective_questions)}문항<br/>
          <b>강사:</b> {', '.join(report.instructors) if report.instructors else '미추출'}
        </div>
        """,
        unsafe_allow_html=True,
    )
    for diagnostic in report.diagnostics:
        st.write(f"- {diagnostic}")

nav_left, nav_center, nav_right, nav_select = st.columns([1, 2, 1, 4])
with nav_left:
    if st.button("〈 이전", disabled=current_index == 0, use_container_width=True):
        target_index = max(0, current_index - 1)
        st.session_state.slide_index = target_index
        st.session_state.slide_id = slide_ids[target_index]
        st.session_state.slide_jump_id = slide_ids[target_index]
        st.rerun()
with nav_center:
    st.markdown(f"<div style='text-align:center;font-weight:800;padding-top:.45rem'>{current_index + 1} / {len(plan)} · {plan[current_index].title}</div>", unsafe_allow_html=True)
with nav_right:
    if st.button("다음 〉", disabled=current_index >= len(plan) - 1, use_container_width=True):
        target_index = min(len(plan) - 1, current_index + 1)
        st.session_state.slide_index = target_index
        st.session_state.slide_id = slide_ids[target_index]
        st.session_state.slide_jump_id = slide_ids[target_index]
        st.rerun()
with nav_select:
    st.selectbox(
        "장표 바로가기",
        options=slide_ids,
        format_func=lambda slide_id: f"{slide_ids.index(slide_id) + 1}. {plan[slide_ids.index(slide_id)].title}",
        label_visibility="collapsed",
        key="slide_jump_id",
        on_change=_sync_slide_jump,
    )

current_slide = plan[current_index]
left, right = st.columns([0.42, 0.58], gap="large")
with left:
    with st.container(border=True):
        _render_editor(report, current_slide)
with right:
    st.markdown('<div class="panel-title">출력 미리보기</div>', unsafe_allow_html=True)

    # st.markdown()은 줄 앞 공백이 있는 HTML을 Markdown 코드 블록(<pre>)으로
    # 해석할 수 있습니다. 미리보기는 독립 HTML 컴포넌트로 렌더링하여
    # 소스 코드가 아니라 실제 16:9 슬라이드 화면을 표시합니다.
    preview_document = f"""
    <!doctype html>
    <html lang="ko">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
          html, body {{
            margin: 0;
            padding: 0;
            background: transparent;
            overflow: hidden;
          }}
          * {{ box-sizing: border-box; }}
        </style>
        {PREVIEW_CSS}
      </head>
      <body>
        {render_slide_html(report, current_slide, photo_names)}
      </body>
    </html>
    """
    components.html(preview_document, height=610, scrolling=False)

st.session_state.report = report

st.divider()
action_left, action_right = st.columns([1, 1])
with action_left:
    if st.button("PPT 생성 / 최신 내용 반영", type="primary", use_container_width=True):
        try:
            with st.spinner("편집 가능한 PPTX를 생성하고 있습니다..."):
                st.session_state.ppt_bytes = generate_pptx(report, _photo_payload())
            st.success("PPT 생성이 완료되었습니다.")
        except Exception as exc:
            st.error(f"PPT 생성 중 오류가 발생했습니다: {exc}")
with action_right:
    st.download_button(
        "PPT 다운로드",
        data=st.session_state.get("ppt_bytes") or b"",
        file_name=_safe_filename(report.report_title or report.course_name),
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        disabled=not bool(st.session_state.get("ppt_bytes")),
        use_container_width=True,
    )

st.caption("v1.2: 멀티캠퍼스 주황색 결과보고서 UI를 HTML 미리보기와 편집 가능한 PPTX에 동일하게 적용합니다.")
