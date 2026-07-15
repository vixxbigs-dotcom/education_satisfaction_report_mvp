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
    page_title="교육만족도 결과보고서 생성기 (Beta)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {padding-top: 2.75rem !important; padding-bottom: 2rem; max-width: 1800px;}
      /* Streamlit의 메인 영역이 자체 스크롤을 담당합니다.
         overflow: visible로 바꾸면 전체 페이지 스크롤이 사라질 수 있으므로
         세로 스크롤을 명시적으로 복원합니다. */
      section[data-testid="stMain"] {overflow-y:auto !important; overflow-x:hidden !important;}
      .stMainBlockContainer {overflow: visible !important;}
      .app-title {font-size:clamp(1.65rem, 2.25vw, 2.25rem); line-height:1.42; font-weight:800; margin:.25rem 0 .28rem; color:#ED702C; white-space:normal; word-break:keep-all; overflow:visible; padding:.32rem 0 .2rem;}
      .app-sub {color:#6b7280; margin-bottom:1.1rem;}
      .panel-title {font-size:1.45rem; font-weight:800; margin-bottom:.65rem;}
      div[data-testid="stFileUploader"] {border:1px solid #d9e1e8; border-radius:12px; padding:.4rem .8rem;}
      .status-card {background:#fff7f2; border:1px solid #f3d3c1; border-radius:10px; padding:.8rem 1rem; margin:.4rem 0 1rem;}
      .small-muted {font-size:1rem; color:#6b7280;}
      div[data-testid="stVerticalBlockBorderWrapper"] {border-radius:14px;}
      .stNumberInput label, .stSelectbox label, .stCheckbox label, .stFileUploader label {font-size:1.08rem !important; font-weight:700 !important;}
      .stDataFrame, .stDataEditor {font-size:1.05rem;}
      button[data-testid="stBaseButton-primary"],
      div.stButton > button[kind="primary"] {background:#ED702C !important; border-color:#ED702C !important; color:#fff !important;}
      button[data-testid="stBaseButton-primary"]:hover,
      div.stButton > button[kind="primary"]:hover {background:#D85C1E !important; border-color:#D85C1E !important; color:#fff !important;}
      #editor-preview-split-marker {height:0; margin:0; padding:0;}
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
    # 표지에서 실제로 보이는 항목만 노출합니다.
    report.company_name = live_text("회사 이름", report.company_name, key="field_company_name")
    report.course_name = live_text("교육 과정명", report.course_name, key="field_course_name")
    report.schedule = live_text("교육 일정", report.schedule, key="field_schedule")

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
        total = st.number_input(
            "총 수강인원 (0은 미입력)",
            min_value=0,
            value=int(report.total_participants or 0),
            step=1,
            key="field_total_participants",
        )
        report.total_participants = int(total) if total else None
    with right:
        report.response_count = st.number_input(
            "응답자 수",
            min_value=0,
            value=int(report.response_count),
            step=1,
            key="field_response_count",
        )
    st.caption("문항별 평균은 각 객관식 결과 장표의 응답 분포를 수정하면 자동으로 다시 계산됩니다.")

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
    original_answers = list(q.answers)
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
    if q.answers != original_answers:
        # The slide plan is built before the editor is drawn. Rerun once after
        # applying subjective edits so continuation pages and the preview use
        # the exact same answer list immediately.
        st.session_state.report = report
        st.session_state.ppt_bytes = None
        st.rerun()
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


def _render_editor(report: ReportData, slide) -> bool:
    editable_kinds = {
        "cover",
        "overview",
        "curriculum",
        "survey_structure",
        "summary",
        "objective",
        "subjective",
        "photos",
    }
    if slide.kind not in editable_kinds:
        return False

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
    return True

def _sync_slide_jump() -> None:
    selected_id = st.session_state.get("slide_jump_id")
    if selected_id:
        st.session_state.slide_id = selected_id


st.markdown('<div class="app-title">교육만족도 결과보고서 생성기 (Beta)</div>', unsafe_allow_html=True)
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
editable_kinds = {"cover", "overview", "curriculum", "survey_structure", "summary", "objective", "subjective", "photos"}

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

if current_slide.kind in editable_kinds:
    st.markdown('<div id="editor-preview-split-marker"></div>', unsafe_allow_html=True)
    left, right = st.columns([0.42, 0.58], gap="small")
    with left:
        with st.container(border=True):
            _render_editor(report, current_slide)
    with right:
        st.markdown('<div class="panel-title">출력 미리보기</div>', unsafe_allow_html=True)
        components.html(preview_document, height=610, scrolling=False)

    # Streamlit columns are fixed by default. This lightweight client-side
    # splitter adds a draggable vertical divider without changing report data
    # or forcing the current slide to rerun. The last width is remembered in
    # the browser for the remainder of the session.
    components.html(
        r"""
        <script>
        (() => {
          const doc = window.parent.document;
          const marker = doc.getElementById('editor-preview-split-marker');
          if (!marker) return;

          const markerRect = marker.getBoundingClientRect();
          const scope = marker.closest('[data-testid="stVerticalBlock"]') || doc;
          const rows = Array.from(scope.querySelectorAll('[data-testid="stHorizontalBlock"]'));
          const row = rows.find((candidate) => {
            const rect = candidate.getBoundingClientRect();
            const cols = candidate.querySelectorAll(':scope > [data-testid="stColumn"]');
            return cols.length === 2 && rect.top >= markerRect.bottom - 3;
          });
          if (!row) return;

          const columns = row.querySelectorAll(':scope > [data-testid="stColumn"]');
          if (columns.length !== 2) return;
          const left = columns[0];
          const right = columns[1];

          row.querySelectorAll('[data-report-splitter="true"]').forEach((node) => node.remove());
          row.style.position = 'relative';
          row.style.gap = '0px';
          row.style.alignItems = 'stretch';

          const splitter = doc.createElement('div');
          splitter.setAttribute('data-report-splitter', 'true');
          splitter.title = '좌우로 드래그하여 입력부와 미리보기 너비 조절';
          Object.assign(splitter.style, {
            position: 'absolute',
            top: '0',
            bottom: '0',
            width: '18px',
            transform: 'translateX(-9px)',
            cursor: 'col-resize',
            zIndex: '50',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            touchAction: 'none',
            userSelect: 'none'
          });

          const line = doc.createElement('div');
          Object.assign(line.style, {
            width: '2px',
            height: '100%',
            minHeight: '420px',
            background: '#D8D8D8',
            borderRadius: '999px',
            transition: 'background .15s ease, width .15s ease'
          });
          splitter.appendChild(line);
          row.appendChild(splitter);

          let stored = 42;
          try {
            const parsed = Number(window.parent.sessionStorage.getItem('education_report_split_pct'));
            if (Number.isFinite(parsed)) stored = parsed;
          } catch (_) {}

          const apply = (pct) => {
            const safe = Math.max(27, Math.min(68, pct));
            left.style.flex = `0 0 ${safe}%`;
            left.style.width = `${safe}%`;
            left.style.maxWidth = `${safe}%`;
            right.style.flex = '1 1 0';
            right.style.width = `${100 - safe}%`;
            right.style.maxWidth = 'none';
            left.style.paddingRight = '14px';
            right.style.paddingLeft = '14px';
            splitter.style.left = `${safe}%`;
            return safe;
          };

          let current = apply(stored);
          let dragging = false;

          const move = (event) => {
            if (!dragging) return;
            const rect = row.getBoundingClientRect();
            current = apply(((event.clientX - rect.left) / rect.width) * 100);
          };
          const stop = () => {
            if (!dragging) return;
            dragging = false;
            line.style.background = '#D8D8D8';
            line.style.width = '2px';
            doc.body.style.cursor = '';
            doc.body.style.userSelect = '';
            try {
              window.parent.sessionStorage.setItem('education_report_split_pct', String(current));
            } catch (_) {}
          };

          splitter.addEventListener('pointerdown', (event) => {
            dragging = true;
            splitter.setPointerCapture(event.pointerId);
            line.style.background = '#ED702C';
            line.style.width = '4px';
            doc.body.style.cursor = 'col-resize';
            doc.body.style.userSelect = 'none';
            event.preventDefault();
          });
          splitter.addEventListener('pointermove', move);
          splitter.addEventListener('pointerup', stop);
          splitter.addEventListener('pointercancel', stop);
          splitter.addEventListener('mouseenter', () => { line.style.background = '#ED702C'; });
          splitter.addEventListener('mouseleave', () => { if (!dragging) line.style.background = '#D8D8D8'; });
        })();
        </script>
        """,
        height=0,
        scrolling=False,
    )
else:
    # 편집 입력부가 없는 구분장/목차/감사 페이지에서는 이전 장표에서
    # 삽입된 드래그 분할선과 열 너비 인라인 스타일을 완전히 정리합니다.
    components.html(
        r'''<script>
        (() => {
          const doc = window.parent.document;
          const splitters = Array.from(doc.querySelectorAll('[data-report-splitter="true"]'));
          splitters.forEach((splitter) => {
            const row = splitter.parentElement;
            if (row) {
              const columns = row.querySelectorAll(':scope > [data-testid="stColumn"]');
              columns.forEach((column) => {
                column.style.removeProperty('flex');
                column.style.removeProperty('width');
                column.style.removeProperty('max-width');
                column.style.removeProperty('padding-left');
                column.style.removeProperty('padding-right');
              });
              row.style.removeProperty('position');
              row.style.removeProperty('gap');
              row.style.removeProperty('align-items');
            }
            splitter.remove();
          });
        })();
        </script>''',
        height=0,
        scrolling=False,
    )

    st.markdown('<div class="panel-title" style="text-align:center;">출력 미리보기</div>', unsafe_allow_html=True)
    preview_left, preview_center, preview_right = st.columns([0.11, 0.78, 0.11])
    with preview_center:
        compact_preview_document = preview_document.replace(
            '</head>',
            '<style>html,body{width:100%;min-height:100%;display:flex;justify-content:center;align-items:flex-start;}body>.ppt-stage{width:min(100%,1120px);margin:0 auto;padding:8px;}</style></head>',
        )
        components.html(compact_preview_document, height=555, scrolling=False)

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

st.caption("v1.7.1 Beta: 전체 화면 세로 스크롤 복원, 상단 잘림 보정, 비편집 장표 중앙 미리보기")
