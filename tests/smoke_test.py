from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.excel_parser import parse_excel
from src.ppt_renderer import generate_pptx
from src.preview_renderer import render_slide_html
from src.slide_plan import build_slide_plan


def run(paths):
    for path in paths:
        file_path = Path(path)
        data = parse_excel(file_path.read_bytes(), file_path.name)
        assert data.response_count > 0, f"응답자 수 인식 실패: {file_path.name}"
        assert data.objective_questions, f"객관식 인식 실패: {file_path.name}"
        assert all(not q.question.lstrip().startswith(("1. ", "2. ", "1) ", "2) ")) for q in data.objective_questions), \
            f"문항 앞 원본 번호 제거 실패: {file_path.name}"
        plan = build_slide_plan(data)
        survey_slide = next(item for item in plan if item.kind == "survey_structure")
        survey_html = render_slide_html(data, survey_slide)
        assert "1. 1." not in survey_html and "2. 1." not in survey_html, f"문항 번호 중복: {file_path.name}"
        subjective_slide = next((item for item in plan if item.kind == "subjective"), None)
        if subjective_slide and subjective_slide.payload.get("answers"):
            subjective_html = render_slide_html(data, subjective_slide)
            assert '<li>' in subjective_html or 'class="empty-opinion"' in subjective_html, f"주관식 불릿 미적용: {file_path.name}"
        ppt = generate_pptx(data)
        assert ppt[:2] == b"PK", f"PPTX 생성 실패: {file_path.name}"
        output = ROOT / "tests" / f"_smoke_{file_path.stem}.pptx"
        output.write_bytes(ppt)
        print(
            f"OK | {file_path.name} | 응답 {data.response_count}명 | "
            f"객관식 {len(data.objective_questions)} | 주관식 {len(data.subjective_questions)} | {output.name}"
        )


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("사용법: python tests/smoke_test.py 파일1.xlsx [파일2.xlsx ...]")
    run(sys.argv[1:])
