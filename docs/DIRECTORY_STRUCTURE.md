# VSCode 기준 디렉터리 구조

```text
education_satisfaction_report_mvp/
│
├─ app.py
│   Streamlit 진입점입니다. 파일 업로드, 페이지 이동, 장표별 입력 폼,
│   실시간 미리보기, PPT 생성·다운로드를 담당합니다.
│
├─ src/
│   ├─ models.py
│   │   엑셀 형식과 무관하게 사용하는 표준 데이터 구조입니다.
│   │
│   ├─ excel_parser.py
│   │   시트 탐색, 행/열 방향 판별, 리커트 변환, 메타데이터와 강사명 추출을 담당합니다.
│   │
│   ├─ slide_plan.py
│   │   객관식 문항 수, 주관식 응답 수, 사진 수에 따라 장표 수를 동적으로 계산합니다.
│   │
│   ├─ preview_renderer.py
│   │   현재 장표를 16:9 HTML/CSS로 렌더링합니다.
│   │
│   └─ ppt_renderer.py
│       동일한 ReportData를 사용해 편집 가능한 PPTX를 생성합니다.
│
├─ assets/
│   └─ 결과보고서_표준양식_참고용.pptx
│       사용자가 제공한 목표 양식입니다. 2차 개발에서 좌표·스타일 매핑 기준으로 사용합니다.
│
├─ tests/
│   └─ smoke_test.py
│       제공된 세 종류 엑셀을 파싱하고 PPT를 생성하는 최소 통합 테스트입니다.
│
├─ docs/
│   ├─ DIRECTORY_STRUCTURE.md
│   └─ NEXT_PHASE.md
│
├─ requirements.txt
├─ run.bat
├─ run.sh
└─ .gitignore
```

## 권장 개발 순서

1. `src/excel_parser.py`에서 추가 엑셀 유형 지원
2. `src/models.py`의 표준 필드 확정
3. `src/slide_plan.py`에서 장표 추가·삭제 규칙 확정
4. `src/preview_renderer.py`와 `src/ppt_renderer.py`를 동일 규격으로 고도화
5. 마지막에 `app.py`의 사용성과 디자인 개선
