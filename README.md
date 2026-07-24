# 교육만족도 결과보고서 생성기 V2 (Beta)
🌍https://educationsatisfactionreportmvp-tdhju6x2yaawphxhcmxg2b.streamlit.app/

형식이 다른 설문 엑셀과 구글폼 응답 파일을 자동 분석하고, 장표별 내용을 웹에서 수정한 뒤 **편집 가능한 16:9 PPTX**를 생성하는 Streamlit 앱입니다.

## V2 주요 기능

- 문항이 **열 / 응답자가 행**인 일반·구글폼 로우데이터 지원
- 문항이 **행 / 응답자가 열**인 전치형 데이터 지원
- 여러 시트 중 설문 원본 시트 자동 선택
- 응답 데이터를 기반으로 문항 유형 자동 분류
  - 1~5점, 1~7점, 0~10점 척도
  - 단일선택
  - 복수선택
  - 주관식
- 추천 의향 0~10점 문항은 **평균 점수 + 점수별 응답 분포**만 표시
  - 추천자·중립자·비추천자·NPS는 계산하지 않음
- 문항에서 실명 강사 자동 추출 및 강사 수 제한 없이 그룹화
- 강사가 2명 이상이면 강사 비교 장표 자동 생성
- 공통 척도 문항 → 강사 비교 → 강사별 상세 → 선택형 → 주관식 순으로 자동 구성
- 단일선택 원형그래프, 복수선택 가로막대그래프 생성
- 규칙형 종합 시사점 초안 생성 후 웹에서 수정 가능
- HTML 미리보기와 PPT가 동일한 `ReportData`와 `slide_plan`을 사용
- 오렌지 컬러, 제목, 폰트, 레이아웃 등 기존 멀티캠퍼스 결과보고서 스타일 유지
- PPT의 텍스트·표·도형·차트 편집 가능

## 실행 방법

### Windows

```bash
cd education_satisfaction_report_mvp
py -m venv .venv
source .venv/Scripts/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

PowerShell에서는 가상환경을 다음과 같이 활성화합니다.

```powershell
.venv\Scripts\Activate.ps1
```

## 샘플 로우데이터 추가

아래 폴더에 `.xlsx` 또는 `.xlsm` 파일을 추가하면 앱의 **샘플 로우데이터** 목록에 자동 표시됩니다.

```text
sample_data/
├─ 01_표준형_만족도_로우데이터_샘플.xlsx
├─ 02_전치형_만족도_로우데이터_샘플.xlsx
├─ 03_구글폼_혼합문항_샘플.xlsx
└─ 04_다강사_복수선택_샘플.xlsx
```

실제 배포본에는 개인정보와 고객사 민감정보를 제거한 파일만 넣으세요.

## 주요 폴더 구조

```text
education_satisfaction_report_mvp/
├─ app.py
├─ requirements.txt
├─ assets/
│  ├─ 결과보고서_표준양식_참고용.pptx
│  ├─ 만족도_조사_로우데이터_템플릿.xlsx
│  ├─ multicampus_section_background.png
│  └─ multicampus_logo.png
├─ sample_data/
├─ src/
│  ├─ models.py
│  ├─ excel_parser.py
│  ├─ slide_plan.py
│  ├─ preview_renderer.py
│  ├─ ppt_renderer.py
│  ├─ live_input.py
│  └─ live_input_frontend/
├─ docs/
└─ tests/
```

## 문항 판별 기준

프로그램은 문항 제목만으로 유형을 결정하지 않습니다. 실제 응답값의 숫자 비율, 값 범위, 고유값 수, 반복성, 한 셀 내 복수 토큰 여부를 우선 분석하고 문항 문구는 보조 신호로만 사용합니다.

```text
응답값 수집
→ 숫자형 척도 여부 분석
→ 선택지 반복 패턴 분석
→ 단일/복수선택 판단
→ 나머지는 주관식 처리
```

## V2 장표 순서

```text
교육 개요
→ 만족도 요약(척도별 분리)
→ 공통 척도 문항
→ 강사 비교(2명 이상)
→ 강사별 상세
→ 단일선택 / 복수선택
→ 주관식
→ 종합 시사점
→ 현장 사진
```

## 테스트

```bash
python tests/smoke_test.py sample_data/03_구글폼_혼합문항_샘플.xlsx sample_data/04_다강사_복수선택_샘플.xlsx
```

상세 변경 사항은 `docs/CHANGELOG_V2.md`에서 확인할 수 있습니다.
