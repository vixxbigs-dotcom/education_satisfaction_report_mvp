# 📊 교육만족도 결과보고서 생성기 (Beta)

엑셀 설문 데이터를 분석해  
교육 결과보고서 PPT를 자동으로 생성하는 Streamlit 웹앱입니다.

---

## ✨ 주요 기능

- 📁 엑셀 파일 업로드
- 🔍 설문 문항 자동 추출
- 📈 객관식 평균 및 응답 분포 계산
- 💬 주관식 답변 자동 정리
- ✏️ 추출 결과 실시간 수정
- 👀 16:9 PPT 미리보기
- 🖼️ 현장 사진 슬라이드 생성
- 📥 편집 가능한 PPTX 다운로드

---

## 🧩 지원하는 엑셀 구조

다음 형식을 지원합니다.

- 문항이 열로 배치된 설문
- 문항이 행으로 배치된 설문
- 숫자형 1~5점 응답
- 문자형 5점 척도 응답
- 여러 시트가 포함된 엑셀
- 객관식과 주관식이 함께 있는 설문

예시 척도:

```text
매우 그렇다
그렇다
보통이다
그렇지 않다
매우 그렇지 않다
```

---

## 🗂️ 프로젝트 구조

```text
education_satisfaction_report_mvp/
├─ app.py
├─ requirements.txt
├─ README.md
│
├─ assets/
│  ├─ multicampus_logo.png
│  ├─ multicampus_section_background.png
│  └─ 결과보고서_표준양식_참고용.pptx
│
├─ src/
│  ├─ __init__.py
│  ├─ excel_parser.py
│  ├─ models.py
│  ├─ slide_plan.py
│  ├─ preview_renderer.py
│  ├─ ppt_renderer.py
│  ├─ text_utils.py
│  └─ live_input.py
│
└─ docs/
   └─ 변경 이력 문서
```

---

## 🚀 처음 실행하기

### 1. 프로젝트 폴더 열기

VSCode에서 프로젝트 폴더를 엽니다.

```text
education_satisfaction_report_mvp
```

### 2. 가상환경 생성

```bash
py -m venv .venv
```

`py`가 안 되면 아래 명령을 사용합니다.

```bash
python -m venv .venv
```

### 3. 가상환경 활성화

#### Git Bash

```bash
source .venv/Scripts/activate
```

#### PowerShell

```powershell
.venv\Scripts\Activate.ps1
```

PowerShell 실행 정책 오류가 나오면:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

그다음 다시 활성화합니다.

### 4. 패키지 설치

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 5. 앱 실행

```bash
python -m streamlit run app.py
```

브라우저 주소:

```text
http://localhost:8501
```

---

## 🔁 다음 실행부터

가상환경을 새로 만들 필요는 없습니다.

### Git Bash

```bash
source .venv/Scripts/activate
python -m streamlit run app.py
```

### PowerShell

```powershell
.venv\Scripts\Activate.ps1
python -m streamlit run app.py
```

---

## 📝 사용 방법

1. 엑셀 파일을 업로드합니다.
2. 자동 추출 결과를 확인합니다.
3. 필요한 내용을 수정합니다.
4. 슬라이드 미리보기를 확인합니다.
5. PPT 생성 버튼을 누릅니다.
6. 완성된 PPTX를 다운로드합니다.

---

## 🧾 자동 추출 항목

- 과정명
- 교육일정
- 교육방식
- 교육대상
- 교육목표
- 강사명
- 응답자 수
- 객관식 문항
- 주관식 문항
- 문항별 평균
- 응답 분포

엑셀에서 찾지 못한 값은 공란으로 표시됩니다.

---

## 📊 PPT 구성

- 표지
- 목차
- 교육 개요
- 커리큘럼
- 설문 구성
- 만족도 요약
- 객관식 문항별 결과
- 주관식 설문 결과
- 현장 사진
- 감사 페이지

문항 수가 많으면 슬라이드가 자동으로 추가됩니다.

---

## 🖼️ 현장 사진

사진은 최대 6장씩 자동 배치됩니다.

지원 형식:

```text
JPG
JPEG
PNG
WEBP
BMP
GIF
TIFF
ZIP
```

ZIP 내부 폴더명은 슬라이드 소제목으로 사용할 수 있습니다.

---

## ⚠️ 사용 시 참고

- PPT 비율은 16:9입니다.
- 텍스트와 도형은 편집 가능합니다.
- 객관식 점수는 로우데이터 기준으로 다시 계산합니다.
- 주관식의 `없음`, `해당 없음` 등은 선택적으로 제외할 수 있습니다.
- 사진이 없으면 사진 슬라이드는 비워둘 수 있습니다.
- Sourcery 로그인은 필요하지 않습니다.

---

## 🛠️ 자주 발생하는 오류

### `No module named 'src.excel_parser'`

프로젝트 구조를 확인합니다.

```text
education_satisfaction_report_mvp/
├─ app.py
└─ src/
   └─ excel_parser.py
```

`src` 폴더가 한 단계 더 안쪽에 들어가면 안 됩니다.

### 화면이 바뀌지 않을 때

Streamlit을 재시작합니다.

```bash
Ctrl + C
python -m streamlit run app.py
```

브라우저 캐시가 남아 있으면:

```text
Ctrl + F5
```

### 전체 화면 스크롤이 안 될 때

최신 `app.py`가 적용됐는지 확인합니다.

---

## 🧪 현재 버전

```text
v1.7.1 Beta
```

현재 버전에는 다음 내용이 반영되어 있습니다.

- 실시간 입력 반영
- 현재 슬라이드 유지
- 좌우 패널 크기 조절
- PPT 미리보기 UI 개선
- 객관식·주관식 레이아웃 개선
- 실제 PPT 표 높이 정상화
- 멀티캠퍼스 로고 및 배경 적용
- 전체 화면 스크롤 복구

---

## 📌 권장 환경

- Windows 10 이상
- Python 3.10 이상
- VSCode
- Chrome 또는 Edge
- PowerPoint 설치 권장

---

## 📬 문의

오류가 발생하면 아래 내용을 함께 확인합니다.

- 사용한 엑셀 파일
- 오류 메시지
- 터미널 로그
- 문제가 발생한 슬라이드
- 현재 프로젝트 버전
