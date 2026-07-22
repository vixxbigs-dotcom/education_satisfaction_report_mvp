# V2.8 Beta 변경사항

## 웹 글꼴 선택

상단 `출력 설정`에서 웹 미리보기와 실제 PPT에 사용할 글꼴을 선택할 수 있습니다.

기본 선택지:

- 나눔스퀘어
- 나눔고딕
- 맑은 고딕

선택을 바꾸면 현재 슬라이드의 웹 미리보기가 즉시 변경됩니다. 이미 생성한 PPT 바이트는 무효화되므로 `PPT 생성 / 최신 내용 반영`을 다시 눌러야 합니다.

## 폰트 파일 위치

폰트 파일은 ZIP에 포함하지 않습니다. 아래 위치에 직접 넣습니다.

```text
assets/fonts/
├─ font_config.json
├─ NanumSquareR.ttf
├─ NanumSquareB.ttf
├─ NanumGothic.ttf
└─ NanumGothicBold.ttf
```

맑은 고딕은 Windows 기본 글꼴이므로 파일이 필요하지 않습니다.

## 반영 범위

- 웹 PPT 미리보기
- PPT 텍스트 상자
- PPT 표
- PPT 차트 축, 범례, 데이터 라벨
- PPT 주관식 응답

실제 PPT에는 선택한 Windows 글꼴명이 지정되며, PPT를 여는 PC에도 해당 글꼴이 설치되어 있어야 같은 모양으로 표시됩니다.
