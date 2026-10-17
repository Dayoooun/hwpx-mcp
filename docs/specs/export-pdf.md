# export_to_pdf — 맥·윈도우 어디서나 한컴 엔진으로 PDF

## 왜

PDF 는 지금까지 윈도우 한/글 COM 으로만 뽑았다. 맥 사용자는 윈도우 PC 를 따로 두고
SSH 로 보내야 했다. REA 분석으로 한/글 for Mac 의 Quick Look 미리보기가 쓰는 한컴 렌더
엔진(`libSimpleEngineLib`)을 화면 없이 부를 수 있음을 확인했다(문서당 약 0.5초, 40개
대조 쪽수 39/40). 같은 도구 하나가 실행 환경을 보고 엔진을 고른다.

## 동작

- 입력: `file_path`(HWP/HWPX) 또는 `doc_id`(열린 문서 — 고친 내용 그대로). 출력: `output_path`.
- `engine`: `auto`(기본) · `hancom-mac` · `hancom-windows`.
  - `auto` 는 macOS → `hancom-mac`, Windows → `hancom-windows`, 그 밖 → 거절.
- 맥: 패키지에 든 `hwp2pdf`(universal) 가 설치된 한/글 for Mac 의 엔진을 불러 그린다. 화면·포커스를 쓰지 않는다.
- 윈도우: PowerShell 로 `HWPFrame.HwpObject` 를 띄워 `Open` → `SaveAs(PDF)` → `Quit`. 경로는 환경변수로 넘긴다.
- PDF 는 출력 폴더 안의 임시 이름으로 쓴 뒤 이름을 바꾼다. 실패하면 출력 경로에 아무것도 남기지 않고, 이미 있던 파일도 그대로 둔다.
- 결과: `{ output_path, pages, engine }`.

## Acceptance

- AC-1: `auto` 는 darwin 에서 hancom-mac, win32 에서 hancom-windows 를 고르고, 그 밖의 플랫폼은 이유와 필요한 설치를 적은 오류를 낸다.
- AC-2: 명시한 엔진이 현재 플랫폼에서 못 도는 것이면(예: 리눅스에서 hancom-mac) 오류로 거절한다.
- AC-3: 윈도우 스크립트는 입력·출력 경로를 스크립트 본문에 넣지 않는다(따옴표·`$` 가 든 경로에도 주입이 없다).
- AC-4: 엔진이 없는 환경에서 `export_to_pdf` 는 isError 로 오고 출력 파일을 만들지 않는다.
- AC-5: (맥 + 한/글) `file_path` 로 변환하면 PDF 가 생기고 쪽 수와 본문 글자가 들어 있다.
- AC-6: (맥 + 한/글) `doc_id` 로 변환하면 저장 전에 고친 글이 PDF 에 나온다.
- AC-7: (맥 + 한/글) 열 수 없는 문서는 isError 이고, 이미 있던 출력 파일은 바이트 그대로다.
- AC-8: [manual] 윈도우 한/글 2024 에서 같은 도구 호출로 PDF 가 나온다(노트북 실측).
