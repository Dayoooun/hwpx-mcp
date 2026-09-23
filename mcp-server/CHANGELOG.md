# Changelog

`@kimdayoun/hwpx-mcp` 의 주요 변경 사항을 기록합니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르고,
버전은 [유의적 버전](https://semver.org/lang/ko/) 을 따릅니다.

## [0.3.2] - 2026-09-22

외부 사용 리뷰에서 보고된 8건을 MCP stdio 클라이언트로 직접 재현한 뒤 수정했습니다.
리뷰 대응 내역은 [#10](https://github.com/Dayoooun/hwpx-mcp/pull/10) 에 정리돼 있습니다.

### Fixed

- **표 행·열 삽입이 저장본을 깨뜨리던 문제.** 복제한 행·셀의 본문을 비우는 정규식
  `<(hp|hs):t([^>]*)>` 에 태그명 뒤 경계가 없어 `<hp:tc>` 를 `<hp:t>` 로 삼켰습니다.
  `<hp:subList>` 이하가 통째로 사라지고 닫는 태그만 남아 한/글이 파일을 열지 못했습니다.
  `insert_table_row` · `insert_table_column` 양쪽 해당.
- **`copy_paragraph` 직후 텍스트를 교체하면 원본까지 바뀌던 문제.** 저장 파이프라인이
  텍스트 갱신을 문단 복제보다 먼저 적용해, 복제 후 인덱스로 지정한 대상이 복제 전 XML 의
  원본 문단으로 해석됐습니다. 복제·이동을 모든 텍스트 갱신 앞으로 옮겼습니다.
  중간에 `save_document` 를 끼우던 우회가 더 이상 필요 없습니다.
- **`create_document` · `save_document` 의 경로 무시.** 새 문서가 실제 위치 없이
  `new-document.hwpx` 라는 상대 파일명을 들고 있어 저장이 서버 프로세스 cwd 로 갔습니다.
  목적지가 없으면 조용히 cwd 로 가지 않고 명시적으로 실패합니다.
- **병합으로 덮인 셀에 쓰면 성공이라 답하고 저장 시 사라지던 문제.** HWPX 는 덮인 위치에
  `<hp:tc>` 를 두지 않아 직렬화 대상이 없습니다. 이제 어느 마스터 셀에 써야 하는지
  알려주며 거부합니다.
- **`get_tool_guide` 의 `topic` 이 무시되던 문제.** 모르는 값에 조용히 전체 참조를
  반환하던 것을 가능한 값 목록을 담은 오류로 바꿨습니다.
- **복제 문단·행·열의 `linesegarray` 미초기화.** 원본의 고정 줄 배치를 물려받아 긴 글을
  넣으면 글자가 겹쳤습니다. `get_section_xml` → 문자열 치환 → `set_section_xml` 수작업이
  필요 없어집니다.

### Changed

- **필수 인자 누락 시 원인을 밝힙니다.** `section_index` 를 빠뜨리면
  `Failed to insert paragraph` 가 떠 문서 손상처럼 보였습니다. 스키마의 `required` 를
  디스패처에서 검사해 어떤 인자가 빠졌는지 이름으로 답합니다.
  예: `Missing required arguments for insert_paragraph: section_index, after_index`
- **`save_document` 응답에 절대경로를 돌려줍니다.** `path` 와 `backup_path` 가 추가됐고,
  `output_path` 외에 `file_path` 도 받습니다.
- **`batch_fill_table` 이 실패한 셀을 보고하고 나머지를 계속 채웁니다.** 반환값에 `failed`
  목록(좌표·값·이유)이 추가됐습니다. 이전에는 덮인 셀 하나가 배치 전체를 중단시켰습니다.

### Added

- **`create_document({file_path})`** — 목적지를 미리 지정하면 이후 `save_document` 가
  인자 없이도 그 경로에 씁니다.
- **`paraPrIDRef` · `charPrIDRef` 노출** — `get_paragraph` · `get_paragraphs` 가 스타일
  값과 함께 원본 숫자 ID 를 돌려줍니다. XML 을 직접 조립할 때 `section0.xml` 을 정규식으로
  파지 않아도 됩니다.

### 알려진 제한

- `updateTableCell` 은 셀마다 문서 전체를 직렬화해 undo 스택에 쌓습니다. 실측상 100×30 표
  전수 쓰기 1,725ms 중 1,501ms(87%)가 여기서 발생합니다. 구조 변경이라 이 릴리스의 수정
  범위에 포함하지 않았습니다.

## [0.3.1] - 2026-09-08

- 저장 임시 파일·백업 경로의 링크 공격 방어 및 실패 시 원본 보존
- 여러 `hp:t`에 분리된 텍스트의 저장 누락 수정 및 특수문자 보존
- 개인 Windows 문서에 의존하던 테스트를 합성 HWPX 회귀 테스트로 교체
- 의존성 갱신 및 기존 절대·상대경로 사용법의 호환성 검사 추가
- 작업 폴더 제한은 추가하지 않았으며, 수식 저장 누락은 이 릴리스의 수정 범위에 포함하지 않음

## 0.3.0 - 2026-01-28

- **대규모 XML persistence 수정**: 8개 조작에 대한 save-reload 지원 추가
  - `insertTableRow` / `deleteTableRow`
  - `insertTableColumn` / `deleteTableColumn`
  - `copyParagraph` / `moveParagraph` (+ 같은 섹션 인덱스 버그 수정)
  - `setHeader` / `setFooter`
- **depth-aware element indexing**: 중첩 태그 내부 요소를 무시하는 안전한 파싱
- **undo/redo 안전성**: pending 배열 초기화로 메모리/XML 비동기화 방지
- **replaceText 수정**: XML entity 불일치 해결
- **mergeCells 수정**: indexOf 모호성 해결
- **테스트**: 24개 E2E 테스트 (16 기본 + 8 persistence)

## 0.2.0

- **신규 기능**: 테이블 셀 내 내어쓰기(Hanging Indent) 자동 적용
  - `update_table_cell` 시 마커(○, 1., 가., (1) 등) 감지하여 자동 내어쓰기
  - 멀티라인 텍스트의 각 줄에 독립적으로 내어쓰기 적용
  - `set_table_cell_hanging_indent`, `get_table_cell_hanging_indent` 도구 추가
- **버그 수정**: 병렬 테이블 업데이트 시 XML 손상 문제 해결
  - 문서별 Lock 추가로 병렬 요청 직렬화 (race condition 방지)
  - `findTableCellInXml()` 중첩 테이블 처리 개선 (balanced bracket 매칭)
  - 여러 테이블 동시 수정 후 저장 시 "Broken tag structure" 오류 수정
- **버그 수정**: 여러 테이블에 내어쓰기 적용 시 stale position 문제 해결
  - 테이블 인덱스 내림차순 처리로 위치 변경 영향 방지
  - 각 테이블 처리 시 위치 정보 재계산
- **테스트 강화**: Red Team 스트레스 테스트 추가 (238개 테스트)
  - 50~200개 테이블 대량 수정 테스트
  - 중첩 테이블 + 내어쓰기 + 이미지 복합 테스트
  - 병렬 업데이트 시나리오 테스트

## 0.1.1

- **버그 수정**: `update_table_cell` 후 `save_document` 시 빈 셀 변경사항이 저장되지 않던 문제 수정
  - Self-closing XML run 태그 (`<hp:run ... />`) 처리 지원 추가
  - ID 기반 테이블 매칭으로 정확한 XML 업데이트 구현
  - 원본 XML 구조를 최대한 보존하면서 텍스트만 수정

## 0.1.0

- 최초 릴리스

[0.3.2]: https://github.com/Dayoooun/hwpx-mcp/compare/v0.3.1...v0.3.2
[0.3.1]: https://github.com/Dayoooun/hwpx-mcp/releases/tag/v0.3.1
