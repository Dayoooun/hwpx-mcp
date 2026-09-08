# HWPX MCP Server — `@kimdayoun/hwpx-mcp`

[![npm](https://img.shields.io/npm/v/@kimdayoun/hwpx-mcp)](https://www.npmjs.com/package/@kimdayoun/hwpx-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![MCP](https://img.shields.io/badge/MCP-Compatible-blue)](https://modelcontextprotocol.io/)

> **이 저장소가 원본입니다.** 정본: https://github.com/Dayoooun/hwpx-mcp
> npm 배포본은 **`@kimdayoun/hwpx-mcp`** 하나뿐입니다.
> `hwpx-mcp`, `hwpx-mcp-server` 등 스코프 없는 동명 패키지는 이 프로젝트와 무관한 제3자 배포본입니다.

HWP/HWPX 문서를 AI로 읽고 편집할 수 있는 Model Context Protocol (MCP) 서버입니다.

## 특징

- **125개 도구**: 문서의 모든 요소를 프로그래밍 방식으로 제어
- **HWPX 완전 편집**: 텍스트, 테이블, 이미지, 스타일, 머리글/꼬리글 등
- **HWP 읽기 지원**: 레거시 HWP 바이너리 포맷 읽기
- **XML 무결성 보장**: 모든 편집이 HWPML 규격에 맞게 XML에 저장
- **24개 E2E 테스트 통과**: save-reload 검증 완료

## 설치

```bash
npm install -g @kimdayoun/hwpx-mcp
```

설치 없이 바로 실행하려면:

```bash
npx -y @kimdayoun/hwpx-mcp
```

소스에서 빌드하려면:

```bash
git clone https://github.com/Dayoooun/hwpx-mcp.git
cd hwpx-mcp/mcp-server
npm install
npm run build
```

## 설정

### Claude Code (.vscode/mcp.json)

프로젝트 루트에 `.vscode/mcp.json` 파일 생성:

```json
{
  "mcpServers": {
    "hwpx-mcp": {
      "command": "npx",
      "args": ["-y", "@kimdayoun/hwpx-mcp"]
    }
  }
}
```

### Claude Desktop (claude_desktop_config.json)

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "hwpx-mcp": {
      "command": "npx",
      "args": ["-y", "@kimdayoun/hwpx-mcp"]
    }
  }
}
```

## 도구 목록 (125개)

### 문서 관리 (7개)
| 도구 | 설명 |
|------|------|
| `get_tool_guide` | 도구 가이드 조회 |
| `open_document` | 문서 열기 (HWPX/HWP) |
| `create_document` | 새 HWPX 문서 생성 |
| `save_document` | 문서 저장 |
| `close_document` | 문서 닫기 |
| `list_open_documents` | 열린 문서 목록 |
| `get_document_metadata` | 메타데이터 조회 |
| `set_document_metadata` | 메타데이터 설정 |

### 문서 조회 (9개)
| 도구 | 설명 |
|------|------|
| `get_document_text` | 전체 텍스트 |
| `get_document_structure` | 문서 구조 |
| `get_document_outline` | 문서 개요 |
| `get_paragraphs` | 문단 목록 |
| `get_paragraph` | 특정 문단 상세 |
| `get_word_count` | 단어 수 |
| `get_insert_context` | 삽입 컨텍스트 |
| `find_insert_position_after_header` | 헤더 다음 위치 찾기 |
| `find_insert_position_after_table` | 테이블 다음 위치 찾기 |

### 텍스트 편집 (11개)
| 도구 | 설명 |
|------|------|
| `insert_paragraph` | 문단 삽입 |
| `update_paragraph_text` | 문단 텍스트 수정 |
| `update_paragraph_text_preserve_styles` | 스타일 유지 텍스트 수정 |
| `append_text_to_paragraph` | 문단에 텍스트 추가 |
| `delete_paragraph` | 문단 삭제 |
| `copy_paragraph` | 문단 복사 |
| `move_paragraph` | 문단 이동 |
| `search_text` | 텍스트 검색 |
| `replace_text` | 텍스트 치환 |
| `batch_replace` | 일괄 치환 |
| `find_paragraph_by_text` | 텍스트로 문단 찾기 |

### 서식 (스타일) (9개)
| 도구 | 설명 |
|------|------|
| `get_text_style` | 문자 스타일 조회 |
| `set_text_style` | 문자 스타일 설정 (폰트, 크기, 볼드 등) |
| `get_paragraph_style` | 문단 스타일 조회 |
| `set_paragraph_style` | 문단 스타일 설정 (정렬, 줄간격 등) |
| `get_styles` | 스타일 목록 |
| `get_char_shapes` | 문자 모양 전체 |
| `get_para_shapes` | 문단 모양 전체 |
| `apply_style` | 스타일 적용 |
| `get_column_def` | 단 설정 조회 |
| `set_column_def` | 단 설정 |

### 내어쓰기 (8개)
| 도구 | 설명 |
|------|------|
| `get_hanging_indent` | 내어쓰기 조회 |
| `set_hanging_indent` | 내어쓰기 설정 |
| `set_auto_hanging_indent` | 자동 내어쓰기 |
| `remove_hanging_indent` | 내어쓰기 제거 |
| `get_table_cell_hanging_indent` | 셀 내어쓰기 조회 |
| `set_table_cell_hanging_indent` | 셀 내어쓰기 설정 |
| `set_table_cell_auto_hanging_indent` | 셀 자동 내어쓰기 |
| `remove_table_cell_hanging_indent` | 셀 내어쓰기 제거 |

### 테이블 (23개)
| 도구 | 설명 |
|------|------|
| `get_tables` | 테이블 목록 |
| `get_table` | 테이블 상세 |
| `get_table_cell` | 셀 읽기 |
| `get_table_as_csv` | CSV 내보내기 |
| `get_table_map` | 테이블 맵 |
| `get_tables_summary` | 테이블 요약 |
| `get_tables_by_section` | 섹션별 테이블 |
| `find_table_by_header` | 헤더로 테이블 찾기 |
| `find_empty_tables` | 빈 테이블 찾기 |
| `get_element_index_for_table` | 요소 인덱스 |
| `insert_table` | 테이블 삽입 |
| `insert_nested_table` | 중첩 테이블 |
| `delete_table` | 테이블 삭제 |
| `update_table_cell` | 셀 수정 |
| `replace_text_in_cell` | 셀 텍스트 치환 |
| `set_cell_properties` | 셀 속성 |
| `insert_table_row` | 행 추가 |
| `delete_table_row` | 행 삭제 |
| `insert_table_column` | 열 추가 |
| `delete_table_column` | 열 삭제 |
| `merge_cells` | 셀 병합 |
| `split_cell` | 셀 분할 |
| `copy_table` | 테이블 복사 |
| `move_table` | 테이블 이동 |

### 머리글/꼬리글/각주 (8개)
| 도구 | 설명 |
|------|------|
| `get_header` | 머리글 조회 |
| `set_header` | 머리글 설정 |
| `get_footer` | 꼬리글 조회 |
| `set_footer` | 꼬리글 설정 |
| `get_footnotes` | 각주 목록 |
| `insert_footnote` | 각주 삽입 |
| `get_endnotes` | 미주 목록 |
| `insert_endnote` | 미주 삽입 |

### 이미지 (7개)
| 도구 | 설명 |
|------|------|
| `get_images` | 이미지 목록 |
| `insert_image` | 이미지 삽입 |
| `update_image_size` | 이미지 크기 변경 |
| `delete_image` | 이미지 삭제 |
| `render_mermaid` | Mermaid 다이어그램 삽입 |
| `insert_image_in_cell` | 셀에 이미지 삽입 |
| `render_mermaid_in_cell` | 셀에 Mermaid 삽입 |

### 북마크/하이퍼링크 (4개)
| 도구 | 설명 |
|------|------|
| `get_bookmarks` | 북마크 목록 |
| `insert_bookmark` | 북마크 삽입 |
| `get_hyperlinks` | 하이퍼링크 목록 |
| `insert_hyperlink` | 하이퍼링크 삽입 |

### 수식/메모 (6개)
| 도구 | 설명 |
|------|------|
| `get_equations` | 수식 목록 |
| `insert_equation` | 수식 삽입 |
| `get_memos` | 메모 목록 |
| `insert_memo` | 메모 삽입 |
| `delete_memo` | 메모 삭제 |

### 도형 (3개)
| 도구 | 설명 |
|------|------|
| `insert_line` | 선 삽입 |
| `insert_rect` | 사각형 삽입 |
| `insert_ellipse` | 타원 삽입 |

### 섹션/페이지 (5개)
| 도구 | 설명 |
|------|------|
| `get_sections` | 섹션 목록 |
| `insert_section` | 섹션 삽입 |
| `delete_section` | 섹션 삭제 |
| `get_page_settings` | 페이지 설정 조회 |
| `set_page_settings` | 페이지 설정 |

### 실행 취소/다시 실행 (2개)
| 도구 | 설명 |
|------|------|
| `undo` | 실행 취소 |
| `redo` | 다시 실행 |

### 내보내기 (2개)
| 도구 | 설명 |
|------|------|
| `export_to_text` | TXT 내보내기 |
| `export_to_html` | HTML 내보내기 |

### 고급/디버깅 (8개)
| 도구 | 설명 |
|------|------|
| `get_section_xml` | 섹션 XML 조회 |
| `set_section_xml` | 섹션 XML 설정 |
| `get_raw_section_xml` | 원본 섹션 XML |
| `set_raw_section_xml` | 원본 섹션 XML 설정 |
| `analyze_xml` | XML 분석 |
| `repair_xml` | XML 복구 |
| `chunk_document` | 문서 청킹 |
| `invalidate_reading_cache` | 읽기 캐시 무효화 |

### 검색/인덱싱 (5개)
| 도구 | 설명 |
|------|------|
| `search_chunks` | 청크 검색 |
| `get_chunk_context` | 청크 컨텍스트 |
| `extract_toc` | 목차 추출 |
| `build_position_index` | 위치 인덱스 빌드 |
| `get_position_index` | 위치 인덱스 조회 |
| `search_position_index` | 위치 인덱스 검색 |
| `get_chunk_at_offset` | 오프셋 청크 조회 |

## 사용 예시

### 예시 1: 테이블 편집
```
사용자: 사업계획서.hwpx를 열어서 3번째 테이블의 2행 1열을 수정해줘

AI 동작:
1. open_document("사업계획서.hwpx")
2. get_tables() → 테이블 목록 확인
3. get_table(section=0, index=2) → 구조 확인
4. update_table_cell(section=0, table=2, row=1, col=0, text="수정 내용")
5. save_document()
```

### 예시 2: 스타일 복사
```
사용자: 기존 양식의 스타일을 참고해서 새 문단을 추가해줘

AI 동작:
1. get_paragraph_style(sec=0, para=10) → {align: "Justify", lineSpacing: 145}
2. get_text_style(sec=0, para=10) → {fontName: "맑은 고딕", fontSize: 14}
3. insert_paragraph(sec=0, after=10, text="새 내용")
4. set_paragraph_style(sec=0, para=11, align="justify", line_spacing=145)
5. set_text_style(sec=0, para=11, font_name="맑은 고딕", font_size=14)
6. save_document()
```

### 예시 3: 자동 내어쓰기
```
사용자: 테이블 1행 1열에 "1. 항목\n2. 항목" 넣고 자동 내어쓰기 적용해줘

AI 동작:
1. update_table_cell(section=0, table=0, row=0, col=0, text="1. 항목\n2. 항목")
2. set_table_cell_auto_hanging_indent(section=0, table=0, row=0, col=0)
3. save_document()
```

## 테스트

```bash
cd mcp-server
npm test                    # vitest 단위 테스트
npm run test:security       # 빌드 후 실제 MCP 저장 경로 공격·반복 저장·실패 정리 검사
node test-mcp-e2e.mjs       # 16개 기본 E2E 테스트
node test-new-persistence-e2e.mjs  # 8개 persistence E2E 테스트
```

문단 저장 회귀 테스트는 개인 PC의 문서 경로 대신 합성 HWPX를 사용합니다.
여러 `hp:t`로 나뉜 텍스트·빈 문자열·XML 특수문자·동일한 문단 ID/본문의 대상 선택을 검사하며,
중첩 표 삭제는 최상위 표 목록과 셀 내부 중첩 표의 보존을 각각 확인합니다.

## 알려진 제한사항

- **각주/미주/북마크/하이퍼링크 삽입**: 메모리에서만 동작, save 후 XML 미반영 (읽기는 정상)
- **HWP 파일**: 읽기 전용 (편집 불가)
- **secPr 문단 스타일**: 첫 번째 특수 문단에서 스타일 reload 제한
- **파일 접근 범위**: 현재 서버는 작업 폴더 밖의 절대경로·상위경로·디렉터리 심볼릭 링크를 차단하지 않습니다. 신뢰하는 로컬 MCP 클라이언트에서만 사용하고, OS 권한 또는 컨테이너로 접근 범위를 제한하세요. `stdio` 실행을 네트워크 인증·격리로 간주하면 안 됩니다.
- **저장 안전성**: `save_document`는 목적지와 같은 디렉터리 안에 비공개 임시 디렉터리를 만들고 검증 후 rename합니다. 기존 `.tmp` 파일은 사용하지 않으며, `.bak`이 심볼릭 링크 등 일반 파일이 아니면 저장을 거부합니다. 이 조치는 작업 폴더 밖 접근 제한을 대신하지 않습니다.

## 변경 이력

### 0.3.1 (2026-09-08)
- 저장 임시 파일·백업 경로의 링크 공격 방어 및 실패 시 원본 보존
- 여러 `hp:t`에 분리된 텍스트의 저장 누락 수정 및 특수문자 보존
- 개인 Windows 문서에 의존하던 테스트를 합성 HWPX 회귀 테스트로 교체
- 의존성 갱신 및 기존 절대·상대경로 사용법의 호환성 검사 추가
- 작업 폴더 제한은 추가하지 않았으며, 수식 저장 누락은 이 릴리스의 수정 범위에 포함하지 않음
### 0.3.0 (2026-01-28)
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

### 0.2.0
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

### 0.1.1
- **버그 수정**: `update_table_cell` 후 `save_document` 시 빈 셀 변경사항이 저장되지 않던 문제 수정
  - Self-closing XML run 태그 (`<hp:run ... />`) 처리 지원 추가
  - ID 기반 테이블 매칭으로 정확한 XML 업데이트 구현
  - 원본 XML 구조를 최대한 보존하면서 텍스트만 수정

### 0.1.0
- 최초 릴리스

## 라이선스

MIT
