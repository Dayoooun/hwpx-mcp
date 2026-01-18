# HWPX MCP Server - Enhanced Edition

[![GitHub](https://img.shields.io/badge/GitHub-Dayoooun%2Fhwp--extension-blue?logo=github)](https://github.com/Dayoooun/hwp-extension)
[![Fork](https://img.shields.io/badge/Forked%20from-mjyoo2%2Fhwp--extension-gray?logo=github)](https://github.com/mjyoo2/hwp-extension)

> 🚀 **Original 프로젝트를 Fork하여 안정성과 기능을 대폭 개선한 버전입니다.**

AI 도구(Claude 등)와 연동하여 한글(HWPX) 문서를 자동으로 편집할 수 있는 MCP(Model Context Protocol) 서버입니다.

---

## ✨ Enhanced Features (개선된 기능)

원본 프로젝트 대비 다음과 같은 **핵심 문제들을 해결**했습니다:

### 🔧 Critical Bug Fixes

| 문제 | 원본 상태 | 개선 후 |
|------|----------|---------|
| **테이블 저장 실패** | 셀 수정 후 저장해도 변경사항 사라짐 | ✅ 완벽하게 저장됨 |
| **텍스트 겹침 현상** | 저장 후 한글에서 열면 글자가 겹쳐 표시 | ✅ 정상 표시 |
| **파일 손상** | 저장 시 가끔 파일이 손상됨 | ✅ 원자적 쓰기로 100% 안전 |
| **다중 셀 손상** | 같은 행에 여러 셀 수정 시 XML 손상 | ✅ 인덱스 관리로 안전 |
| **자간/줄간격 손실** | 저장 후 스타일 정보 유실 | ✅ 모든 스타일 보존 |

### 🛠 Technical Improvements

1. **Atomic File Writing (원자적 파일 쓰기)**
   - 임시 파일 → ZIP 검증 → 원자적 이동
   - 저장 중 오류 발생해도 원본 파일 보호

2. **Smart Lineseg Reset (스마트 줄 레이아웃 초기화)**
   - 텍스트 수정 시 `lineseg` 자동 초기화
   - 한글 프로그램이 열 때 자동으로 줄바꿈 재계산
   - 텍스트 겹침 현상 완전 해결

3. **Depth-based XML Parsing (깊이 기반 XML 파싱)**
   - 기존 lazy regex의 중첩 구조 오인식 문제 해결
   - 복잡한 테이블(중첩 테이블, subList 등) 완벽 지원

4. **Complete Style Preservation (스타일 완전 보존)**
   - `charPr`, `spacing` 등 원본 스타일 100% 유지
   - 불완전한 직렬화 로직 제거로 데이터 무결성 보장

5. **Safe Multi-Cell Updates (안전한 다중 셀 업데이트)**
   - 같은 행(row)의 여러 셀을 동시에 수정해도 안전
   - 행별 그룹화 + 역순 처리로 인덱스 손상 방지

---

## 📦 Installation

### MCP 서버 설치

```bash
git clone https://github.com/Dayoooun/hwp-extension.git
cd hwp-extension/mcp-server
npm install
npm run build
```

### Claude Code 연동

`~/.claude/claude_desktop_config.json` 또는 `.vscode/mcp.json`에 추가:

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["/path/to/hwp-extension/mcp-server/dist/index.js"]
    }
  }
}
```

---

## 🔌 MCP Tools (77개)

### 📁 문서 관리 (Document Management) - 5개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `create_document` | 새 빈 HWPX 문서 생성 | `title?`, `creator?` |
| `open_document` | HWPX 문서 열기 | `file_path` |
| `close_document` | 열린 문서 닫기 | `doc_id` |
| `save_document` | 문서 저장 (백업/무결성 검증 지원) | `doc_id`, `output_path?`, `create_backup?`, `verify_integrity?` |
| `list_open_documents` | 현재 열린 문서 목록 조회 | - |

### 📄 문서 정보 (Document Info) - 5개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_document_text` | 문서 전체 텍스트 추출 | `doc_id` |
| `get_document_structure` | 문서 구조 조회 (섹션/단락/테이블/이미지 수) | `doc_id` |
| `get_document_metadata` | 메타데이터 조회 (제목, 저자, 날짜 등) | `doc_id` |
| `set_document_metadata` | 메타데이터 수정 | `doc_id`, `title?`, `creator?`, `subject?`, `description?` |
| `get_word_count` | 글자수/단어수 통계 | `doc_id` |

### 📝 단락 (Paragraphs) - 8개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_paragraphs` | 단락 목록 조회 (텍스트/스타일 포함) | `doc_id`, `section_index?` |
| `get_paragraph` | 특정 단락 상세 정보 | `doc_id`, `section_index`, `paragraph_index` |
| `insert_paragraph` | 새 단락 삽입 | `doc_id`, `section_index`, `after_index`, `text` |
| `delete_paragraph` | 단락 삭제 | `doc_id`, `section_index`, `paragraph_index` |
| `update_paragraph_text` | 단락 텍스트 내용 수정 | `doc_id`, `section_index`, `paragraph_index`, `text`, `run_index?` |
| `append_text_to_paragraph` | 기존 단락에 텍스트 추가 | `doc_id`, `section_index`, `paragraph_index`, `text` |
| `copy_paragraph` | 단락을 다른 위치로 복사 | `doc_id`, `source_section`, `source_paragraph`, `target_section`, `target_after` |
| `move_paragraph` | 단락을 다른 위치로 이동 | `doc_id`, `source_section`, `source_paragraph`, `target_section`, `target_after` |

### 🎨 텍스트 스타일 (Text Styling) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_text_style` | 글자 서식 조회 (폰트/크기/색상 등) | `doc_id`, `section_index`, `paragraph_index`, `run_index?` |
| `set_text_style` | 글자 서식 설정 | `doc_id`, `section_index`, `paragraph_index`, `bold?`, `italic?`, `underline?`, `strikethrough?`, `font_name?`, `font_size?`, `font_color?`, `background_color?` |
| `get_paragraph_style` | 문단 서식 조회 (정렬/줄간격/여백 등) | `doc_id`, `section_index`, `paragraph_index` |
| `set_paragraph_style` | 문단 서식 설정 | `doc_id`, `section_index`, `paragraph_index`, `align?`, `line_spacing?`, `margin_*?`, `first_line_indent?` |

### 🔍 검색/치환 (Search & Replace) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `search_text` | 문서 내 텍스트 검색 (정규식 지원, **테이블 셀 포함**) | `doc_id`, `query`, `case_sensitive?`, `regex?`, `include_tables?` |
| `replace_text` | 텍스트 찾아 바꾸기 | `doc_id`, `old_text`, `new_text`, `case_sensitive?`, `regex?`, `replace_all?` |
| `replace_text_in_cell` | **특정 테이블 셀 내 텍스트 치환** | `doc_id`, `section_index`, `table_index`, `row`, `col`, `old_text`, `new_text` |
| `batch_replace` | 여러 텍스트 일괄 치환 | `doc_id`, `replacements[]` (old_text, new_text 쌍 배열) |

### 📊 테이블 (Tables) - 12개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_tables` | 문서 내 모든 테이블 목록 | `doc_id` |
| `get_table` | 특정 테이블 전체 데이터 조회 | `doc_id`, `section_index`, `table_index` |
| `get_table_cell` | 특정 셀 내용 조회 | `doc_id`, `section_index`, `table_index`, `row`, `col` |
| `update_table_cell` | 셀 내용 수정 (스타일 보존) | `doc_id`, `section_index`, `table_index`, `row`, `col`, `text`, `char_shape_id?` |
| `set_cell_properties` | 셀 속성 설정 (크기/배경색/정렬) | `doc_id`, `section_index`, `table_index`, `row`, `col`, `width?`, `height?`, `background_color?`, `vertical_align?` |
| `insert_table` | 새 테이블 삽입 | `doc_id`, `section_index`, `after_index`, `rows`, `cols`, `width?` |
| `insert_table_row` | 테이블에 행 삽입 | `doc_id`, `section_index`, `table_index`, `after_row`, `cell_texts?` |
| `delete_table_row` | 테이블에서 행 삭제 | `doc_id`, `section_index`, `table_index`, `row_index` |
| `insert_table_column` | 테이블에 열 삽입 | `doc_id`, `section_index`, `table_index`, `after_col` |
| `delete_table_column` | 테이블에서 열 삭제 | `doc_id`, `section_index`, `table_index`, `col_index` |
| `insert_nested_table` | **셀 안에 중첩 테이블 삽입 (표 안에 표)** | `doc_id`, `section_index`, `parent_table_index`, `row`, `col`, `nested_rows`, `nested_cols`, `data?` |
| `get_table_as_csv` | 테이블을 CSV 형식으로 추출 | `doc_id`, `section_index`, `table_index`, `delimiter?` |

### 📐 페이지 설정 (Page Settings) - 2개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_page_settings` | 페이지 설정 조회 (용지 크기/여백) | `doc_id`, `section_index?` |
| `set_page_settings` | 페이지 설정 변경 | `doc_id`, `section_index?`, `width?`, `height?`, `margin_*?`, `orientation?` |

### 🖼️ 이미지 (Images) - 5개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_images` | 문서 내 모든 이미지 목록 | `doc_id` |
| `insert_image` | 이미지 파일 삽입 (BinData 자동 등록) | `doc_id`, `section_index`, `after_index`, `image_path`, `width?`, `height?` |
| `update_image_size` | 기존 이미지 크기 변경 | `doc_id`, `section_index`, `image_index`, `width`, `height` |
| `delete_image` | 이미지 삭제 | `doc_id`, `section_index`, `image_index` |
| `render_mermaid` | **Mermaid 다이어그램을 이미지로 삽입** | `doc_id`, `mermaid_code`, `after_index`, `section_index?`, `width?`, `height?`, `theme?`, `background_color?` |

### ✏️ 도형 (Shapes) - 3개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `insert_line` | 선 도형 삽입 | `doc_id`, `section_index`, `after_index`, `x1`, `y1`, `x2`, `y2`, `stroke_color?`, `stroke_width?` |
| `insert_rect` | 사각형 도형 삽입 | `doc_id`, `section_index`, `after_index`, `x`, `y`, `width`, `height`, `fill_color?`, `stroke_color?` |
| `insert_ellipse` | 타원 도형 삽입 | `doc_id`, `section_index`, `after_index`, `cx`, `cy`, `rx`, `ry`, `fill_color?`, `stroke_color?` |

### 📑 머리글/바닥글 (Header/Footer) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_header` | 머리글 내용 조회 | `doc_id`, `section_index?` |
| `set_header` | 머리글 설정 | `doc_id`, `text`, `section_index?`, `apply_page_type?` (both/even/odd) |
| `get_footer` | 바닥글 내용 조회 | `doc_id`, `section_index?` |
| `set_footer` | 바닥글 설정 | `doc_id`, `text`, `section_index?`, `apply_page_type?` (both/even/odd) |

### 📌 각주/미주 (Footnotes/Endnotes) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_footnotes` | 문서 내 모든 각주 목록 | `doc_id` |
| `insert_footnote` | 특정 위치에 각주 삽입 | `doc_id`, `section_index`, `paragraph_index`, `text` |
| `get_endnotes` | 문서 내 모든 미주 목록 | `doc_id` |
| `insert_endnote` | 특정 위치에 미주 삽입 | `doc_id`, `section_index`, `paragraph_index`, `text` |

### 🔗 북마크/하이퍼링크 (Bookmarks/Hyperlinks) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_bookmarks` | 문서 내 모든 북마크 목록 | `doc_id` |
| `insert_bookmark` | 특정 위치에 북마크 삽입 | `doc_id`, `section_index`, `paragraph_index`, `name` |
| `get_hyperlinks` | 문서 내 모든 하이퍼링크 목록 | `doc_id` |
| `insert_hyperlink` | 하이퍼링크 삽입 | `doc_id`, `section_index`, `paragraph_index`, `url`, `text` |

### ➗ 수식 (Equations) - 2개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_equations` | 문서 내 모든 수식 목록 | `doc_id` |
| `insert_equation` | 수식 삽입 (HWP 수식 스크립트 형식) | `doc_id`, `section_index`, `after_index`, `script` |

### 💬 메모 (Memos/Comments) - 3개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_memos` | 문서 내 모든 메모/주석 목록 | `doc_id` |
| `insert_memo` | 메모/주석 삽입 | `doc_id`, `section_index`, `paragraph_index`, `content`, `author?` |
| `delete_memo` | 메모/주석 삭제 | `doc_id`, `memo_id` |

### 📚 섹션 (Sections) - 5개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_sections` | 문서 내 모든 섹션 목록 | `doc_id` |
| `insert_section` | 새 섹션 삽입 | `doc_id`, `after_index` |
| `delete_section` | 섹션 삭제 | `doc_id`, `section_index` |
| `get_section_xml` | **섹션 Raw XML 조회 (AI 문서 조작용)** | `doc_id`, `section_index?` |
| `set_section_xml` | **섹션 Raw XML 교체 (HWPML 형식 필수)** | `doc_id`, `xml`, `section_index?`, `validate?` |

### 🎭 스타일 정의 (Style Definitions) - 4개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_styles` | 문서에 정의된 스타일 목록 | `doc_id` |
| `get_char_shapes` | 글자 모양(CharShape) 정의 목록 | `doc_id` |
| `get_para_shapes` | 문단 모양(ParaShape) 정의 목록 | `doc_id` |
| `apply_style` | 단락에 스타일 적용 | `doc_id`, `section_index`, `paragraph_index`, `style_id` |

### 📰 단 설정 (Column Layout) - 2개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `get_column_def` | 단(Column) 설정 조회 | `doc_id`, `section_index?` |
| `set_column_def` | 단 설정 변경 (다단 편집) | `doc_id`, `count`, `section_index?`, `type?`, `same_size?`, `gap?` |

### 📤 내보내기 (Export) - 2개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `export_to_text` | 문서를 텍스트 파일로 내보내기 | `doc_id`, `output_path` |
| `export_to_html` | 문서를 HTML 파일로 내보내기 | `doc_id`, `output_path` |

### ↩️ 실행 취소 (Undo/Redo) - 2개

| Tool | 설명 | 주요 파라미터 |
|------|------|--------------|
| `undo` | 마지막 변경 실행 취소 | `doc_id` |
| `redo` | 실행 취소한 변경 다시 실행 | `doc_id` |

---

### 사용 예시

```typescript
// 문서 열기
await mcp.open_document({ file_path: "report.hwpx" })

// 테이블 셀 수정
await mcp.update_table_cell({
  doc_id: "...",
  section_index: 0,
  table_index: 0,
  row: 0,
  col: 1,
  text: "수정된 내용"
})

// 중첩 테이블 삽입 (표 안에 표)
await mcp.insert_nested_table({
  doc_id: "...",
  section_index: 0,
  parent_table_index: 0,
  row: 1,
  col: 2,
  nested_rows: 2,
  nested_cols: 2,
  data: [["A1", "A2"], ["B1", "B2"]]
})

// Mermaid 다이어그램 삽입
await mcp.render_mermaid({
  doc_id: "...",
  mermaid_code: "graph TD; A-->B; B-->C;",
  after_index: 0,
  theme: "default"
})

// 저장
await mcp.save_document({ doc_id: "..." })
```

#### 테이블 셀 검색/치환 워크플로우

동일한 텍스트가 여러 곳에 있을 때 **특정 위치**만 수정하는 방법:

```typescript
// 1. 테이블 포함 검색으로 위치 파악
const results = await mcp.search_text({
  doc_id: "...",
  query: "수정대상",
  include_tables: true  // 테이블 셀 포함
})
// 결과: [{ text: "수정대상", location: { type: "table", tableIndex: 2, row: 3, col: 1 } }, ...]

// 2. 원하는 위치의 셀만 정밀 치환
await mcp.replace_text_in_cell({
  doc_id: "...",
  section_index: 0,
  table_index: 2,  // 검색 결과에서 확인한 위치
  row: 3,
  col: 1,
  old_text: "수정대상",
  new_text: "새로운내용"
})
```

---

## 📋 Supported Format

| 포맷 | 확장자 | 읽기 | 쓰기 |
|------|--------|:----:|:----:|
| HWPX | .hwpx | ✅ | ✅ |
| HWP | .hwp | ❌ | ❌ |

> **Note**: HWP(바이너리) 파일은 지원하지 않습니다. 한컴오피스에서 HWPX로 변환 후 사용하세요.

---

## 📝 Release Notes

### v0.4.0 (Enhanced Search & Diagram Support)
- 🆕 **New Feature**: `search_text`에 `include_tables` 옵션 추가 - 테이블 셀 내 텍스트도 검색
- 🆕 **New Feature**: `replace_text_in_cell` - 특정 테이블 셀 내 텍스트만 정밀 치환
- 🆕 **New Feature**: `render_mermaid` - Mermaid 다이어그램을 이미지로 문서에 삽입
  - Flowchart, Sequence, Class Diagram 등 모든 Mermaid 문법 지원
  - 테마 선택 가능 (default, dark, forest, neutral)
- 🆕 **New Feature**: `get_section_xml` / `set_section_xml` - 섹션 Raw XML 직접 조작
  - AI 기반 고급 문서 편집 시나리오 지원
- 🔧 **Improvement**: `insert_image` 완전 개선
  - BinData 폴더에 이미지 자동 저장
  - content.hpf 매니페스트 자동 등록
  - 파일 손상 없이 이미지 삽입 보장

### v0.3.0 (Nested Table Support)
- 🆕 **New Feature**: `insert_nested_table` - 테이블 셀 안에 중첩 테이블 삽입 기능
  - 부모 테이블의 특정 셀에 새 테이블을 삽입
  - 초기 데이터 지정 가능 (2D 배열)
  - HWPX 표준 구조(`treatAsChar`, `hp:subList`) 완벽 준수
- 🔧 **Improvement**: charSpacing 파싱 개선 (속성 순서 무관하게 처리)

### v0.2.1 (Critical Fix)
- 🔥 **Critical Fix**: 같은 행에 여러 셀 동시 수정 시 파일 손상 문제 완전 해결
  - 행(row)별 업데이트 그룹화로 인덱스 불일치 방지
  - 역순(descending) 처리로 안전한 XML 수정 보장

### v0.2.0 (Enhanced Edition)
- 🔥 **Major Fix**: 텍스트 수정 시 lineseg 자동 초기화로 겹침 현상 완전 해결
- 🔧 **Bug Fix**: 중첩 테이블 구조에서 XML 요소 경계 오인식 문제 수정
- 🛡️ **Stability**: 원자적 파일 쓰기로 파일 손상 방지
- 📦 **Preservation**: 원본 charPr/spacing 스타일 완전 보존

### v0.1.0 (Original)
- 최초 릴리스 (mjyoo2/hwp-extension)

---

## 🙏 Credits

- Original Project: [mjyoo2/hwp-extension](https://github.com/mjyoo2/hwp-extension)
- Enhanced by: [Dayoooun](https://github.com/Dayoooun)

---

## 📄 License

MIT

---

## 🤝 Contributing

버그 리포트 및 기능 요청: [GitHub Issues](https://github.com/Dayoooun/hwp-extension/issues)
