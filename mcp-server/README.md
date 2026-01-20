# HWPX MCP Server

HWP/HWPX 문서를 읽고 편집할 수 있는 Model Context Protocol (MCP) 서버입니다.

## 기능

- **문서 열기/저장**: HWPX (편집 가능), HWP (읽기 전용)
- **텍스트 편집**: 단락 추가/수정/삭제, 텍스트 검색/치환
- **테이블 편집**: 테이블 생성, 셀 편집, 행/열 추가/삭제
- **서식 지정**: 글꼴, 크기, 색상, 정렬 등
- **메타데이터**: 제목, 작성자, 날짜 등 편집
- **내보내기**: TXT, HTML 형식으로 내보내기

## 설치

```bash
cd mcp-server
npm install
npm run build
```

## Claude Code에서 사용하기

### 방법 1: 프로젝트별 설정 (.vscode/mcp.json)

프로젝트 루트에 `.vscode/mcp.json` 파일을 생성합니다:

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["C:\\Users\\mjyoo\\hwp-extension\\mcp-server\\dist\\index.js"]
    }
  }
}
```

### 방법 2: 전역 설정

Windows: `%APPDATA%\Code\User\globalStorage\anthropic.claude-code\settings\mcp.json`
macOS: `~/Library/Application Support/Code/User/globalStorage/anthropic.claude-code/settings/mcp.json`
Linux: `~/.config/Code/User/globalStorage/anthropic.claude-code/settings/mcp.json`

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["/path/to/mcp-server/dist/index.js"]
    }
  }
}
```

## Claude Desktop에서 사용하기

`claude_desktop_config.json` 파일에 추가:

Windows: `%APPDATA%\Claude\claude_desktop_config.json`
macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["C:\\Users\\mjyoo\\hwp-extension\\mcp-server\\dist\\index.js"]
    }
  }
}
```

## 사용 가능한 도구

| 도구 | 설명 |
|------|------|
| `open_document` | 문서 열기 |
| `create_document` | 새 문서 생성 |
| `save_document` | 문서 저장 |
| `close_document` | 문서 닫기 |
| `get_document_text` | 전체 텍스트 가져오기 |
| `get_paragraphs` | 단락 목록 가져오기 |
| `insert_paragraph` | 단락 삽입 |
| `update_paragraph_text` | 단락 텍스트 수정 |
| `delete_paragraph` | 단락 삭제 |
| `search_text` | 텍스트 검색 |
| `replace_text` | 텍스트 치환 |
| `get_tables` | 테이블 목록 가져오기 |
| `insert_table` | 테이블 삽입 |
| `update_table_cell` | 테이블 셀 수정 |
| `set_text_style` | 텍스트 스타일 적용 |
| `set_paragraph_style` | 단락 스타일 적용 |
| `export_to_text` | TXT로 내보내기 |
| `export_to_html` | HTML로 내보내기 |

## 예시

```
사용자: test.hwpx 파일을 열어서 내용을 보여줘
AI: [open_document 도구 사용] → [get_document_text 도구 사용]

사용자: 첫 번째 단락에 "안녕하세요"를 추가해줘
AI: [insert_paragraph 도구 사용]

사용자: 2x3 테이블을 추가해줘
AI: [insert_table 도구 사용]
```

## 변경 이력

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
