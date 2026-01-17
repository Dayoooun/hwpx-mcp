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

## 🔌 MCP Tools

| Tool | 설명 |
|------|------|
| `open_document` | HWPX 문서 열기 |
| `get_document_text` | 전체 텍스트 추출 |
| `get_tables` | 테이블 목록 조회 |
| `get_table` | 특정 테이블 상세 정보 |
| `update_table_cell` | 테이블 셀 내용 수정 |
| `search_text` | 텍스트 검색 |
| `replace_text` | 텍스트 치환 |
| `save_document` | 문서 저장 |
| `close_document` | 문서 닫기 |

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

// 저장
await mcp.save_document({ doc_id: "..." })
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
