<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

HWP/HWPX Editor - A VSCode extension for editing Hangul Word Processor documents. Supports both modern HWPX (XML-based, editable) and legacy HWP (binary, read-only) formats.

## Build Commands

**VSCode Extension (root directory):**
```bash
pnpm compile           # Build extension with esbuild
pnpm watch             # Watch mode for development
pnpm package           # Package as VSIX (requires vsce)
```

**MCP Server (mcp-server directory):**
```bash
cd mcp-server
pnpm build             # Compile TypeScript
pnpm start             # Run the server
```

**Testing (MCP Server only):**
```bash
cd mcp-server
pnpm test              # Run all tests (vitest)
pnpm test:watch        # Watch mode for tests
pnpm vitest run <file> # Run specific test file (e.g., pnpm vitest run src/HwpxDocument.test.ts)
```

**Development:** Press F5 in VSCode to launch the extension in debug mode (uses `.vscode/launch.json`).

## Architecture

The project consists of two main components:

### 1. VSCode Extension (root)

```
src/
├── extension.ts              # Entry point - registers HwpxEditorProvider
├── editor/
│   ├── HwpxEditorProvider.ts # CustomEditorProvider implementation
│   └── webview.ts            # HTML/CSS/JS for the editor UI
├── hwpx/
│   ├── HwpxParser.ts         # HWPX format parser (ZIP-based XML)
│   ├── HwpxDocument.ts       # Document model and operations
│   └── types.ts              # HWPML type definitions
└── hwp/
    └── HwpParser.ts          # Legacy HWP binary format parser (read-only)
```

**Data flow:** Extension opens file → Parser extracts content → HwpxDocument holds state → Webview renders UI → User edits sent via postMessage → Document updated → Changes saved back to file.

### 2. MCP Server (mcp-server/)

Standalone Model Context Protocol server for AI integration. Provides 77+ tools for document operations.

```
mcp-server/src/
├── index.ts               # MCP server entry point, tool definitions
├── HwpxDocument.ts        # Core document model (open, save, edit operations)
├── HwpxParser.ts          # HWPX format parser
├── HangingIndentCalculator.ts  # Calculates hanging indent widths
├── types.ts               # Shared type definitions
└── *.test.ts              # Test files (vitest)
```

**Key methods in HwpxDocument.ts:**
- `open()` / `save()` - File I/O with atomic writes
- `getParagraphs()` / `updateParagraphText()` - Paragraph manipulation
- `getTables()` / `updateTableCell()` - Table editing with lineseg reset
- `insertImage()` / `renderMermaid()` - Image insertion with BinData registration

## Key Dependencies

- **hwp.js**: Parses legacy HWP binary format
- **jszip**: Handles HWPX ZIP archive structure
- **@modelcontextprotocol/sdk**: MCP protocol implementation (mcp-server only)

## File Format Notes

- **HWPX files** are ZIP archives containing XML files following the HWPML specification
- **HWP files** are binary OLE compound documents (read-only support via hwp.js)
- Type definitions in `src/hwpx/types.ts` are based on the official Korean HWPML specification

### HWPX ZIP Structure
```
*.hwpx (ZIP archive)
├── Contents/
│   ├── section0.xml      # Main document content (paragraphs, tables)
│   ├── header.xml        # Document settings (styles, charShapes, paraShapes)
│   └── content.hpf       # Manifest (lists all files)
├── BinData/              # Embedded files (images)
├── Preview/              # Thumbnail images
└── version.xml           # Format version info
```

### Unit Conversion
HWPML uses "hwpunit" where **1 point = 100 hwpunit**:
```
hwpunit = points × 100
10pt font → 1000 hwpunit
```

## Reference Documentation

The `reference/docs/` folder contains official HWP specification documents:
- `한글문서파일형식3.0_HWPML_revision1.2.md` - HWPML (XML format) specification
- `한글문서파일형식_5.0_revision1.3.md` - HWP 5.0 binary format specification

These documents define all supported elements including:
- Document structure (HEAD, BODY, TAIL)
- Text elements (paragraphs, runs, special characters)
- Drawing objects (tables, images, shapes, equations)
- Control elements (headers/footers, footnotes/endnotes, bookmarks, fields)
- Script and template support

## 한글 호환성 주의사항 (HWPML Critical Notes)

### 이미지 삽입 시 필수 사항

한글에서 이미지가 정상적으로 표시되려면 다음 XML 구조를 **반드시** 따라야 합니다:

1. **`<hp:curSz>`는 반드시 `width="0" height="0"`으로 설정**
   ```xml
   <hp:orgSz width="15000" height="21000"/>
   <hp:curSz width="0" height="0"/>  <!-- 필수! 실제 크기가 아닌 0,0 -->
   ```
   - `orgSz`에는 실제 크기 (hwpunit = points × 100)
   - `curSz`는 항상 0, 0 (한글이 자동 계산)

2. **이미지는 반드시 `<hp:run>` 태그로 감싸야 함**
   ```xml
   <hp:run charPrIDRef="0">
     <hp:pic ...>...</hp:pic>
     <hp:t/>
   </hp:run>
   ```
   - `<hp:pic>` 앞에 `<hp:run>` 시작
   - `<hp:pic>` 뒤에 `<hp:t/>` 추가
   - 그 후 `</hp:run>` 닫기

3. **테이블 셀 내 이미지 위치 속성**
   ```xml
   <hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1"
          allowOverlap="0" holdAnchorAndSO="0"
          vertRelTo="PARA" horzRelTo="COLUMN"
          vertAlign="TOP" horzAlign="LEFT"
          vertOffset="0" horzOffset="0"/>
   ```

### insert_table 주의사항

- `insertTable()`은 메모리에만 테이블을 추가
- `save()` 시 XML에 반영되려면 `_pendingTableInserts` 배열에 추가 필요
- `applyTableInsertsToXml()`가 실제 XML 생성 담당

### updateParagraphText 동작 방식 (2026-01-23 수정)

**문제**: 복잡한 문서에서 `updateParagraphText` 호출 후 저장 → 재로드 시 텍스트가 이상하게 합쳐짐

**원인**: 한글 문서의 paragraph는 여러 개의 `<hp:run>` 요소를 가질 수 있음
```xml
<hp:p id="123">
  <hp:run><hp:t>첫 번째 텍스트</hp:t></hp:run>
  <hp:run><hp:t>두 번째 텍스트</hp:t></hp:run>
  <hp:run><hp:t>세 번째 텍스트</hp:t></hp:run>
</hp:p>
```

**해결**: `updateParagraphText`가 run 0을 업데이트할 때, 다른 runs도 빈 문자열로 클리어
```typescript
// HwpxDocument.ts - updateParagraphText 메서드
if (runIndex === 0) {
  for (let i = 1; i < paragraph.runs.length; i++) {
    paragraph.runs[i].text = '';  // 다른 runs 클리어
  }
}
```

**주의사항**:
- `updateParagraphText`는 전체 문단 텍스트 교체용
- 특정 run만 수정하려면 runIndex를 명시적으로 지정
- 여러 run의 스타일을 유지하면서 텍스트만 바꾸려면 `update_paragraph_text_preserve_styles` 사용

### 디버깅 팁

한글에서 이미지가 안 보일 때:
1. 한글에서 직접 만든 파일의 XML 구조와 비교
2. `unzip -p file.hwpx Contents/section0.xml`로 XML 추출
3. `<hp:pic>` 요소의 부모 태그 확인 (`<hp:run>` 안에 있어야 함)
4. `<hp:curSz>` 값 확인 (0, 0이어야 함)
