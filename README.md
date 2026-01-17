# HWPX Editor for VSCode

[![GitHub](https://img.shields.io/badge/GitHub-mjyoo2%2Fhwp--extension-blue?logo=github)](https://github.com/mjyoo2/hwp-extension)

VSCode에서 한글(HWPX) 문서를 열고 편집할 수 있는 확장 프로그램입니다.

## Features

### 문서 보기 및 편집
- **HWPX 파일**: 읽기 및 편집 지원 (XML 기반 최신 포맷)

### 지원 기능
- 텍스트 내용 보기 및 편집
- 단락 추가/수정/삭제
- 테이블 보기 및 편집
- 문서 메타데이터 확인
- 문서 구조 탐색

### MCP (Model Context Protocol) 서버
AI 도구(Claude 등)와 연동하여 문서를 자동으로 편집할 수 있는 MCP 서버를 포함합니다.

## Installation

1. VSCode에서 Extensions (Ctrl+Shift+X) 열기
2. "HWPX Editor" 검색
3. Install 클릭

## Usage

### 기본 사용법
1. HWPX 파일을 VSCode에서 열기
2. 자동으로 HWPX Editor가 활성화됨
3. 문서 내용 확인 및 편집

### MCP 서버 사용 (AI 연동)

Command Palette (Ctrl+Shift+P)에서:
- `HWPX: Show MCP Server Configuration` - MCP 설정 정보 확인
- `HWPX: Copy MCP Server Path` - MCP 서버 경로 복사

#### Claude Code에서 사용

`.vscode/mcp.json` 파일 생성:

```json
{
  "mcpServers": {
    "hwpx": {
      "command": "node",
      "args": ["${extensionPath}/out/mcp-server.js"]
    }
  }
}
```

## Supported File Formats

| 포맷 | 확장자 | 읽기 | 쓰기 |
|------|--------|------|------|
| HWPX | .hwpx | O | O |

> **Note**: HWP (바이너리) 파일은 지원하지 않습니다. 한컴오피스에서 HWPX로 변환 후 사용해주세요.

## Requirements

- VSCode 1.107.0 이상

## Known Issues

- HWP 파일 미지원 (HWPX로 변환 필요)
- 일부 복잡한 서식은 표시되지 않을 수 있습니다

## Release Notes

### 0.1.1
- **버그 수정**: 테이블 빈 셀 업데이트 후 저장 시 변경사항이 유지되지 않던 문제 수정
  - Self-closing XML run 태그 (`<hp:run ... />`) 처리 지원
  - ID 기반 테이블 매칭으로 안정적인 XML 업데이트
- 테이블 셀 업데이트 테스트 코드 추가

### 0.1.0
- 최초 릴리스
- HWPX 파일 읽기/쓰기 지원
- MCP 서버 포함

## License

MIT

## Contributing

GitHub: https://github.com/mjyoo2/hwp-extension

버그 리포트 및 기능 요청은 GitHub Issues를 이용해주세요.
