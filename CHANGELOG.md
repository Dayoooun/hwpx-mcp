# Changelog

이 저장소는 두 개의 패키지를 담고 있습니다.

| 경로 | 패키지 | 변경 이력 |
|---|---|---|
| `/` | `hwpx-editor` (VS Code 확장) | 이 파일 |
| `/mcp-server` | [`@kimdayoun/hwpx-mcp`](https://www.npmjs.com/package/@kimdayoun/hwpx-mcp) (MCP 서버) | [mcp-server/CHANGELOG.md](./mcp-server/CHANGELOG.md) |

MCP 서버는 독립적으로 버전을 매기고 npm 에 배포합니다. 서버 쪽 변경을 찾는다면
[mcp-server/CHANGELOG.md](./mcp-server/CHANGELOG.md) 를 보세요.

---

## VS Code 확장

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르고,
버전은 [유의적 버전](https://semver.org/lang/ko/) 을 따릅니다.

### [0.1.0] - 2025-01-12

#### Added

- HWPX 파일 읽기/쓰기 지원
- 텍스트 편집 기능
- 테이블 보기 및 편집
- 문서 메타데이터 확인
- MCP (Model Context Protocol) 서버 포함
- AI 도구 연동 지원 (Claude 등)
