# Change: Add `export_to_pdf` using Hancom's own engine on macOS and Windows

## Why

Users need PDFs laid out exactly as 한/글 lays them out. Until now that meant sending files to a Windows PC
running 한/글 COM; Mac users had no route. 한/글 for Mac ships a headless renderer for Quick Look
(`HwpQLExtPreview.appex` → `libSimpleEngineLib.dylib`), found by reverse engineering (REA/Ghidra) the app,
which draws pages straight onto a Core Graphics context without opening a window.

## What Changes

- New MCP tool `export_to_pdf` (`doc_id` or `file_path`, `output_path`, optional `engine`).
  - `auto` picks the engine from the server's OS: macOS → `hancom-mac`, Windows → `hancom-windows`;
    other platforms are refused with the reason.
  - `doc_id` renders the open document including unsaved edits, and leaves it marked unsaved.
  - Output is staged and renamed only after a valid PDF exists; an output path naming the source is refused.
- New native converter `mcp-server/native/hwp2pdf` (Objective-C, universal arm64/x86_64), built by
  `npm run build` on macOS and shipped in the npm package. It dlopens the user's installed engine and
  resolves the engine's libraries from that same app (`HWP_APP`).
- Windows path drives `HWPFrame.HwpObject` from a fixed PowerShell script; paths are passed only as
  environment variables.

## Impact

- Affected specs: `pdf-export` (new capability)
- Affected code: `mcp-server/src/PdfExport.ts` (new), `mcp-server/src/index.ts` (tool + handler),
  `mcp-server/src/HwpxDocument.ts` (`serializeForExport`), `mcp-server/native/` (new), `package.json`
- Engine hazards the converter must contain (measured):
  - `releaseEngine` → libhsp `HspTempClearDir` empties `NSTemporaryDirectory()`, the user's shared temp dir.
  - Concurrent engines draw blank pages while reporting success (3/180).
  - `paperExtent:` returns a 32-byte `CGRect` and needs `objc_msgSend_stret` on x86_64.
- Detailed acceptance criteria: `docs/specs/export-pdf.md` (AC-1…AC-9).
