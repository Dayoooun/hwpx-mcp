# Tasks: add export_to_pdf

## 1. Implementation

- [x] 1.1 `PdfExport.ts`: engine choice, Windows COM script, staged write, timeout reporting, source-path guard
- [x] 1.2 `index.ts`: `export_to_pdf` tool schema and handler (`doc_id` | `file_path`)
- [x] 1.3 `HwpxDocument.serializeForExport()` keeps the unsaved state
- [x] 1.4 `native/hwp2pdf.m`: confine engine temp dir (`HspTempSetDir`), never call `releaseEngine`, per-user
      lock (HOME or passwd home; fail when it cannot be taken), `objc_msgSend_stret` on x86_64,
      libraries resolved from `HWP_APP`
- [x] 1.5 `native/build.sh` wired into `npm run build`; binary shipped in the package

## 2. Testing

- [x] 2.1 Unit: AC-1…AC-4, timeout (`tests/unit/pdf-export.unit.test.ts`)
- [x] 2.2 E2E on macOS + 한/글: AC-5, AC-6 (incl. `isDirty`), AC-7, AC-9 (`tests/e2e/pdf-export.e2e.test.ts`)
- [x] 2.3 Converter: x86_64 slice, moved app, no-HOME lock, shared-temp deletions, concurrency
      (`~/tools/hwp-render/test_hwp2pdf.py`)
- [x] 2.4 AC-8 manual: Windows 한/글 2024 via the packed server

## 3. Validation

- [x] 3.1 40 random documents compared with Windows 한/글 PDFs (pages 39/40 equal)
