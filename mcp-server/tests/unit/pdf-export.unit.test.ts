/**
 * export_to_pdf 엔진 고르기와 윈도우 호출 모양 (docs/specs/export-pdf.md).
 * 실제 변환은 tests/e2e/pdf-export.e2e.test.ts 가 맥 한/글로 본다.
 */
import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pickEngine, windowsInvocation, WINDOWS_SCRIPT, exportPdf } from '../../src/PdfExport';

describe('pickEngine', () => {
  it('export-pdf#AC-1 auto 는 맥에서 hancom-mac, 윈도우에서 hancom-windows', () => {
    expect(pickEngine('auto', 'darwin')).toBe('hancom-mac');
    expect(pickEngine(undefined, 'darwin')).toBe('hancom-mac');
    expect(pickEngine('auto', 'win32')).toBe('hancom-windows');
  });

  it('export-pdf#AC-1 그 밖의 플랫폼은 이유를 적어 거절한다', () => {
    expect(() => pickEngine('auto', 'linux')).toThrow(/macOS.*Windows.*linux/);
  });

  it('export-pdf#AC-2 명시한 엔진이 이 플랫폼 것이 아니면 거절한다', () => {
    expect(() => pickEngine('hancom-mac', 'linux')).toThrow(/runs only on darwin/);
    expect(() => pickEngine('hancom-windows', 'darwin')).toThrow(/runs only on win32/);
    expect(pickEngine('hancom-mac', 'darwin')).toBe('hancom-mac');
    expect(() => pickEngine('libreoffice' as never, 'darwin')).toThrow(/Unknown engine/);
  });
});

describe('windowsInvocation', () => {
  it('export-pdf#AC-3 경로는 환경변수로만 가고 스크립트 본문·인자에 없다', () => {
    const evil = `C:\\a\\x"; Remove-Item C:\\ -Recurse; $(calc) '.hwpx`;
    const out = `C:\\b\\o'$env:USERNAME".pdf`;
    const inv = windowsInvocation('C:\\t\\convert.ps1', evil, out);
    expect(inv.env).toEqual({ HWPX_PDF_IN: evil, HWPX_PDF_OUT: out });
    expect(inv.args.join(' ')).not.toContain(evil);
    expect(inv.args.join(' ')).not.toContain(out);
    expect(WINDOWS_SCRIPT).toContain('$env:HWPX_PDF_IN');
    expect(WINDOWS_SCRIPT).toContain('$env:HWPX_PDF_OUT');
    // 스크립트는 고정 문자열이라 어떤 경로도 그 안에 들어가지 않는다.
    expect(WINDOWS_SCRIPT).not.toMatch(/C:\\/);
  });
});

describe('exportPdf', () => {
  it('export-pdf#AC-4 엔진이 없으면 실패하고 출력 파일을 만들지 않는다', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-unit-'));
    try {
      const src = path.join(dir, 'a.hwpx');
      fs.writeFileSync(src, 'x');
      const out = path.join(dir, 'a.pdf');
      await expect(exportPdf(src, out, { platform: 'darwin', env: { HWP_APP: path.join(dir, 'none.app') } }))
        .rejects.toThrow(/not installed/);
      await expect(exportPdf(src, out, { platform: 'linux' })).rejects.toThrow(/macOS/);
      expect(fs.readdirSync(dir)).toEqual(['a.hwpx']);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
