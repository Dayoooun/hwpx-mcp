/**
 * export_to_pdf 종단간 — 실제 MCP 서버에 한컴 엔진으로 PDF 를 뽑게 하고, PDF 를 서버가 아닌
 * 쪽에서 읽어 판정한다 (docs/specs/export-pdf.md).
 *
 * 엔진이 있는 곳(맥 + 한/글 for Mac)에서만 변환 시험이 돈다. CI(리눅스)는 "엔진 없음" 시험만 돈다.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { McpClient } from '../helpers/mcp-client';

const hasMacEngine = process.platform === 'darwin'
  && fs.existsSync('/Applications/Hancom Office HWP.app/Contents/PlugIns/HwpQLExtPreview.appex')
  && fs.existsSync(path.resolve(__dirname, '../../native/hwp2pdf'));

let workDir: string;
let mcp: McpClient;

beforeAll(async () => {
  workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-pdf-e2e-'));
  mcp = await McpClient.start(workDir);
}, 180_000);

afterAll(() => {
  mcp?.close();
  if (workDir) fs.rmSync(workDir, { recursive: true, force: true });
});

/** Page count and decoded content-stream bytes of a PDF, read without the server. */
function readPdf(file: string): { pages: number; streams: string } {
  const buf = fs.readFileSync(file);
  expect(buf.subarray(0, 5).toString('latin1')).toBe('%PDF-');
  const raw = buf.toString('latin1');
  const pages = (raw.match(/\/Type\s*\/Page\b/g) || []).length;
  let streams = '';
  for (const m of raw.matchAll(/stream\r?\n/g)) {
    const start = m.index! + m[0].length;
    const end = raw.indexOf('endstream', start);
    try { streams += zlib.inflateSync(buf.subarray(start, end)).toString('latin1'); } catch { /* not deflate */ }
  }
  return { pages, streams };
}

async function docWithText(name: string, text: string) {
  const file = path.join(workDir, `${name}.hwpx`);
  const { doc_id } = await mcp.ok('create_document', { file_path: file });
  await mcp.ok('insert_paragraph', { doc_id, section_index: 0, after_index: -1, text });
  return { doc_id, file };
}

describe('export_to_pdf', () => {
  it('도구 목록에 있다', async () => {
    expect(await mcp.listTools()).toContain('export_to_pdf');
  });

  it('doc_id 와 file_path 를 둘 다 주거나 둘 다 빼면 isError', async () => {
    const out = path.join(workDir, 'x.pdf');
    expect((await mcp.call('export_to_pdf', { output_path: out })).isError).toBe(true);
    expect((await mcp.call('export_to_pdf', { output_path: out, doc_id: 'a', file_path: 'b' })).isError).toBe(true);
    expect(fs.existsSync(out)).toBe(false);
  });

  it.runIf(!hasMacEngine && process.platform !== 'win32')('export-pdf#AC-4 엔진이 없으면 isError 이고 PDF 를 만들지 않는다', async () => {
    const { doc_id, file } = await docWithText('noengine', '엔진 없음');
    await mcp.ok('save_document', { doc_id });
    const out = path.join(workDir, 'noengine.pdf');
    const r = await mcp.call('export_to_pdf', { file_path: file, output_path: out });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/macOS|Windows|not installed|converter missing/);
    expect(fs.existsSync(out)).toBe(false);
  });

  it.runIf(hasMacEngine)('export-pdf#AC-5 file_path 로 변환하면 쪽과 글이 든 PDF 가 나온다', async () => {
    const { doc_id, file } = await docWithText('disk', '파일에서 바로 PDF 로');
    await mcp.ok('save_document', { doc_id });
    const out = path.join(workDir, 'disk.pdf');
    const r = await mcp.ok('export_to_pdf', { file_path: file, output_path: out });
    expect(r).toMatchObject({ output_path: out, pages: 1, engine: 'hancom-mac' });
    const pdf = readPdf(out);
    expect(pdf.pages).toBe(1);
    // 글자는 CID 로 그려진다 — 텍스트 그리기 연산자가 있어야 글이 이미지가 아니라 글자로 들어간 것이다.
    expect(pdf.streams).toMatch(/\bTJ?\b/);
    expect(fs.readdirSync(workDir).filter(f => f.startsWith('.hwpx-pdf-'))).toEqual([]);
  });

  it.runIf(hasMacEngine)('export-pdf#AC-6 doc_id 로 변환하면 저장 전에 고친 글도 PDF 에 나온다', async () => {
    const { doc_id, file } = await docWithText('unsaved', '첫 문단');
    await mcp.ok('save_document', { doc_id });
    const before = path.join(workDir, 'before.pdf');
    await mcp.ok('export_to_pdf', { file_path: file, output_path: before });
    for (let i = 0; i < 60; i++) {
      await mcp.ok('insert_paragraph', { doc_id, section_index: 0, after_index: i, text: `저장 안 한 문단 ${i}` });
    }
    const out = path.join(workDir, 'unsaved.pdf');
    const r = await mcp.ok('export_to_pdf', { doc_id, output_path: out });
    // 61 문단은 한 쪽에 안 들어간다. 디스크 파일(1문단)을 그렸다면 1쪽이다.
    expect(r.pages).toBeGreaterThan(1);
    expect(readPdf(out).pages).toBe(r.pages);
    expect(readPdf(before).pages).toBe(1);
    // 내보내기는 저장이 아니다 — 저장 안 한 편집은 여전히 저장 안 한 것으로 보여야 닫기 전에 저장한다.
    const { documents } = await mcp.ok('list_open_documents', {});
    expect(documents.find((d: { id: string }) => d.id === doc_id)?.isDirty).toBe(true);
    await mcp.ok('save_document', { doc_id });
    const saved = path.join(workDir, 'saved.pdf');
    expect((await mcp.ok('export_to_pdf', { file_path: file, output_path: saved })).pages).toBe(r.pages);
  });

  it('export-pdf#AC-9 출력 경로가 원본 문서면 거절하고 원본을 그대로 둔다', async () => {
    const { doc_id, file } = await docWithText('self', '덮어쓰면 안 되는 원본');
    await mcp.ok('save_document', { doc_id });
    const before = fs.readFileSync(file);
    const viaFile = await mcp.call('export_to_pdf', { file_path: file, output_path: file });
    const viaDoc = await mcp.call('export_to_pdf', { doc_id, output_path: file });
    const viaAlias = await mcp.call('export_to_pdf', { file_path: file, output_path: path.join(workDir, '.', 'self.hwpx') });
    for (const r of [viaFile, viaDoc, viaAlias]) {
      expect(r.isError).toBe(true);
      expect(r.raw).toMatch(/source document/);
    }
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it.runIf(hasMacEngine)('export-pdf#AC-7 열 수 없는 문서는 isError 이고 있던 출력 파일은 그대로다', async () => {
    const bad = path.join(workDir, 'broken.hwpx');
    fs.writeFileSync(bad, 'PK\u0003\u0004 not really a zip');
    const out = path.join(workDir, 'keep.pdf');
    fs.writeFileSync(out, 'previous');
    const r = await mcp.call('export_to_pdf', { file_path: bad, output_path: out });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/could not open|failed/);
    expect(fs.readFileSync(out, 'utf8')).toBe('previous');
    expect(fs.readdirSync(workDir).filter(f => f.startsWith('.hwpx-pdf-'))).toEqual([]);
  });
});
