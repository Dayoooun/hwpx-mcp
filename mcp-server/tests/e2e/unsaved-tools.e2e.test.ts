/**
 * 종단간: 저장 단계가 없는 도구는 성공이라 답하지 않고 오류를 돌려준다.
 * 근거와 목록은 src/unsavedTools.ts, 저장본 측정은 tests/module/unsaved-tools.test.ts.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { McpClient } from '../helpers/mcp-client';

let workDir: string;
let mcp: McpClient;

beforeAll(async () => {
  workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-e2e-unsaved-'));
  mcp = await McpClient.start(workDir);
}, 180_000);

afterAll(async () => {
  await mcp?.close();
  if (workDir) fs.rmSync(workDir, { recursive: true, force: true });
});

describe('저장되지 않는 도구', () => {
  it('insert_footnote 는 오류를 돌려주고 문서를 바꾸지 않는다', async () => {
    const id = (await mcp.ok('create_document', { file_path: path.join(workDir, 'u.hwpx') })).doc_id;
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '본문' });
    const r = await mcp.call('insert_footnote', { doc_id: id, section_index: 0, paragraph_index: 0, text: '각주' });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/not available/);
    expect((await mcp.ok('get_footnotes', { doc_id: id })).footnotes).toEqual([]);
  });

  it('문서 정보는 저장할 자리가 없는 새 문서에서 오류를 돌려주고 바꾸지 않는다', async () => {
    const id = (await mcp.ok('create_document', { file_path: path.join(workDir, 'm.hwpx'), title: '처음' })).doc_id;
    const r = await mcp.call('set_document_metadata', { doc_id: id, title: '바꾼 제목' });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/cannot save title/);
    expect((await mcp.ok('get_document_metadata', { doc_id: id })).metadata?.title).not.toBe('바꾼 제목');
  });

  it('일부만 저장되는 서식은 저장 안 되는 항목이 있을 때만 거절한다', async () => {
    const id = (await mcp.ok('create_document', { file_path: path.join(workDir, 's.hwpx') })).doc_id;
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '서식' });
    const at = { doc_id: id, section_index: 0, paragraph_index: 1 };
    const color = await mcp.call('set_text_style', { ...at, bold: true, font_color: '#FF0000' });
    expect(color.isError).toBe(true);
    expect(color.raw).toMatch(/cannot save font_color/);
    await mcp.ok('set_text_style', { ...at, bold: true });
    const margin = await mcp.call('set_paragraph_style', { ...at, align: 'center', margin_left: 20 });
    expect(margin.isError).toBe(true);
    await mcp.ok('set_paragraph_style', { ...at, align: 'center' });

    const tools = (await mcp.request('tools/list', {})).result.tools as Array<{ name: string; inputSchema: any }>;
    const textStyle = tools.find(t => t.name === 'set_text_style')!;
    expect(textStyle.inputSchema.properties.font_color.description).toMatch(/Not saved yet/);
    expect(textStyle.inputSchema.properties.bold.description ?? '').not.toMatch(/Not saved yet/);
  });

  it('set_cell_properties 는 오류를 돌려준다', async () => {
    const id = (await mcp.ok('create_document', { file_path: path.join(workDir, 'c.hwpx') })).doc_id;
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 1, cols: 1 });
    const r = await mcp.call('set_cell_properties', { doc_id: id, section_index: 0, table_index: 0, row: 0, col: 0, background_color: '#123456' });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/not available/);
  });

  it('문서 정보를 빈 값으로 지우는 요청은 오류다', async () => {
    const id = (await mcp.ok('create_document', { file_path: path.join(workDir, 'e.hwpx') })).doc_id;
    const r = await mcp.call('set_document_metadata', { doc_id: id, title: '' });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/cannot clear title/);
  });

  it('.hwp 파일은 빈 문서로 열리지 않고 오류가 난다', async () => {
    const file = path.join(workDir, 'old.hwp');
    fs.writeFileSync(file, Buffer.from('not parsed'));
    const r = await mcp.call('open_document', { file_path: file });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/not supported/);
  });
});
