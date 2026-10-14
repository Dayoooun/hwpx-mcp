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

  it('.hwp 파일은 빈 문서로 열리지 않고 오류가 난다', async () => {
    const file = path.join(workDir, 'old.hwp');
    fs.writeFileSync(file, Buffer.from('not parsed'));
    const r = await mcp.call('open_document', { file_path: file });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/not supported/);
  });
});
