/**
 * 종단간: 실제 MCP 서버 프로세스를 stdio 로 띄워 사용자와 같은 경로로 부른다.
 *
 * 서버는 HWPX_MCP_SERVER 로 고른다 (tests/helpers/mcp-client.ts).
 *   local        → 이 저장소의 dist/index.js (기본)
 *   npm:0.3.3    → 게시된 버전
 * 같은 파일을 여러 버전에 돌리면 "어느 버전부터 고쳐졌나"가 표로 나온다
 * (npm run test:versions).
 *
 * 판정은 저장된 파일을 서버가 아닌 쪽(JSZip·XML)에서 직접 읽어 한다. 서버의
 * 성공 응답만 믿지 않는다 — 이 저장소가 반복해서 겪은 결함이 "성공이라 했는데
 * 저장본은 틀림"이었다.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { McpClient } from '../helpers/mcp-client';
import { assertBalanced, rowAddrsOfFirstTable } from '../helpers/hwpx';

const server = McpClient.serverLabel();
let workDir: string;
let mcp: McpClient;

beforeAll(async () => {
  workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-e2e-'));
  mcp = await McpClient.start(workDir);
}, 180_000);

afterAll(() => {
  mcp?.close();
  if (workDir) fs.rmSync(workDir, { recursive: true, force: true });
});

async function savedSection(file: string, index = 0): Promise<string> {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const f = zip.file(`Contents/section${index}.xml`);
  if (!f) throw new Error(`section${index}.xml missing in ${file}`);
  return f.async('string');
}

/** Top-level order of a saved section: table wrappers as 'T', paragraphs as their text. */
function topLevelOrder(xml: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (const m of xml.matchAll(/<(\/?)hp:p\b[^>]*?(\/?)>/g)) {
    if (m[1]) {
      depth--;
      if (depth === 0) {
        const block = xml.slice(start, m.index! + m[0].length);
        if (/<hp:tbl\b/.test(block)) out.push('T');
        else {
          const t = [...block.matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(x => x[1]).join('');
          if (t) out.push(t);
        }
      }
    } else if (!m[2]) {
      if (depth === 0) start = m.index!;
      depth++;
    }
  }
  return out;
}

async function newDoc(name: string): Promise<{ id: string; file: string }> {
  const file = path.join(workDir, `${name}.hwpx`);
  const created = await mcp.ok('create_document', { file_path: file });
  return { id: created.doc_id, file };
}

describe(`MCP stdio 종단간 [${server}]`, () => {
  it('서버가 뜨고 도구 목록을 준다', async () => {
    const tools = await mcp.listTools();
    expect(tools).toContain('create_document');
    expect(tools).toContain('save_document');
    expect(tools.length).toBeGreaterThan(50);
  });

  it('create_document(file_path) → save 가 그 경로에 파일을 만든다 (리뷰 1차 #2)', async () => {
    const { id, file } = await newDoc('path');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '경로 확인' });
    await mcp.ok('save_document', { doc_id: id });
    expect(fs.existsSync(file)).toBe(true);
    expect(topLevelOrder(await savedSection(file))).toContain('경로 확인');
  });

  it('문단 → 표 → 문단 이 저장본에서도 같은 순서다 (0.3.3)', async () => {
    const { id, file } = await newDoc('order');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '제목' });
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 2, cols: 2 });
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: 1, text: '표 아래' });
    await mcp.ok('save_document', { doc_id: id });
    expect(topLevelOrder(await savedSection(file))).toEqual(['제목', 'T', '표 아래']);
  });

  it('① 행 삽입 뒤 rowAddr 가 행마다 유일하다 — 한/글이 멈추지 않는 조건', async () => {
    const { id, file } = await newDoc('row');
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 4, cols: 4 });
    await mcp.ok('insert_table_row', { doc_id: id, section_index: 0, table_index: 0, after_row: 1, cell_texts: ['a', 'b', 'c', 'd'] });
    await mcp.ok('save_document', { doc_id: id });
    const xml = await savedSection(file);
    expect(rowAddrsOfFirstTable(xml)).toEqual(['0', '1', '2', '3', '4']);
    expect(assertBalanced(xml)).toEqual({});
  });

  it('③ 두 구역 문서: get_table_map 의 구역 안 순번으로 쓰면 본문 표에 들어간다', async () => {
    const { id, file } = await newDoc('sections');
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 2, cols: 2 });
    await mcp.ok('insert_section', { doc_id: id, after_index: 0 });
    await mcp.ok('insert_table', { doc_id: id, section_index: 1, after_index: 0, rows: 3, cols: 3 });
    const map = (await mcp.ok('get_table_map', { doc_id: id })).table_map;
    const body = map.find((m: { section_index: number }) => m.section_index === 1);
    expect(body.table_index_in_section).toBe(0);
    await mcp.ok('update_table_cell', { doc_id: id, section_index: 1, table_index: body.table_index_in_section, row: 1, col: 1, text: '본문 표' });
    await mcp.ok('save_document', { doc_id: id });
    // 구역 파일이 둘 다 있고, 글은 본문 구역(section1)의 표에 있다.
    expect(await savedSection(file, 1)).toContain('>본문 표<');
    expect(await savedSection(file, 0)).not.toContain('>본문 표<');
  });

  it('④ 열 삽입 뒤 칸 폭 합이 표 폭과 같다', async () => {
    const { id, file } = await newDoc('col');
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 4, cols: 4 });
    await mcp.ok('insert_table_column', { doc_id: id, section_index: 0, table_index: 0, after_col: 1 });
    await mcp.ok('save_document', { doc_id: id });
    const xml = await savedSection(file);
    const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    const width = +t.match(/<hp:sz width="(\d+)"/)![1];
    const firstRow = t.slice(0, t.indexOf('</hp:tr>'));
    const sum = [...firstRow.matchAll(/<hp:cellSz width="(\d+)"/g)].reduce((a, m) => a + +m[1], 0);
    expect(sum).toBe(width);
  });

  it('⑥ 실패한 호출은 isError: true 로 온다', async () => {
    const missing = await mcp.call('insert_paragraph', { text: '구역 번호 없음' });
    expect(missing.isError).toBe(true);
    expect(missing.raw).toMatch(/section_index/);

    const unknownDoc = await mcp.call('get_paragraphs', { doc_id: 'no-such-doc' });
    expect(unknownDoc.isError).toBe(true);
  });

  it('성공한 호출은 isError 가 아니다', async () => {
    const { id } = await newDoc('ok');
    const r = await mcp.call('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '정상' });
    expect(r.isError).toBe(false);
  });

  // Last on purpose: every call above has run by now, so this covers stdout
  // from the whole session, not just start-up.
  it('서버가 stdout 에 JSON-RPC 외의 출력을 쓰지 않는다 (MCP stdio 규약)', () => {
    expect(() => mcp.assertCleanStdout()).not.toThrow();
  });
});
