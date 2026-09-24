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

  /**
   * Rewrite section0 of a saved file and reopen it through the server — for
   * shapes 한/글 writes but the tools cannot build (mixed character shapes in one
   * paragraph, a heading and a table in one paragraph, repeated paragraph ids).
   */
  async function reopenEdited(file: string, name: string, edit: (xml: string) => string): Promise<string> {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const xml = await zip.file('Contents/section0.xml')!.async('string');
    const next = edit(xml);
    if (next === xml) throw new Error(`fixture edit for ${name} matched nothing`);
    zip.file('Contents/section0.xml', next);
    const edited = path.join(workDir, `${name}.hwpx`);
    fs.writeFileSync(edited, await zip.generateAsync({ type: 'nodebuffer' }));
    return (await mcp.ok('open_document', { file_path: edited })).doc_id;
  }

  /** (charPrIDRef, own text) of each text run of the saved paragraph holding `marker`. */
  function runsOfParagraphWith(xml: string, marker: string): Array<{ char: string; text: string }> {
    const at = xml.indexOf(marker);
    if (at < 0) return [];
    const p = xml.slice(xml.lastIndexOf('<hp:p ', at), xml.indexOf('</hp:p>', at));
    return [...p.matchAll(/<hp:run charPrIDRef="(\d+)"[^>]*>([\s\S]*?)<\/hp:run>/g)]
      .map(m => ({ char: m[1], text: [...m[2].matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(t => t[1]).join('') }))
      .filter(r => r.text);
  }

  it('⑤ 글자 모양이 섞인 문단을 update_paragraph_text 로 통째로 바꾸면 새 글이 모두 첫 글자 모양이다', async () => {
    // 회신 ⑤: 앞은 보통·뒤는 굵게인 문단을 통째로 바꾸면 셋째 줄쯤부터 굵게 바뀐다.
    // 0.3.4 까지 처리부가 run 이 여럿이면 길이 비율로 나눠 담는 쪽으로 넘겼다.
    const { id, file } = await newDoc('mixed-src');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: 'placeholder' });
    await mcp.ok('save_document', { doc_id: id });
    const doc = await reopenEdited(file, 'mixed', x => x.replace(
      /(<hp:run charPrIDRef=")(\d+)(">)<hp:t>placeholder<\/hp:t><\/hp:run>/,
      '$1$2$3<hp:t>보통으로 쓴 앞부분 문장입니다. </hp:t></hp:run><hp:run charPrIDRef="7"><hp:t>여기부터 굵게 쓴 뒷부분 문장입니다.</hp:t></hp:run>'));

    const para = (await mcp.ok('get_paragraphs', { doc_id: doc, section_index: 0 })).paragraphs
      .find((p: { text: string }) => p.text.includes('보통으로'));
    const NEW = '문단 전체를 새 문장으로 바꿉니다. 이 문장은 길어서 한 줄을 넘기고 둘째 줄과 셋째 줄까지 이어집니다.';
    await mcp.ok('update_paragraph_text', { doc_id: doc, section_index: 0, paragraph_index: para.index, text: NEW });
    const out = path.join(workDir, 'mixed-out.hwpx');
    await mcp.ok('save_document', { doc_id: doc, output_path: out });

    const xml = await savedSection(out);
    expect(assertBalanced(xml)).toEqual({});
    const firstShape = runsOfParagraphWith(xml, NEW.slice(0, 8))[0]?.char;
    // 새 글은 한 run 에 통째로 있고, 굵은 run(7) 에는 글이 남지 않는다.
    expect(runsOfParagraphWith(xml, NEW.slice(0, 8))).toEqual([{ char: firstShape, text: NEW }]);
    expect(firstShape).not.toBe('7');
  });

  it('⑤ 글자 모양을 run 마다 지키려면 update_paragraph_text_preserve_styles 가 그대로 나눠 담는다', async () => {
    const { id, file } = await newDoc('mixed-keep-src');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: 'placeholder' });
    await mcp.ok('save_document', { doc_id: id });
    const doc = await reopenEdited(file, 'mixed-keep', x => x.replace(
      /(<hp:run charPrIDRef=")(\d+)(">)<hp:t>placeholder<\/hp:t><\/hp:run>/,
      '$1$2$3<hp:t>굵지않음 </hp:t></hp:run><hp:run charPrIDRef="7"><hp:t>굵음</hp:t></hp:run>'));
    const para = (await mcp.ok('get_paragraphs', { doc_id: doc, section_index: 0 })).paragraphs
      .find((p: { text: string }) => p.text.includes('굵지않음'));
    await mcp.ok('update_paragraph_text_preserve_styles', { doc_id: doc, section_index: 0, paragraph_index: para.index, text: '보통문장 굵은' });
    const out = path.join(workDir, 'mixed-keep-out.hwpx');
    await mcp.ok('save_document', { doc_id: doc, output_path: out });
    const runs = runsOfParagraphWith(await savedSection(out), '보통');
    expect(runs.map(r => r.char)).toContain('7');
    expect(runs.map(r => r.text).join('')).toBe('보통문장 굵은');
  });

  it('② 제목 글과 목차 표를 품은 문단을 preserve_styles 로 고치면 제목만 바뀌고 저장본이 정상이다', async () => {
    // 회신 ②: 연구보고서 "제1장 연구의 개요(작성중)1" 은 목차 표를 품은 문단이었다.
    // 0.3.3 은 한/글 원본 60건 중 44건에서 표 칸 글자를 바꿨다. 한/글처럼 문단 id 를 모두 0 으로 둔다.
    const { id, file } = await newDoc('toc-src');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '표지' });
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 3, cols: 2 });
    for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++)
      await mcp.ok('update_table_cell', { doc_id: id, section_index: 0, table_index: 0, row: r, col: c, text: `목차${r}${c}` });
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: 1, text: '본문' });
    await mcp.ok('save_document', { doc_id: id });
    const doc = await reopenEdited(file, 'toc', x => x
      .replace(/(<hp:p [^>]*><hp:run[^>]*>)(<hp:tbl)/, '$1<hp:t>제1장 연구의 개요(작성중)1</hp:t></hp:run><hp:run charPrIDRef="0">$2')
      .replace(/<hp:p id="[^"]*"/g, '<hp:p id="0"'));

    const heading = (await mcp.ok('get_paragraphs', { doc_id: doc, section_index: 0 })).paragraphs
      .find((p: { text: string }) => p.text.includes('제1장'));
    await mcp.ok('update_paragraph_text_preserve_styles', { doc_id: doc, section_index: 0, paragraph_index: heading.index, text: '시험 문구' });
    const out = path.join(workDir, 'toc-out.hwpx');
    await mcp.ok('save_document', { doc_id: doc, output_path: out, verify_integrity: true });

    const xml = await savedSection(out);
    expect(assertBalanced(xml)).toEqual({});
    expect(xml.replace(/<[^>]+>/g, '')).toContain('시험 문구');
    expect([...xml.matchAll(/<hp:t>(목차\d\d)<\/hp:t>/g)].map(m => m[1]))
      .toEqual(['목차00', '목차01', '목차10', '목차11', '목차20', '목차21']);
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

  it('한 행의 칸을 모두 덮는 병합은 isError 로 거절되고 저장본의 모든 행에 칸이 남는다', async () => {
    const { id, file } = await newDoc('merge-empty-row');
    await mcp.ok('insert_table', { doc_id: id, section_index: 0, after_index: 0, rows: 3, cols: 3 });
    const r = await mcp.call('merge_cells', { doc_id: id, section_index: 0, table_index: 0, start_row: 0, start_col: 0, end_row: 1, end_col: 2 });
    expect(r.isError).toBe(true);
    expect(r.raw).toMatch(/row 1 would have no cell of its own/);
    await mcp.ok('save_document', { doc_id: id });
    const t = await savedSection(file);
    const tbl = t.slice(t.indexOf('<hp:tbl'), t.indexOf('</hp:tbl>'));
    const tcPerRow = [...tbl.matchAll(/<hp:tr(?:\s[^>]*)?>([\s\S]*?)<\/hp:tr>/g)].map(m => (m[1].match(/<hp:tc[\s>]/g) ?? []).length);
    expect(tcPerRow).toEqual([3, 3, 3]);
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

  it('② (나) 닫는 태그가 어긋난 문서는 verify_integrity 저장이 실패하고 파일을 남기지 않는다', async () => {
    const { id, file } = await newDoc('verify-src');
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: -1, text: '첫 문단' });
    await mcp.ok('insert_paragraph', { doc_id: id, section_index: 0, after_index: 0, text: '둘째 문단' });
    await mcp.ok('save_document', { doc_id: id });

    // 회신 ② 저장본과 같은 증상: <hp:p> 닫힘이 하나 모자란 section.
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const xml = await zip.file('Contents/section0.xml')!.async('string');
    const cut = xml.lastIndexOf('</hp:p>');
    zip.file('Contents/section0.xml', xml.slice(0, cut) + xml.slice(cut + '</hp:p>'.length));
    const broken = path.join(workDir, 'verify-broken.hwpx');
    fs.writeFileSync(broken, await zip.generateAsync({ type: 'nodebuffer' }));

    const opened = await mcp.ok('open_document', { file_path: broken });
    const out = path.join(workDir, 'verify-out.hwpx');
    const saved = await mcp.call('save_document', { doc_id: opened.doc_id, output_path: out, verify_integrity: true });
    // 0.3.3: 성공 + integrity_verified: true 로 깨진 파일을 썼다.
    expect(saved.isError).toBe(true);
    expect(saved.raw).toMatch(/Malformed XML: Contents\/section0\.xml/);
    expect(fs.existsSync(out)).toBe(false);
  });

  // Last on purpose: every call above has run by now, so this covers stdout
  // from the whole session, not just start-up.
  it('서버가 stdout 에 JSON-RPC 외의 출력을 쓰지 않는다 (MCP stdio 규약)', () => {
    expect(() => mcp.assertCleanStdout()).not.toThrow();
  });
});
