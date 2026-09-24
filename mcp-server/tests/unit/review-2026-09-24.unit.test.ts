/**
 * 단위: 정창기 박사님 2차 회신 ①~⑥ 의 핵심 로직을 파일 저장 없이 직접 부른다.
 *
 * 회귀(tests/regression)·모듈(tests/module)은 저장 → 다시 열기로 판정하고,
 * 종단간(tests/e2e)은 MCP 서버를 stdio 로 띄워 판정한다. 여기서는 각 수정이 기대는
 * 순수 함수·메모리 동작의 경계값과 실패 경로를 본다. 비공개 XML 헬퍼는 문서
 * 인스턴스에서 (doc as any) 로 부른다 — 입력·출력이 문자열이라 파일 없이 판정된다.
 */
import { describe, it, expect } from 'vitest';
import { HwpxDocument } from '../../src/HwpxDocument';
import { success, error, findMissingArgs } from '../../src/ToolResult';

const doc = () => HwpxDocument.createNew('u', 'unit');
const priv = (d: HwpxDocument) => d as unknown as Record<string, (...a: unknown[]) => any>;
/** Text of the single text block a tool result carries. */
const textOf = (r: { content: Array<{ type: string; text?: string }> }) => {
  const c = r.content[0];
  if (c.type !== 'text' || c.text === undefined) throw new Error(`expected a text block, got ${c.type}`);
  return c.text;
};

/** A <hp:tc> in Hancom's shape: address/span/size after the sub-list. */
const tc = (col: number, row: number, text: string, opts: { colSpan?: number; rowSpan?: number; width?: number } = {}) =>
  `<hp:tc><hp:subList><hp:p><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p></hp:subList>` +
  `<hp:cellAddr colAddr="${col}" rowAddr="${row}"/>` +
  `<hp:cellSpan colSpan="${opts.colSpan ?? 1}" rowSpan="${opts.rowSpan ?? 1}"/>` +
  `<hp:cellSz width="${opts.width ?? 1000}" height="100"/></hp:tc>`;

describe('① 행 추가: rowAddr 다시 매기기 · 새 행 칸 격자', () => {
  it('shiftTableRowAddrs 는 fromRow 이상인 칸만 delta 만큼 옮긴다', () => {
    const tbl = `<hp:tbl><hp:tr>${tc(0, 0, 'a')}</hp:tr><hp:tr>${tc(0, 1, 'b')}</hp:tr><hp:tr>${tc(0, 2, 'c')}</hp:tr></hp:tbl>`;
    const out: string = priv(doc()).shiftTableRowAddrs(tbl, 1, +1);
    expect([...out.matchAll(/rowAddr="(\d+)"/g)].map(m => m[1])).toEqual(['0', '2', '3']);
  });

  it('shiftTableRowAddrs 는 칸 안 중첩 표의 rowAddr 를 건드리지 않는다', () => {
    const nested = `<hp:tbl><hp:tr>${tc(0, 5, '안')}</hp:tr></hp:tbl>`;
    const outer = `<hp:tbl><hp:tr>${tc(0, 1, nested.replace(/<\/?hp:t>/g, ''))}</hp:tr></hp:tbl>`;
    const withNested = outer.replace('<hp:t></hp:t>', nested);
    const out: string = priv(doc()).shiftTableRowAddrs(withNested, 0, +1);
    expect(out).toContain('rowAddr="5"');   // 중첩 표 칸 그대로
    expect(out).toContain('rowAddr="2"');   // 바깥 칸만 1 → 2
  });

  it('gridCellsForNewRow 는 세로 병합에 덮인 열까지 열마다 칸 하나를 준다', () => {
    // 행0: A(열0, 세로 2칸) B(열1)  /  행1: C(열1) — 열 0 은 위 병합에 덮였다.
    const rows = [
      { xml: `<hp:tr>${tc(0, 0, 'A', { rowSpan: 2 })}${tc(1, 0, 'B')}</hp:tr>` },
      { xml: `<hp:tr>${tc(1, 1, 'C')}</hp:tr>` },
    ];
    const cells: string[] = priv(doc()).gridCellsForNewRow(rows, 1);
    expect(cells.map(c => c.match(/colAddr="(\d+)"/)![1])).toEqual(['0', '1']);
  });

  it('insertTableRow 는 세로 병합 한가운데에 넣으면 거부하고 병합 끝 뒤에는 넣는다', () => {
    const d = doc();
    d.insertTable(0, 0, 4, 2);
    d.mergeCells(0, 0, 0, 0, 1, 0);            // 0~1행 세로 병합
    expect(() => d.insertTableRow(0, 0, 0)).toThrow(/split the merged cell at \(0, 0\).*Insert after row 1/);
    expect(d.insertTableRow(0, 0, 1)).toBe(true);
    const rowAddrs = d.findTable(0, 0)!.rows.map(r => r.cells[0].rowAddr);
    expect(rowAddrs).toEqual([0, 1, 2, 3, 4]);
  });
});

describe('② 표를 품은 문단: 문단의 자기 run 만 고른다', () => {
  it('findDirectChildRuns 는 표 칸 안 문단의 run 을 세지 않는다', () => {
    const p = '<hp:p id="0"><hp:run charPrIDRef="0"><hp:t>제1장</hp:t></hp:run>' +
      '<hp:run charPrIDRef="0"><hp:tbl><hp:tr><hp:tc><hp:subList>' +
      '<hp:p id="0"><hp:run charPrIDRef="0"><hp:t>목차 칸</hp:t></hp:run></hp:p>' +
      '</hp:subList></hp:tc></hp:tr></hp:tbl><hp:t></hp:t></hp:run></hp:p>';
    const runs: Array<{ xml: string }> = priv(doc()).findDirectChildRuns(p);
    expect(runs).toHaveLength(2);
    expect(runs[0].xml).toContain('제1장');
    expect(runs[1].xml).toContain('<hp:tbl>');   // 표 run 은 문단의 자기 run
  });

  it('ownRunText 는 run 안 표 칸의 글자를 빼고 run 자신의 글자만 준다', () => {
    const run = '<hp:run charPrIDRef="0"><hp:t>앞</hp:t><hp:tbl><hp:tr><hp:tc><hp:subList>' +
      '<hp:p><hp:run><hp:t>칸</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr></hp:tbl><hp:t>뒤</hp:t></hp:run>';
    expect(priv(doc()).ownRunText(run).replace(/<[^>]+>/g, '')).toBe('앞뒤');
  });

  it('updateParagraphText 는 표 요소를 가리키면 성공이라 하지 않고 대신 쓸 도구를 말한다', () => {
    const d = doc();
    d.insertTable(0, 0, 2, 2);
    const t = d.content.sections[0].elements.findIndex(e => e.type === 'table');
    expect(() => d.updateParagraphText(0, t, 0, 'x')).toThrow(/is a table, not a paragraph.*update_table_cell/);
  });
});

describe('③ 구역별 표 번호', () => {
  it('getTableMap 은 문서 전체 순번과 구역 안 순번을 함께 준다', () => {
    const d = doc();
    d.insertTable(0, 0, 1, 1);                 // 표지 구역
    d.insertSection(0);
    d.insertTable(1, -1, 1, 1);                // 본문 구역 첫 표
    d.insertTable(1, 0, 1, 1);                 // 본문 구역 둘째 표
    expect(d.getTableMap().map(t => [t.section_index, t.table_index, t.table_index_in_section]))
      .toEqual([[0, 0, 0], [1, 1, 0], [1, 2, 1]]);
  });

  it('구역 안 순번으로 findTable 하면 그 구역의 표가 나온다', () => {
    const d = doc();
    d.insertTable(0, 0, 1, 1);
    d.insertSection(0);
    d.insertTable(1, -1, 3, 3);
    const body = d.getTableMap().find(t => t.section_index === 1)!;
    expect(d.findTable(1, body.table_index_in_section)!.rows).toHaveLength(3);
    expect(d.findTable(1, body.table_index)).toBeNull();   // 전체 순번을 넣으면 없는 표
  });
});

describe('④ 열 추가 뒤 칸 폭: fitColumnsToTableWidth', () => {
  const tbl = (width: number, colCnt: number, rows: string[]) =>
    `<hp:tbl rowCnt="${rows.length}" colCnt="${colCnt}"><hp:sz width="${width}" height="1"/>${rows.map(r => `<hp:tr>${r}</hp:tr>`).join('')}</hp:tbl>`;
  const widthsOf = (xml: string) => xml.split('</hp:tr>').filter(r => r.includes('<hp:tr'))
    .map(r => [...r.matchAll(/<hp:cellSz width="(\d+)"/g)].map(m => +m[1]));

  it('칸 폭 합이 표 폭보다 크면 비율대로 줄이고 끝수는 마지막 칸에 준다', () => {
    const x = tbl(1000, 3, [tc(0, 0, 'a', { width: 400 }) + tc(1, 0, 'b', { width: 400 }) + tc(2, 0, 'c', { width: 700 })]);
    const [w] = widthsOf(priv(doc()).fitColumnsToTableWidth(x));
    expect(w.reduce((a: number, b: number) => a + b, 0)).toBe(1000);
    expect(w).toEqual([266, 266, 468]);
  });

  it('가로 병합 칸은 걸친 칸들의 줄인 폭의 합이 된다', () => {
    const x = tbl(900, 3, [
      tc(0, 0, 'm', { colSpan: 2, width: 1200 }) + tc(2, 0, 'c', { width: 600 }),
      tc(0, 1, 'a', { width: 600 }) + tc(1, 1, 'b', { width: 600 }) + tc(2, 1, 'd', { width: 600 }),
    ]);
    const [top, bottom] = widthsOf(priv(doc()).fitColumnsToTableWidth(x));
    expect(bottom).toEqual([300, 300, 300]);
    expect(top).toEqual([600, 300]);
  });

  it('이미 맞거나 어느 열 폭도 알 수 없으면 그대로 둔다', () => {
    const fit = tbl(1000, 2, [tc(0, 0, 'a', { width: 500 }) + tc(1, 0, 'b', { width: 500 })]);
    expect(priv(doc()).fitColumnsToTableWidth(fit)).toBe(fit);
    const onlyMerged = tbl(1000, 2, [tc(0, 0, 'm', { colSpan: 2, width: 3000 })]);
    expect(priv(doc()).fitColumnsToTableWidth(onlyMerged)).toBe(onlyMerged);
  });
});

describe('⑤ 글자 모양이 섞인 문단: 메모리에서 run 0 교체', () => {
  it('run 0 을 바꾸면 새 글은 run 0 에만 있고 나머지 run 은 빈다', () => {
    const d = doc();
    d.insertParagraph(0, -1, '');
    const p = d.content.sections[0].elements.findIndex(e => e.type === 'paragraph');
    d.updateParagraphRuns(0, p, [{ text: '보통 ', charPrIDRef: 0 }, { text: '굵게', charPrIDRef: 7 }] as any);
    d.updateParagraphText(0, p, 0, '새 문장 전체');
    expect(d.getParagraph(0, p)!.runs.map(r => [r.charPrIDRef, r.text])).toEqual([[0, '새 문장 전체'], [7, '']]);
  });

  it('preserve_styles 는 같은 문단을 run 길이 비율로 나눈다 (update_paragraph_text 와 다른 도구)', () => {
    const d = doc();
    d.insertParagraph(0, -1, '');
    const p = d.content.sections[0].elements.findIndex(e => e.type === 'paragraph');
    d.updateParagraphRuns(0, p, [{ text: 'ab', charPrIDRef: 0 }, { text: 'cd', charPrIDRef: 7 }] as any);
    expect(d.updateParagraphTextPreserveStyles(0, p, 'WXYZ')).toBe(true);
    expect(d.getParagraph(0, p)!.runs.map(r => r.text)).toEqual(['WX', 'YZ']);
  });

  it('run 이 없는 문단에 run 0 을 쓰면 run 을 만든다 · 없는 run 번호는 무시한다', () => {
    const d = doc();
    d.insertParagraph(0, -1, '');
    const p = d.content.sections[0].elements.findIndex(e => e.type === 'paragraph');
    d.updateParagraphRuns(0, p, [] as any);
    d.updateParagraphText(0, p, 0, '첫 글');
    expect(d.getParagraph(0, p)!.text).toBe('첫 글');
    d.updateParagraphText(0, p, 5, '무시');
    expect(d.getParagraph(0, p)!.text).toBe('첫 글');
  });
});

describe('⑥ 실패한 호출의 isError', () => {
  it('error() 는 isError: true 와 함께 JSON 본문에 오류 문장을 담는다', () => {
    const r = error('Document not found');
    expect(r.isError).toBe(true);
    expect(JSON.parse(textOf(r))).toEqual({ error: 'Document not found' });
  });

  it('success() 에는 isError 가 없다', () => {
    const r = success({ message: 'ok' });
    expect('isError' in r).toBe(false);
    expect(JSON.parse(textOf(r))).toEqual({ message: 'ok' });
  });

  it('findMissingArgs 는 빠지거나 null 인 필수값을 이름으로 돌려준다', () => {
    const req = new Map([['insert_paragraph', ['doc_id', 'section_index', 'after_index']]]);
    expect(findMissingArgs(req, 'insert_paragraph', { doc_id: 'd', section_index: null })).toEqual(['section_index', 'after_index']);
    expect(findMissingArgs(req, 'insert_paragraph', { doc_id: 'd', section_index: 0, after_index: -1 })).toEqual([]);
    expect(findMissingArgs(req, 'insert_paragraph', undefined)).toEqual(['doc_id', 'section_index', 'after_index']);
    expect(findMissingArgs(req, 'no_such_tool', {})).toEqual([]);
  });

  it('0 과 빈 문자열은 빠진 값이 아니다', () => {
    const req = new Map([['t', ['a', 'b']]]);
    expect(findMissingArgs(req, 't', { a: 0, b: '' })).toEqual([]);
  });
});
