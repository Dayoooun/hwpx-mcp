/**
 * 회귀: 정창기 박사님 2차 회신 (2026-09-24, 0.3.3 사용 결과) — ①~⑥
 *
 * 각 테스트는 회신의 재현 순서를 그대로 옮겼고, 0.3.3 에서 실패하는 것을 먼저
 * 확인했다. 판정은 저장본(다시 열기 포함)으로 한다. 한/글이 파일을 여는지는
 * 여기서 볼 수 없으므로, 회신에서 "이걸 고치니 한/글이 열렸다"고 확인된 조건
 * (rowAddr 유일성 등)을 대리 지표로 쓴다.
 */
import { describe, it, expect } from 'vitest';
import { HwpxDocument } from '../../src/HwpxDocument';
import {
  sectionXml, roundTrip, withSectionXml, assertBalanced,
  rowAddrsOfFirstTable, paragraphText, cellText,
} from '../helpers/hwpx';

describe('① insert_table_row 뒤 한/글이 멈춤 — rowAddr 중복', () => {
  it('4×4 표 1행 뒤에 넣으면 모든 행의 rowAddr 가 0..n-1 로 유일하다', async () => {
    const doc = HwpxDocument.createNew('r1', 'row');
    doc.insertTable(0, 0, 4, 4);
    doc.insertTableRow(0, 0, 1, ['a', 'b', 'c', 'd']);

    const addrs = rowAddrsOfFirstTable(await sectionXml(await doc.save()));
    // 회신 실측: tr1=1 / tr2=2(새 행) / tr3=2(중복) / tr4=3. rowAddr 만 순번대로
    // 고치자 한/글 2020 이 정상으로 열었다.
    expect(addrs).toEqual(['0', '1', '2', '3', '4']);
  });

  it('다시 연 문서에서도 새 행 아래 행들이 한 칸씩 밀린다', async () => {
    const seed = HwpxDocument.createNew('r2', 'row');
    seed.insertTable(0, 0, 4, 4);
    const { doc } = await roundTrip(seed);
    doc.insertTableRow(0, 0, 0, ['x', 'y', 'z', 'w']);

    const { buf, doc: back } = await roundTrip(doc);
    expect(rowAddrsOfFirstTable(await sectionXml(buf))).toEqual(['0', '1', '2', '3', '4']);
    expect(cellText(back, 0, 0, 1, 0)).toBe('x');
  });

  it('세로 병합(rowSpan=2) 한가운데에 행을 넣으면 거부한다', async () => {
    const seed = HwpxDocument.createNew('r3', 'row');
    seed.insertTable(0, 0, 4, 3);
    seed.mergeCells(0, 0, 0, 0, 1, 0);             // (0,0)~(1,0)
    const { doc } = await roundTrip(seed);

    // 회신: after_row 0 에 넣으면 rowSpan=2 칸이 새 행으로 복제돼 병합 영역이 겹쳤다.
    expect(() => doc.insertTableRow(0, 0, 0, ['x', 'y', 'z'])).toThrow(/merge|병합/i);
  });

  it('병합 영역 아래(경계 밖)에는 넣을 수 있다', async () => {
    const seed = HwpxDocument.createNew('r4', 'row');
    seed.insertTable(0, 0, 4, 3);
    seed.mergeCells(0, 0, 0, 0, 1, 0);
    const { doc } = await roundTrip(seed);

    expect(doc.insertTableRow(0, 0, 1, ['x', 'y', 'z'])).toBe(true);
    const { buf, doc: back } = await roundTrip(doc);
    const xml = await sectionXml(buf);
    expect(rowAddrsOfFirstTable(xml)).toEqual(['0', '1', '2', '3', '4']);
    expect(assertBalanced(xml)).toEqual({});

    // 새 행이 열 0..2 를 빈틈없이 덮어야 한다. 템플릿 행(1)은 열 0 이 위 병합에
    // 덮여 <hp:tc> 가 없는데, 그 행을 칸 단위로 복제하면 새 행도 열 0 을 잃었다
    // (CodeRabbit 지적, colCnt=3 인데 새 행 실효 열 수 2).
    const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    const newRow = t.split('</hp:tr>').filter(r => r.includes('<hp:tr'))[2];
    const covered = [...newRow.matchAll(/<hp:cellAddr colAddr="(\d+)"[^/]*\/>\s*<hp:cellSpan colSpan="(\d+)"/g)]
      .flatMap(m => Array.from({ length: +m[2] }, (_, k) => +m[1] + k)).sort();
    expect(covered).toEqual([0, 1, 2]);
    expect([0, 1, 2].map(c => cellText(back, 0, 0, 2, c))).toEqual(['x', 'y', 'z']);
  });

  it('병합 아래 새 행은 메모리에서도 열 0..2 를 덮는다 (저장 전 조회와 저장본이 같다)', async () => {
    const seed = HwpxDocument.createNew('r5', 'row');
    seed.insertTable(0, 0, 4, 3);
    seed.mergeCells(0, 0, 0, 0, 1, 0);
    const { doc } = await roundTrip(seed);

    doc.insertTableRow(0, 0, 1, ['x', 'y', 'z']);
    const memoryRow = doc.findTable(0, 0)!.rows[2].cells;
    expect(memoryRow.map(c => [c.colAddr, c.colSpan])).toEqual([[0, 1], [1, 1], [2, 1]]);
    expect(memoryRow.map(c => c.paragraphs[0].runs[0].text)).toEqual(['x', 'y', 'z']);
  });

  it('주소가 <hp:tc colAddr rowAddr> 속성에만 있는 표도 새 행이 칸을 갖고 아래 행이 밀린다', async () => {
    // 한/글은 늘 <hp:cellAddr> 자식을 쓰지만(원본 209/209) 손으로 만든 파일은
    // 주소를 <hp:tc> 속성에 둔다. 격자 복제가 자식만 읽어 새 <hp:tr> 이 비었다.
    const cell = (c: number, r: number, t: string) =>
      `<hp:tc colAddr="${c}" rowAddr="${r}"><hp:subList><hp:p id="p${c}${r}"><hp:run><hp:t>${t}</hp:t></hp:run></hp:p></hp:subList></hp:tc>`;
    const table =
      `<hp:tbl id="t1" rowCnt="2" colCnt="2"><hp:tr>${cell(0, 0, 'A')}${cell(1, 0, 'B')}</hp:tr>` +
      `<hp:tr>${cell(0, 1, 'C')}${cell(1, 1, 'D')}</hp:tr></hp:tbl>`;
    const doc = await withSectionXml(HwpxDocument.createNew('r6', 'row'),
      xml => xml.replace(/<\/hs:sec>/, `${table}</hs:sec>`));

    doc.insertTableRow(0, 0, 0, ['n1', 'n2']);
    const { buf: out, doc: back } = await roundTrip(doc);
    const t = (await sectionXml(out)).match(/<hp:tbl id="t1"[\s\S]*?<\/hp:tbl>/)![0];
    const rows = t.split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    expect(rows.map(r => [...r.matchAll(/<hp:tc [^>]*rowAddr="(\d+)"/g)].map(m => m[1]))).toEqual([
      ['0', '0'], ['1', '1'], ['2', '2'],
    ]);
    expect(assertBalanced(t)).toEqual({});
    const reopened = back.getTableMap().find(m => m.rows === 3)!;
    expect(reopened).toBeDefined();
    const ti = reopened.table_index_in_section;
    expect([0, 1].map(c => cellText(back, 0, ti, 1, c))).toEqual(['n1', 'n2']);
    expect([0, 1].map(c => cellText(back, 0, ti, 2, c))).toEqual(['C', 'D']);
  });
});

describe('② 표를 품은 문단을 preserve_styles 로 고치면 글이 사라지거나 표 칸이 바뀜', () => {
  /** "제1장 …" 글자와 목차 표를 한 문단에 품은 문서 (연구보고서 구조). */
  async function tocDoc(): Promise<{ doc: HwpxDocument; heading: number }> {
    const seed = HwpxDocument.createNew('t', 'toc');
    seed.insertParagraph(0, -1, '표지');
    seed.insertTable(0, 0, 3, 2);
    for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) seed.updateTableCell(0, 0, r, c, `목차${r}${c}`);
    seed.insertParagraph(0, 1, '본문');
    const doc = await withSectionXml(seed, x => x
      .replace(/(<hp:p [^>]*><hp:run[^>]*>)(<hp:tbl)/, '$1<hp:t>제1장 연구의 개요(작성중)1</hp:t></hp:run><hp:run charPrIDRef="0">$2')
      .replace(/<hp:p id="[^"]*"/g, '<hp:p id="0"'));   // 한/글처럼 id 반복
    const heading = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && paragraphTextOf(e).includes('제1장'));
    return { doc, heading };
  }
  const paragraphTextOf = (e: { data: { runs: Array<{ text: string }> } }) => e.data.runs.map(r => r.text).join('');

  it('제목 글만 바뀌고 목차 표와 그 칸 글자는 그대로다', async () => {
    const { doc, heading } = await tocDoc();
    expect(heading).toBeGreaterThanOrEqual(0);
    expect(doc.updateParagraphTextPreserveStyles(0, heading, '시험 문구')).toBe(true);

    const { buf, doc: back } = await roundTrip(doc);
    const xml = await sectionXml(buf);
    expect(assertBalanced(xml)).toEqual({});
    expect(paragraphText(back, 0, heading)).toBe('시험 문구');
    for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) expect(cellText(back, 0, 0, r, c)).toBe(`목차${r}${c}`);
  });

  it('표 요소를 가리킨 update_paragraph_text 는 성공이라 하지 않는다', async () => {
    const seed = HwpxDocument.createNew('t2', 'table-el');
    seed.insertTable(0, 0, 2, 2);
    const { doc } = await roundTrip(seed);
    const t = doc.content.sections[0].elements.findIndex(e => e.type === 'table');

    // 회신 곁가지: "Paragraph updated" 라 답하고 아무것도 안 바꿨다.
    // preserve_styles 는 "not found or no runs" 로 정직하게 거부했다 — 그쪽으로 맞춘다.
    expect(() => doc.updateParagraphText(0, t, 0, '바꿈')).toThrow(/paragraph/i);
  });
});

describe('③ table_index 기준이 도구마다 다름 (구역 2개 이상)', () => {
  async function twoSectionDoc(): Promise<HwpxDocument> {
    const d = HwpxDocument.createNew('s', 'sections');
    d.insertTable(0, 0, 2, 2);                   // 표지 구역의 표
    d.insertSection(0);
    d.insertTable(1, -1, 3, 3);                  // 본문 구역의 표
    return d;
  }

  it('get_table_map 이 구역 안 순번(table_index_in_section)을 함께 준다', async () => {
    const map = (await twoSectionDoc()).getTableMap();
    const body = map.find(m => m.section_index === 1)!;
    expect(body.table_index_in_section).toBe(0);
    expect(body.table_index).toBe(1);            // 문서 전체 순번은 그대로
  });

  it('맵이 준 구역 안 순번으로 쓰면 그 표에 들어간다', async () => {
    const doc = await twoSectionDoc();
    const body = doc.getTableMap().find(m => m.section_index === 1)!;
    expect(doc.updateTableCell(1, body.table_index_in_section, 1, 1, '본문 표')).toBe(true);
    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 1, 0, 1, 1)).toBe('본문 표');
    expect(cellText(back, 0, 0, 1, 1)).toBe('');  // 표지 표는 그대로
  });
});

describe('④ insert_table_column 뒤 표가 본문 폭을 넘음', () => {
  it('열을 넣어도 표 전체 폭은 그대로이고 칸 폭 합이 표 폭과 같다', async () => {
    const doc = HwpxDocument.createNew('c', 'col');
    doc.insertTable(0, 0, 4, 4);
    const before = await sectionXml(await doc.save());
    const tableWidth = +before.match(/<hp:tbl\b[\s\S]*?<hp:sz width="(\d+)"/)![1];

    doc.insertTableColumn(0, 0, 1);
    const xml = await sectionXml(await doc.save());
    const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    expect(+t.match(/<hp:sz width="(\d+)"/)![1]).toBe(tableWidth);
    for (const row of t.split('</hp:tr>').filter(r => r.includes('<hp:tr'))) {
      const widths = [...row.matchAll(/<hp:cellSz width="(\d+)"/g)].map(m => +m[1]);
      expect(widths).toHaveLength(5);
      // 회신 실측: 칸 폭 11765 가 그대로 하나 늘어 합계 58825 > 본문 폭 51024.
      expect(widths.reduce((a, b) => a + b, 0)).toBe(tableWidth);
    }
  });
});

describe('⑤ 글자 모양이 섞인 문단을 통째로 바꾸면 뒤쪽이 굵게 바뀜', () => {
  it('update_paragraph_text(run 0) 은 첫 run 의 글자 모양으로 전체를 쓴다', async () => {
    const seed = HwpxDocument.createNew('m', 'mixed');
    seed.insertParagraph(0, -1, 'placeholder');
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      '<hp:run charPrIDRef="0"><hp:t>보통 앞부분 </hp:t></hp:run><hp:run charPrIDRef="7"><hp:t>굵은 뒷부분</hp:t></hp:run>'));
    const p = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text.includes('보통')));

    doc.updateParagraphText(0, p, 0, '전체를 새 문장으로 바꿉니다');
    const xml = await sectionXml(await doc.save());
    const para = xml.slice(xml.lastIndexOf('<hp:p ', xml.indexOf('전체를')), xml.indexOf('</hp:p>', xml.indexOf('전체를')));
    // 새 글은 첫 run(보통) 한 곳에만 있고, 굵은 run 에는 글이 남지 않는다.
    const runs = [...para.matchAll(/<hp:run charPrIDRef="(\d+)">([\s\S]*?)<\/hp:run>/g)]
      .map(m => ({ char: m[1], text: [...m[2].matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(t => t[1]).join('') }))
      .filter(r => r.text);
    expect(runs).toEqual([{ char: '0', text: '전체를 새 문장으로 바꿉니다' }]);
  });
});

describe('② (나) 저장 검증이 깨진 XML 을 통과시킴 — 그림 칸 쓰기에서 실측', () => {
  /**
   * 한/글 원본 209건 첫 표 (0,0) 칸 쓰기 중 18건(보도자료 양식 등)의 저장본이
   * 파싱되지 않았다. 칸 첫 run 이 로고 그림 <hp:pic> 이고, 그 안의
   * <hc:transMatrix .../> 를 칸 글 정규식 <(hp|hs|hc):t[^>]*> 가 <hc:t> 로 읽어
   * "<hc:transMatrix …>시험</hc:t>" 을 만들었다. 0.3.3 의 저장 검증은 이걸 통과시켰다.
   */
  const picCell =
    '<hp:run charPrIDRef="0"><hp:pic id="1" zOrder="0"><hp:renderingInfo>' +
    '<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>' +
    '<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>' +
    '</hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/></hp:imgRect></hp:pic><hp:t/></hp:run>';

  async function logoTable(): Promise<HwpxDocument> {
    const seed = HwpxDocument.createNew('pic', 'logo');
    seed.insertTable(0, 0, 1, 2);
    seed.updateTableCell(0, 0, 0, 1, '제목칸');
    // (0,0) 칸의 run 을 한/글 원본과 같은 "그림 + 빈 글" run 으로 바꾼다.
    return withSectionXml(seed, x => x.replace(
      /(<hp:tc\b[\s\S]*?<hp:subList\b[\s\S]*?<hp:p\b[^>]*>)<hp:run\b[\s\S]*?<\/hp:run>/,
      (_m, open) => open + picCell));
  }

  it('그림이 든 칸에 글을 써도 저장본이 XML 로 파싱되고 그림 행렬은 그대로다', async () => {
    const doc = await logoTable();
    expect(doc.updateTableCell(0, 0, 0, 0, '시험')).toBe(true);

    const xml = await sectionXml(await doc.save());
    const { xmlWellFormednessError } = await import('../../src/XmlWellFormed');
    expect(xmlWellFormednessError(xml)).toBeNull();
    expect(xml).toContain('<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>');
    expect(xml).not.toMatch(/<hc:transMatrix[^>]*>시험/);
    expect(xml).toMatch(/<\/hp:pic><hp:t>시험<\/hp:t><\/hp:run>/);
  });
});
