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
    const xml = await sectionXml(await doc.save());
    expect(rowAddrsOfFirstTable(xml)).toEqual(['0', '1', '2', '3', '4']);
    expect(assertBalanced(xml)).toEqual({});
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
