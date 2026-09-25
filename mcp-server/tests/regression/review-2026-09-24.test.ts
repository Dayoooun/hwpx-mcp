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
import { error } from '../../src/ToolResult';
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

  it('표 뒤에도 자기 글("…(작성중)1" 의 "1")이 있는 문단을 바꿔도 표 칸 글자는 그대로다', async () => {
    // 연구보고서 목차 줄처럼 문단이 [제목 글][목차 표][쪽 번호 글] 순서다. 0.3.3 은 preserve_styles 가
    // 표 뒤 run 을 표 칸 안 run 으로 세어, 칸 글자 여섯 개가 첫 칸 하나로 합쳐지고 나머지가 비었다.
    const seed = HwpxDocument.createNew('t3', 'toc-tail');
    seed.insertParagraph(0, -1, '표지');
    seed.insertTable(0, 0, 3, 2);
    for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) seed.updateTableCell(0, 0, r, c, `목차${r}${c}`);
    seed.insertParagraph(0, 1, '본문');
    const doc = await withSectionXml(seed, x => x
      .replace(/(<hp:p [^>]*><hp:run[^>]*>)(<hp:tbl[\s\S]*?<\/hp:tbl>)/, (_m, open: string, tbl: string) =>
        `${open}<hp:t>제1장 연구의 개요(작성중)</hp:t></hp:run><hp:run charPrIDRef="0">${tbl}<hp:t>1</hp:t>`)
      .replace(/<hp:p id="[^"]*"/g, '<hp:p id="0"'));
    const heading = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && paragraphTextOf(e).includes('제1장'));
    expect(doc.updateParagraphTextPreserveStyles(0, heading, '시험 문구입니다')).toBe(true);

    const { buf, doc: back } = await roundTrip(doc);
    expect(assertBalanced(await sectionXml(buf))).toEqual({});
    expect(paragraphText(back, 0, heading)).toBe('시험 문구입니다');
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

describe('⑤ 뒤: 문단 통째 교체가 성공이라 답하고 글을 엉뚱한 곳에 둠 (0.3.5, 한/글 원본에서 실측)', () => {
  /**
   * 한/글 원본 150건 문단 11,754개 중 710개는 메모리 글과 그 문단 XML 의 "자기 글"
   * (findDirectChildRuns → ownRunText) 이 달랐고, 그런 문단에 update_paragraph_text 를
   * 걸면 100건 표본에서 13/13 이 제자리에 들어가지 않았다(같은 문단은 51/51 제자리).
   *   (가) 261개 — 자기 글이 없다. 글이 전부 글상자 안에 있다(파서가 글상자 글을 이
   *        문단 글로 올린다). 새 글은 글상자 뒤 빈 <hp:t> 에 들어가 화면에 안 보였다.
   *   (나) 449개 — 자기 글은 있는데 run 번호가 어긋난다. 파서는 <hp:t> 하나를 탭·고정폭
   *        빈칸 앞뒤로 여러 run 으로 쪼개므로, run 0 이 <hp:t> 조각 하나만 가리켰다.
   */
  it('(나) 고정폭 빈칸으로 시작하는 문단을 통째로 바꾸면 새 글만 남는다', async () => {
    const seed = HwpxDocument.createNew('fw', 'fwspace');
    seed.insertParagraph(0, -1, 'placeholder');
    // 한/글 보도자료 문단 머리 모양 (hwpx-h-02 실측): 빈칸 + 고정폭 빈칸 run, 그 뒤 본문 run
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      '<hp:run charPrIDRef="0"><hp:t> <hp:fwSpace/></hp:t></hp:run>' +
      '<hp:run charPrIDRef="0"><hp:t>2025년 2분기 해외직접투자액은</hp:t></hp:run>'));
    const p = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text.includes('2분기')));

    doc.updateParagraphText(0, p, 0, '새 문장');
    const { doc: back } = await roundTrip(doc);
    expect(paragraphText(back, 0, p)).toBe('새 문장');
  });

  it('(나) 탭이 든 문단을 통째로 바꾸면 탭 뒤 글이 남지 않는다', async () => {
    const seed = HwpxDocument.createNew('tab', 'tab');
    seed.insertParagraph(0, -1, 'placeholder');
    // 목차 줄 모양 (SO-SUEOP 실측): 한 <hp:t> 안에 글 · 탭 · 쪽 번호
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      '<hp:run charPrIDRef="0"><hp:t>20) 유예 오상원<hp:tab width="30284" leader="3" type="0"/>36</hp:t></hp:run>'));
    const p = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text.includes('유예')));

    doc.updateParagraphText(0, p, 0, '새 목차 줄');
    const { doc: back } = await roundTrip(doc);
    expect(paragraphText(back, 0, p)).toBe('새 목차 줄');
  });

  it('통째 교체 뒤 같은 문단의 다른 run 을 고쳐도 그 글이 저장본에 남는다 (CodeRabbit PR #17)', async () => {
    // "A<hp:tab/>B" 는 메모리에서 ["A", "", "B"] 다. run 0 통째 교체 뒤 run 2 를 고치면 메모리는
    // "newX" 인데, 저장 때 run 2 를 XML 노드로 다시 찾다가 못 찾아 X 가 빠졌다.
    const seed = HwpxDocument.createNew('aw', 'after-whole');
    seed.insertParagraph(0, -1, 'placeholder');
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      '<hp:run charPrIDRef="0"><hp:t>A<hp:tab width="1" leader="0" type="0"/>B</hp:t></hp:run>'));
    const p = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text === 'A'));

    doc.updateParagraphText(0, p, 0, 'new');
    doc.updateParagraphText(0, p, 2, 'X');
    expect(paragraphText(doc, 0, p)).toBe('newX');
    const { doc: back } = await roundTrip(doc);
    expect(paragraphText(back, 0, p)).toBe('newX');
  });

  it('통째 교체 앞의 run 편집은 통째 교체가 덮는다', async () => {
    const seed = HwpxDocument.createNew('bw', 'before-whole');
    seed.insertParagraph(0, -1, 'placeholder');
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      '<hp:run charPrIDRef="0"><hp:t>A</hp:t></hp:run><hp:run charPrIDRef="0"><hp:t>B</hp:t></hp:run>'));
    const p = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text === 'A'));

    doc.updateParagraphText(0, p, 1, 'Y');
    doc.updateParagraphText(0, p, 0, '마지막');
    const { doc: back } = await roundTrip(doc);
    expect(paragraphText(back, 0, p)).toBe('마지막');
  });

  it('(가) 글이 전부 글상자 안에 있는 문단은 빈 문단으로 읽고, 글상자 글은 그 속 문단으로 고친다', async () => {
    const seed = HwpxDocument.createNew('box', 'textbox');
    seed.insertParagraph(0, -1, '앞 문단');
    seed.insertParagraph(0, 0, 'placeholder');
    // 행정업무운영 편람 실측 모양: 문단의 자기 run 에는 글상자와 빈 <hp:t/> 뿐이고 글은 글상자 안
    const box =
      '<hp:rect id="1" zOrder="0"><hp:drawText lastWidth="1000" name="" editable="0"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP">' +
      '<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>글상자 안 제목</hp:t></hp:run></hp:p>' +
      '</hp:subList></hp:drawText></hp:rect>';
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      `<hp:run charPrIDRef="0">${box}<hp:t/></hp:run>`));
    const els = doc.content.sections[0].elements;
    const withBox = els.findIndex((e, i) => e.type === 'paragraph' && els[i + 1]?.type === 'rect');
    const inner = els.findIndex(e => e.type === 'paragraph' && paragraphText(doc, 0, els.indexOf(e)) === '글상자 안 제목');
    // 0.3.4 까지는 글상자를 품은 문단이 글상자 글을 자기 글로 읽었다. 고치면 글상자 옆 빈 자리에 찍혔다.
    expect(paragraphText(doc, 0, withBox)).toBe('');
    expect(inner).toBeGreaterThan(withBox);

    doc.updateParagraphText(0, inner, 0, '새 제목');
    const { doc: back, buf } = await roundTrip(doc);
    const xml = await sectionXml(buf);
    expect(xml).toContain('<hp:t>새 제목</hp:t></hp:run></hp:p></hp:subList>');
    expect(xml).not.toContain('글상자 안 제목');
    expect(paragraphText(back, 0, inner)).toBe('새 제목');
  });

  it('글상자 뒤에 자기 글이 있는 문단을 통째로 바꾸면 다시 열어도 새 글이다 (CodeRabbit PR #17)', async () => {
    // 파서가 문단 전체에서 run 을 읽으면 첫 </hp:run>(글상자 속 문단의 것)에서 run 이 끝나
    // 글상자 뒤 자기 글을 못 읽었다. 저장본에는 새 글이 있는데 다시 열면 글상자 글 "안" 이 나왔다.
    const seed = HwpxDocument.createNew('after-box', 'after-box');
    seed.insertParagraph(0, -1, 'placeholder');
    const box =
      '<hp:rect id="1" zOrder="0"><hp:drawText lastWidth="1000" name="" editable="0"><hp:subList id="">' +
      '<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>안</hp:t></hp:run></hp:p>' +
      '</hp:subList></hp:drawText></hp:rect>';
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      `<hp:run charPrIDRef="0">${box}<hp:t>바깥</hp:t></hp:run>`));
    const els = doc.content.sections[0].elements;
    const outer = els.findIndex((e, i) => e.type === 'paragraph' && els[i + 1]?.type === 'rect');
    expect(paragraphText(doc, 0, outer)).toBe('바깥');

    doc.updateParagraphText(0, outer, 0, '바깥 새 글');
    expect(paragraphText(doc, 0, outer)).toBe('바깥 새 글');
    const { doc: back } = await roundTrip(doc);
    expect(paragraphText(back, 0, outer)).toBe('바깥 새 글');
  });

  it('글상자 속 문단과 그 글상자를 품은 문단을 한 번에 고쳐도 둘 다 저장된다 (CodeRabbit PR #17)', async () => {
    // 파서는 글상자 속 문단을 본문 문단으로도 올린다(메모리: [빈, 바깥 문단, rect, 속 문단]). 두 문단의
    // XML 범위가 겹치고, 저장은 시작이 뒤인 속 문단을 먼저 쓴다. 속 글이 길어지면 미리 잰 바깥 문단의
    // 끝 위치가 문단 한가운데를 가리켜, 바깥 문단의 자기 글을 찾지 못하고 쓰기가 빠졌다.
    const seed = HwpxDocument.createNew('nest', 'nested');
    seed.insertParagraph(0, -1, 'placeholder');
    const box =
      '<hp:rect id="1" zOrder="0"><hp:drawText lastWidth="1000" name="" editable="0"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP">' +
      '<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>안</hp:t></hp:run></hp:p>' +
      '</hp:subList></hp:drawText></hp:rect>';
    const doc = await withSectionXml(seed, x => x.replace(
      /<hp:run charPrIDRef="0"><hp:t>placeholder<\/hp:t><\/hp:run>/,
      `<hp:run charPrIDRef="0">${box}<hp:t>바깥</hp:t></hp:run>`));
    const els = doc.content.sections[0].elements;
    const outer = els.findIndex((e, i) => e.type === 'paragraph' && els[i + 1]?.type === 'rect');
    const inner = els.findIndex((e, i) => i > outer && e.type === 'paragraph' && paragraphText(doc, 0, i) === '안');
    expect(outer).toBeGreaterThanOrEqual(0);
    expect(inner).toBeGreaterThan(outer);   // 바깥 문단이 rect 앞(시작이 먼저), 올려진 속 문단이 뒤

    const LONG = '글상자 안의 글이 훨씬 길어졌습니다. 바깥 문단의 끝 위치가 이만큼 밀립니다.';
    doc.updateParagraphText(0, inner, 0, LONG);
    doc.updateParagraphText(0, outer, 0, '바깥 새 글');
    const xml = await sectionXml(await doc.save());
    const { xmlWellFormednessError } = await import('../../src/XmlWellFormed');
    expect(xmlWellFormednessError(xml)).toBeNull();
    expect(xml).toContain(`<hp:t>${LONG}</hp:t></hp:run></hp:p></hp:subList>`);
    expect(xml).toContain('</hp:rect><hp:t>바깥 새 글</hp:t>');
    expect(xml).not.toContain('>바깥<');
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

describe('PR #16 CodeRabbit 2차 지적 (418b3d7)', () => {
  it('#1 열 삽입 뒤 행 삽입: 메모리 새 행이 새 열 칸까지 갖고 그 칸에 쓸 수 있다', async () => {
    // insertTableColumn 은 새 칸에 colAddr 를 두지 않는다. 격자가 colAddr 없는 칸을
    // 걸러 새 행이 원래 열 수만 덮었고, 새 열 칸 쓰기가 false 였다 (0.3.3 은 true).
    const doc = HwpxDocument.createNew('cr1', 'col-then-row');
    doc.insertTable(0, 0, 2, 2);
    doc.insertTableColumn(0, 0, 1);
    doc.insertTableRow(0, 0, 0, ['a', 'b', 'c']);

    const row = doc.findTable(0, 0)!.rows[1].cells;
    expect(row.map(c => c.paragraphs[0].runs[0].text)).toEqual(['a', 'b', 'c']);
    expect(doc.updateTableCell(0, 0, 1, 2, 'z')).toBe(true);
  });

  it('#1b 가운데 열을 삽입한 뒤 행 삽입: 새 행이 3칸이고 새 열 칸에 쓸 수 있다', async () => {
    // insertTableColumn 이 새 칸에 주소를 주지 않고 뒤 칸도 밀지 않아, 가운데(0열 뒤)에
    // 넣으면 행이 [0, 주소 없음, 1] 이 되고 격자가 열 1 을 두 번 셌다(CodeRabbit 3차).
    const doc = HwpxDocument.createNew('cr1b', 'mid-col-then-row');
    doc.insertTable(0, 0, 2, 2);
    doc.insertTableColumn(0, 0, 0);
    expect(doc.findTable(0, 0)!.rows[0].cells.map(c => c.colAddr)).toEqual([0, 1, 2]);

    doc.insertTableRow(0, 0, 0, ['a', 'b', 'c']);
    const row = doc.findTable(0, 0)!.rows[1].cells;
    expect(row.map(c => c.paragraphs[0].runs[0].text)).toEqual(['a', 'b', 'c']);
    expect(doc.updateTableCell(0, 0, 1, 2, 'z')).toBe(true);
  });

  it('#2 표 id 에 정규식 기호가 있어도 그 표의 칸이 바뀐다', async () => {
    const seed = HwpxDocument.createNew('cr2', 'regex-id');
    seed.insertTable(0, 0, 1, 1);
    seed.insertTable(0, 0, 1, 1);
    // 두 표의 id 를 "1.5" 와 "105" 로: 이스케이프하지 않으면 "1.5" 패턴이 "105" 에도 맞는다.
    let n = 0;
    const doc = await withSectionXml(seed, x => x.replace(/(<hp:tbl\b[^>]*\bid=")[^"]*"/g,
      (_m, open) => `${open}${n++ === 0 ? '105' : '1.5'}"`));
    const tables = doc.content.sections[0].elements.filter(e => e.type === 'table');
    const target = tables.findIndex(t => t.data.id === '1.5');
    expect(target).toBeGreaterThanOrEqual(0);

    doc.updateTableCell(0, target, 0, 0, '맞는 표');
    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, target, 0, 0)).toBe('맞는 표');
    expect(cellText(back, 0, 1 - target, 0, 0)).toBe('');
  });

  it('#3 위 행의 병합 칸을 빌려 좁힐 때 중첩 표가 아니라 그 칸의 colSpan·폭을 줄인다', () => {
    // 좁히기는 "템플릿 행에 열 0 칸이 없어 위 행의 병합 칸(colSpan 2)을 빌리는데,
    // 템플릿 행이 열 1 에서 따로 시작"할 때만 일어난다. 문서 API 로는 이 모양을
    // 만들기 어려워 격자 함수에 행 XML 을 직접 준다.
    //   행0: A(열0, colSpan 2, 폭 2000) — 칸 안에 중첩 표(그 칸 colSpan 7)
    //   행1: B(열1, 폭 1000)            — 열 0 은 위에서 덮였다고 가정
    const nested =
      '<hp:tbl id="n"><hp:tr><hp:tc><hp:subList><hp:p><hp:run><hp:t>안</hp:t></hp:run></hp:p></hp:subList>' +
      '<hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="7" rowSpan="1"/><hp:cellSz width="500" height="1"/></hp:tc></hp:tr></hp:tbl>';
    const tc = (inner: string, col: number, span: number, width: number) =>
      `<hp:tc><hp:subList><hp:p><hp:run>${inner}</hp:run></hp:p></hp:subList>` +
      `<hp:cellAddr colAddr="${col}" rowAddr="0"/><hp:cellSpan colSpan="${span}" rowSpan="1"/>` +
      `<hp:cellSz width="${width}" height="1"/></hp:tc>`;
    const rows = [
      { xml: `<hp:tr>${tc(nested, 0, 2, 2000)}</hp:tr>` },
      { xml: `<hp:tr>${tc('<hp:t>B</hp:t>', 1, 1, 1000)}</hp:tr>` },
    ];

    const doc = HwpxDocument.createNew('cr3', 'grid');
    const cells: string[] = (doc as any).gridCellsForNewRow(rows, 1);
    expect(cells).toHaveLength(2);
    const borrowed = cells[0];
    const tail = borrowed.slice(borrowed.lastIndexOf('</hp:subList>'));
    // 그 칸 자신은 colSpan 1, 폭 1000 (2000 × 1/2) 으로 좁혀진다.
    expect(tail).toMatch(/<hp:cellSpan colSpan="1"/);
    expect(tail).toMatch(/<hp:cellSz width="1000"/);
    // 중첩 표 칸은 그대로다. 예전 코드는 여기(처음 나오는 cellSpan)를 바꿨다.
    expect(borrowed).toContain('<hp:cellSpan colSpan="7" rowSpan="1"/><hp:cellSz width="500"');
  });

  it('#4 이름공간 선언이 빠진 section 은 저장 검증에서 거부된다', async () => {
    const { xmlWellFormednessError } = await import('../../src/XmlWellFormed');
    // xmlns:hp 만 있고 hs 는 선언이 없다 — 구문은 맞지만 한/글이 읽는 문서가 아니다.
    expect(xmlWellFormednessError('<?xml version="1.0"?><hs:sec xmlns:hp="urn:hp"><hp:p/></hs:sec>')).toMatch(/prefix|namespace/i);
    expect(xmlWellFormednessError('<?xml version="1.0"?><hs:sec xmlns:hs="urn:hs" xmlns:hp="urn:hp"><hp:p/></hs:sec>')).toBeNull();
  });
});

describe('표 편집은 호출한 순서대로 저장된다 (CodeRabbit 3차, 0.3.3 에도 있던 결함)', () => {
  /**
   * 저장이 표 편집을 호출 순서가 아니라 종류별 고정 순서(칸 쓰기 → … → 행 삽입 → 행 삭제
   * → 열 삽입 → 열 삭제)로 적용했다. 메모리와 저장본이 달라졌다 (0.3.3 에서도 6건 중 6건).
   */
  const txt = (c: { paragraphs?: Array<{ runs: Array<{ text: string }> }> }) =>
    (c.paragraphs ?? []).map(p => p.runs.map(r => r.text).join('')).join('');
  const grid = (d: HwpxDocument) => d.findTable(0, 0)!.rows.map(r => r.cells.map(txt));

  async function table(rows: number, cols: number) {
    const d = HwpxDocument.createNew('ord', 'order');
    d.insertTable(0, 0, rows, cols);
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) d.updateTableCell(0, 0, r, c, `${r}${c}`);
    return (await roundTrip(d)).doc;
  }
  async function savedEqualsMemory(d: HwpxDocument) {
    const mem = grid(d);
    const { buf, doc: back } = await roundTrip(d);
    expect(assertBalanced(await sectionXml(buf))).toEqual({});
    expect(grid(back)).toEqual(mem);
    return mem;
  }

  it('열 삽입 → 행 삽입(cell_texts) → 새 열 칸 쓰기: 새 행이 [a,b,c] 이고 z 가 남는다', async () => {
    const d = await table(2, 2);
    d.insertTableColumn(0, 0, 1);
    d.insertTableRow(0, 0, 0, ['a', 'b', 'c']);
    d.updateTableCell(0, 0, 1, 2, 'z');
    const mem = await savedEqualsMemory(d);
    expect(mem[1]).toEqual(['a', 'b', 'z']);
  });

  it('행 삽입 → 새 행에 쓰기: 쓴 글이 새 행에 있다', async () => {
    const d = await table(3, 2);
    d.insertTableRow(0, 0, 0);
    d.updateTableCell(0, 0, 1, 0, 'NEW');
    const mem = await savedEqualsMemory(d);
    expect(mem[1][0]).toBe('NEW');
    expect(mem[2]).toEqual(['10', '11']);
  });

  it('행 삭제를 두 번: 두 번째 번호는 첫 삭제 뒤의 번호다', async () => {
    const d = await table(3, 2);
    d.deleteTableRow(0, 0, 0);
    d.deleteTableRow(0, 0, 1);
    expect(await savedEqualsMemory(d)).toEqual([['10', '11']]);
  });

  it('행 삽입을 두 번: 두 번째 번호는 첫 삽입 뒤의 번호다', async () => {
    const d = await table(3, 2);
    d.insertTableRow(0, 0, 0, ['p', 'q']);
    d.insertTableRow(0, 0, 2, ['s', 't']);
    expect(await savedEqualsMemory(d)).toEqual([['00', '01'], ['p', 'q'], ['10', '11'], ['s', 't'], ['20', '21']]);
  });

  it('가운데 열 삽입 → 새 열 칸 쓰기: 새 열에 들어간다', async () => {
    const d = await table(3, 2);
    d.insertTableColumn(0, 0, 0);
    d.updateTableCell(0, 0, 0, 1, 'COL');
    expect((await savedEqualsMemory(d))[0]).toEqual(['00', 'COL', '01']);
  });

  it('행 삽입 → 그 아래 두 행 세로 병합: 병합이 새 번호의 3·4행에 걸린다', async () => {
    // 0.3.3 과 83465c8 은 병합을 행 삽입보다 먼저 적용해, 새 번호 (3,0)-(4,0) 을
    // 삽입 전 표에 걸었다. 4행이 없어 병합이 빠지고 아래 행 rowAddr 도 겹쳤다.
    const d = await table(4, 2);
    d.insertTableRow(0, 0, 0, ['n', 'n']);
    d.mergeCells(0, 0, 3, 0, 4, 0);
    const { buf, doc: back } = await roundTrip(d);
    expect(assertBalanced(await sectionXml(buf))).toEqual({});
    const shown = back.findTable(0, 0)!.rows.map(r => r.cells.map(c =>
      `${c.rowAddr},${c.colAddr}:${txt(c)}${(c.rowSpan ?? 1) > 1 ? `[${c.rowSpan}x${c.colSpan ?? 1}]` : ''}`));
    expect(shown).toEqual([
      ['0,0:00', '0,1:01'], ['1,0:n', '1,1:n'], ['2,0:10', '2,1:11'],
      ['3,0:20[2x1]', '3,1:21'], ['4,1:31'],
    ]);
  });

  it('열 삭제 → 칸 쓰기: 한 칸씩 당겨진 열 번호의 칸에 들어간다', async () => {
    // deleteTableColumn 이 메모리의 뒤쪽 칸 colAddr 를 줄이지 않았다. 저장 때 XML 은 먼저
    // 열을 지우고 주소를 당기므로, 옛 주소로 찾는 쓰기는 없는 칸을 찾아 버려졌다
    // (CodeRabbit 4차; 0.3.3 에서도 버려짐).
    const d = await table(2, 3);
    d.deleteTableColumn(0, 0, 0);
    d.updateTableCell(0, 0, 0, 1, 'Z');
    d.updateTableCell(0, 0, 1, 0, 'Y');
    expect(await savedEqualsMemory(d)).toEqual([['01', 'Z'], ['Y', '12']]);
  });

  it('열 삭제 → 행 삽입 → 새 행 쓰기: 새 행이 두 칸이고 쓴 글이 남는다', async () => {
    const d = await table(2, 3);
    d.deleteTableColumn(0, 0, 1);
    d.insertTableRow(0, 0, 0);
    d.updateTableCell(0, 0, 1, 1, 'Z');
    expect(await savedEqualsMemory(d)).toEqual([['00', '02'], ['', 'Z'], ['10', '12']]);
  });

  it('행 삭제 뒤 세로 병합 경계: 병합 끝 행 뒤 삽입은 되고 병합 가운데 삽입은 거부된다', async () => {
    // deleteTableRow 가 메모리 rowAddr 를 당기지 않아, 행 삽입의 병합 가르기 검사가
    // 한 행 어긋났다. 병합 아래 삽입은 거부하고 병합을 가르는 삽입은 통과시켰다.
    const merged = async () => {
      const seed = HwpxDocument.createNew('dm', 'delete-then-insert');
      seed.insertTable(0, 0, 4, 2);
      seed.mergeCells(0, 0, 1, 0, 2, 0);               // 1~2행 세로 병합
      const { doc } = await roundTrip(seed);
      doc.deleteTableRow(0, 0, 0);                      // 병합은 이제 0~1행
      return doc;
    };
    const below = await merged();
    expect(below.insertTableRow(0, 0, 1)).toBe(true);
    const through = await merged();
    expect(() => through.insertTableRow(0, 0, 0)).toThrow(/split the merged cell/);
  });

  describe('병합이 한 행의 칸을 모두 덮으면 거부한다 (한/글 2024 가 PDF 를 만들지 못함, 0.3.3 도 같음)', () => {
    // 저장본에 칸 없는 <hp:tr> 이 생겼다. 한/글 2024 는 이 파일에서 PDF 를 만들지 못했고
    // (전체 폭 두 행 병합, 1열 표 세로 병합), 같은 표를 전체 폭보다 좁게 병합하면 변환됐다.
    const tcPerRow = (xml: string) => {
      const tbl = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
      return [...tbl.matchAll(/<hp:tr(?:\s[^>]*)?>([\s\S]*?)<\/hp:tr>/g)].map(m => (m[1].match(/<hp:tc[\s>]/g) ?? []).length);
    };

    it('3×3 표 0·1행 전체 폭 병합은 거부하고 표는 그대로 남는다', async () => {
      const d = await table(3, 3);
      expect(() => d.mergeCells(0, 0, 0, 0, 1, 2)).toThrow(/row 1 would have no cell of its own/);
      const { buf } = await roundTrip(d);
      expect(tcPerRow(await sectionXml(buf))).toEqual([3, 3, 3]);
    });

    it('1열 표의 세로 병합은 거부한다', async () => {
      const d = await table(4, 1);
      expect(() => d.mergeCells(0, 0, 0, 0, 1, 0)).toThrow(/no cell of its own/);
    });

    it('열을 지워 1열이 된 표의 세로 병합도 거부한다', async () => {
      const d = await table(4, 2);
      d.deleteTableColumn(0, 0, 0);
      expect(() => d.mergeCells(0, 0, 0, 0, 1, 0)).toThrow(/no cell of its own/);
    });

    it('아래 행에 칸이 남는 병합은 그대로 되고 저장본의 모든 행에 칸이 있다', async () => {
      const d = await table(3, 3);
      expect(d.mergeCells(0, 0, 0, 0, 1, 1)).toBe(true);
      const { buf } = await roundTrip(d);
      expect(tcPerRow(await sectionXml(buf))).toEqual([2, 1, 3]);
    });
  });
});

describe('병합 뒤 칸 쓰기는 병합된 표의 그 칸에 들어간다 (호출 순서 적용 뒤 드러남)', () => {
  it('(0,0)-(0,1) 병합 뒤 (0,2) 에 쓰면 저장본의 열 2 칸에 들어간다', async () => {
    // 메모리 행은 덮인 칸을 남겨 [0,1,2] 이고 XML 행은 [0(2칸),2] 다. 칸 쓰기를 병합 뒤에
    // 적용하게 되면서, 메모리 위치(2)로 XML 칸을 고르면 없는 칸이 됐다. 열 주소로 고른다.
    const doc = HwpxDocument.createNew('mg', 'merge-then-write');
    doc.insertTable(0, 0, 2, 3);
    doc.mergeCells(0, 0, 0, 0, 0, 1);
    doc.updateTableCell(0, 0, 0, 0, 'A');
    doc.updateTableCell(0, 0, 0, 2, 'C');
    doc.updateTableCell(0, 0, 1, 1, 'E');

    const { buf, doc: back } = await roundTrip(doc);
    const xml = await sectionXml(buf);
    expect(assertBalanced(xml)).toEqual({});
    const t = back.findTable(0, 0)!;
    const txt = (c: { paragraphs: Array<{ runs: Array<{ text: string }> }> }) => c.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('');
    expect(t.rows[0].cells.map(txt)).toEqual(['A', 'C']);
    expect(t.rows[0].cells.map(c => c.colAddr)).toEqual([0, 2]);
    expect(t.rows[1].cells.map(txt)).toEqual(['', 'E', '']);
  });
});
describe('⑥ 실패한 호출이 isError: false 로 옴', () => {
  /** 회신 ⑥: 필수값 누락·가려진 칸 쓰기처럼 실패한 호출도 isError 가 false 였다(0.3.3). */
  it('도구 오류 결과는 isError: true 이고, 가려진 칸 쓰기는 오류로 막힌다', () => {
    const r = error('Missing required argument for insert_paragraph: section_index');
    expect(r.isError).toBe(true);
    expect(JSON.parse((r.content[0] as { text: string }).text)).toEqual({ error: 'Missing required argument for insert_paragraph: section_index' });

    const doc = HwpxDocument.createNew('r6', 'covered');
    doc.insertTable(0, 0, 2, 2);
    doc.mergeCells(0, 0, 0, 0, 0, 1);
    expect(() => doc.updateTableCell(0, 0, 0, 1, 'x')).toThrow(/covered by the merged cell/);
  });
});
