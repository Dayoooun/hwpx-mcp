/**
 * 모듈: 문서 API 를 한/글 원본에서 흔한 구조 위에서 부르고, 저장 → 다시 열기로 판정.
 *
 * 한/글 원본 파일은 출처·저작권이 제각각이라 저장소에 넣지 않는다. 대신 원본에서
 * 측정한 구조적 특징을 코드로 재현한다 (각 describe 의 주석에 측정 근거).
 * 실제 원본으로 돌리는 확인은 scripts/corpus-check.mjs (로컬 전용, CI 제외).
 */
import { describe, it, expect } from 'vitest';
import { HwpxDocument } from '../../src/HwpxDocument';
import { sectionXml, roundTrip, withSectionXml, assertBalanced, paragraphText, cellText, paragraphIndexOf } from '../helpers/hwpx';

/** 한/글은 여러 문단에 같은 id 를 쓴다 (325 섹션 중 209 섹션에서 id="0"/"2147483648" 반복). */
const allIdsZero = (x: string) => x.replace(/<hp:p id="[^"]*"/g, '<hp:p id="0"');

describe('같은 id 가 반복되는 문단', () => {
  it('중간 문단을 고치면 그 문단만 바뀐다', async () => {
    const seed = HwpxDocument.createNew('m1', 'ids');
    ['가', '나', '다', '라'].forEach((t, i) => seed.insertParagraph(0, i - 1, t));
    const doc = await withSectionXml(seed, allIdsZero);
    const idx = paragraphIndexOf(doc, 0, '다');

    doc.updateParagraphText(0, idx, 0, '다-수정');
    const { doc: back } = await roundTrip(doc);
    expect(back.getParagraphs(0).map(p => p.text).filter(Boolean)).toEqual(['가', '나', '다-수정', '라']);
  });
});

describe('표를 품은 문단', () => {
  /** 한/글 원본 60건 중 표를 품은 문단 대부분이 글자 run + 표 run 구조였다. */
  async function docWithTableHost() {
    const seed = HwpxDocument.createNew('m2', 'host');
    seed.insertParagraph(0, -1, '앞 문단');
    seed.insertTable(0, 0, 2, 2);
    seed.updateTableCell(0, 0, 0, 0, '칸 가');
    seed.updateTableCell(0, 0, 1, 1, '칸 라');
    seed.insertParagraph(0, 1, '뒤 문단');
    return withSectionXml(seed, x => allIdsZero(
      x.replace(/(<hp:p [^>]*><hp:run[^>]*>)(<hp:tbl)/, '$1<hp:t>표 제목</hp:t></hp:run><hp:run charPrIDRef="0">$2')));
  }

  it('preserve_styles 로 제목을 바꿔도 표와 칸 글자는 그대로다', async () => {
    const doc = await docWithTableHost();
    const host = doc.content.sections[0].elements.findIndex(
      e => e.type === 'paragraph' && e.data.runs.some((r: { text: string }) => r.text === '표 제목'));
    expect(doc.updateParagraphTextPreserveStyles(0, host, '새 제목')).toBe(true);

    const { buf, doc: back } = await roundTrip(doc);
    expect(assertBalanced(await sectionXml(buf))).toEqual({});
    expect(paragraphText(back, 0, host)).toBe('새 제목');
    expect(cellText(back, 0, 0, 0, 0)).toBe('칸 가');
    expect(cellText(back, 0, 0, 1, 1)).toBe('칸 라');
  });

  it('표 뒤 문단을 고치면 표 칸이 아니라 그 문단이 바뀐다', async () => {
    const doc = await docWithTableHost();
    const after = paragraphIndexOf(doc, 0, '뒤 문단');
    doc.updateParagraphText(0, after, 0, '뒤-수정');

    const { doc: back } = await roundTrip(doc);
    expect(back.getParagraphs(0).some(p => p.text === '뒤-수정')).toBe(true);
    expect(cellText(back, 0, 0, 0, 0)).toBe('칸 가');
  });
});

describe('머리말을 품은 첫 문단', () => {
  /** 한/글 원본 325 섹션 중 88 섹션에서 파서가 머리말·글상자 안 문단을 본문 문단으로 올린다. */
  it('머리말 뒤 본문 문단을 고쳐도 머리말은 그대로다', async () => {
    const seed = HwpxDocument.createNew('m3', 'header');
    seed.insertParagraph(0, -1, '본문 1');
    seed.insertParagraph(0, 0, '본문 2');
    const doc = await withSectionXml(seed, x => allIdsZero(x.replace(
      /(<hp:p id="[^"]*"[^>]*><hp:run[^>]*>)(<hp:t>본문 1<\/hp:t>)/,
      '$1<hp:ctrl><hp:header id="1" applyPageType="BOTH"><hp:subList id=""><hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>머리말</hp:t></hp:run></hp:p></hp:subList></hp:header></hp:ctrl>$2')));

    const target = paragraphIndexOf(doc, 0, '본문 2');
    doc.updateParagraphText(0, target, 0, '본문 2-수정');
    const xml = await sectionXml(await doc.save());
    expect(xml).toContain('<hp:t>머리말</hp:t>');
    expect(xml).toContain('<hp:t>본문 2-수정</hp:t>');
    expect(xml).toContain('<hp:t>본문 1</hp:t>');
  });
});

describe('구역이 둘 이상인 문서', () => {
  it('insert_section 뒤 저장해도 구역 수와 각 구역 내용이 유지된다', async () => {
    const doc = HwpxDocument.createNew('m4', 'sections');
    doc.insertParagraph(0, -1, '표지');
    doc.insertSection(0);
    doc.insertParagraph(1, 0, '본문');
    doc.insertTable(1, 1, 2, 2);
    doc.updateTableCell(1, 0, 0, 0, '본문 표');

    const { doc: back } = await roundTrip(doc);
    expect(back.content.sections).toHaveLength(2);
    expect(back.getParagraphs(0).map(p => p.text)).toContain('표지');
    expect(back.getParagraphs(1).map(p => p.text)).toContain('본문');
    expect(cellText(back, 1, 0, 0, 0)).toBe('본문 표');
  });

  it('맨 앞에 구역을 넣으면 기존 구역 파일이 한 칸씩 밀린다', async () => {
    const doc = HwpxDocument.createNew('m5', 'sections');
    doc.insertParagraph(0, -1, '원래 첫 구역');
    doc.insertSection(-1);
    doc.insertParagraph(0, 0, '새 첫 구역');

    const { doc: back } = await roundTrip(doc);
    expect(back.content.sections).toHaveLength(2);
    expect(back.getParagraphs(0).map(p => p.text)).toContain('새 첫 구역');
    expect(back.getParagraphs(1).map(p => p.text)).toContain('원래 첫 구역');
  });
});

describe('병합이 있는 표의 행·열 삽입', () => {
  it('가로 병합 행 아래에 행을 넣어도 병합과 rowAddr 가 맞다', async () => {
    const seed = HwpxDocument.createNew('m6', 'merge');
    seed.insertTable(0, 0, 3, 3);
    seed.mergeCells(0, 0, 0, 0, 0, 2);            // 첫 행 가로 병합
    const { doc } = await roundTrip(seed);
    doc.insertTableRow(0, 0, 0, ['새', '행', '값']);

    const xml = await sectionXml(await doc.save());
    expect(assertBalanced(xml)).toEqual({});
    const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    const rows = t.split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    expect(rows.map(r => [...new Set([...r.matchAll(/rowAddr="(\d+)"/g)].map(m => m[1]))].join('/'))).toEqual(['0', '1', '2', '3']);
    expect(rows[0]).toMatch(/colSpan="3"/);         // 병합 유지
  });

  it('병합이 있는 표에 열을 넣어도 병합 칸 폭이 걸친 칸 폭의 합과 같다', async () => {
    const seed = HwpxDocument.createNew('m7', 'merge');
    seed.insertTable(0, 0, 2, 3);
    seed.mergeCells(0, 0, 0, 0, 0, 1);            // (0,0)~(0,1)
    const { doc } = await roundTrip(seed);
    doc.insertTableColumn(0, 0, 2);

    const xml = await sectionXml(await doc.save());
    const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    const width = +t.match(/<hp:sz width="(\d+)"/)![1];
    const [row0, row1] = t.split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    const w0 = [...row0.matchAll(/<hp:cellSz width="(\d+)"/g)].map(m => +m[1]);
    const w1 = [...row1.matchAll(/<hp:cellSz width="(\d+)"/g)].map(m => +m[1]);
    expect(w0.reduce((a, b) => a + b, 0)).toBe(width);
    expect(w1.reduce((a, b) => a + b, 0)).toBe(width);
    expect(w0[0]).toBe(w1[0] + w1[1]);              // 병합 칸 = 아래 두 칸
  });
});
