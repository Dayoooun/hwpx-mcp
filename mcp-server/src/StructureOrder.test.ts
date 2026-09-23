/**
 * 문서 구조(문단·표 순서)와 셀 여백 회귀 테스트.
 *
 * 모든 시나리오는 한/글로 렌더해서 찾은 증상이다. "메모리 조회는 정상인데
 * 저장본만 틀어진다"는 유형이라, 판정은 반드시 저장된 XML 로 한다.
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

async function sectionXml(buf: Buffer): Promise<string> {
  const file = (await JSZip.loadAsync(buf)).file('Contents/section0.xml');
  if (!file) throw new Error('section0.xml missing');
  return file.async('string');
}

/** 저장본 최상위 문단을 순서대로: 표를 품은 문단은 'T', 나머지는 글자. */
function topLevel(xml: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (const m of xml.matchAll(/<(\/?)hp:p\b[^>]*?(\/?)>/g)) {
    if (m[1]) {
      depth--;
      if (depth === 0) {
        const block = xml.slice(start, m.index! + m[0].length);
        if (/<hp:tbl\b/.test(block)) out.push('T');
        else out.push([...block.matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(t => t[1]).join(''));
      }
    } else if (!m[2]) {
      if (depth === 0) start = m.index!;
      depth++;
    }
  }
  return out;
}

async function reopen(buf: Buffer): Promise<HwpxDocument> {
  return HwpxDocument.createFromBuffer('r', 'r.hwpx', buf);
}

describe('문단·표가 요청한 순서대로 저장된다', () => {
  it('새 문서에서 문단을 연달아 넣어도 순서가 뒤집히지 않는다', async () => {
    const doc = HwpxDocument.createNew('o1', 'order');
    expect(doc.insertParagraph(0, -1, 'P1')).toBe(0);
    expect(doc.insertParagraph(0, 0, 'P2')).toBe(1);

    // 실측: 수정 전 저장본은 [빈, P2, P1] — P1 이 맨 뒤로 갔다.
    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['P1', 'P2']);
  });

  it('문단 → 표 → 문단 순서가 저장본에서도 같다', async () => {
    const doc = HwpxDocument.createNew('o2', 'order');
    doc.insertParagraph(0, -1, '제목');
    doc.insertTable(0, 0, 1, 1);
    doc.insertParagraph(0, 1, '표 아래');

    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['제목', 'T', '표 아래']);
  });

  it('제목 → 표 → 제목 → 표 (한/글 렌더에서 틀어진 A 문서 구성)', async () => {
    const doc = HwpxDocument.createNew('o3', 'order');
    doc.insertParagraph(0, -1, 'A');
    doc.insertTable(0, 0, 2, 2);
    doc.insertParagraph(0, 1, 'A-2');
    doc.insertTable(0, 2, 2, 2);

    // 실측: 수정 전 저장본은 [A, A-2, T, T] 였다.
    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['A', 'T', 'A-2', 'T']);
  });

  it('메모리가 보여주는 순서와 저장본 순서가 같다', async () => {
    const doc = HwpxDocument.createNew('o4', 'order');
    doc.insertParagraph(0, -1, '하나');
    doc.insertParagraph(0, 0, '둘');
    doc.insertTable(0, 1, 1, 1);
    doc.insertParagraph(0, 2, '셋');

    const memory = doc.content.sections[0].elements
      .map(e => e.type === 'table' ? 'T' : e.data.runs.map((r: { text: string }) => r.text).join(''))
      .filter(t => t !== '');
    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(memory);
  });

  it('저장 → 다시 열기 → 표 뒤에 문단 추가해도 위치가 맞다', async () => {
    const seed = HwpxDocument.createNew('o5', 'order');
    seed.insertParagraph(0, -1, '위');
    seed.insertTable(0, 0, 1, 1);
    const doc = await reopen(await seed.save());

    const elements = doc.content.sections[0].elements;
    const tableAt = elements.findIndex(e => e.type === 'table');
    expect(tableAt).toBeGreaterThanOrEqual(0);
    doc.insertParagraph(0, tableAt, '아래');

    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['위', 'T', '아래']);
  });
});

describe('다시 연 문서에서 복제한 문단을 고쳐도 원본은 그대로다', () => {
  it('저장 → 다시 열기 → 복제 → 복제본 수정', async () => {
    const seed = HwpxDocument.createNew('c1', 'copy');
    seed.insertParagraph(0, -1, 'A');
    seed.insertParagraph(0, 0, 'B');
    const doc = await reopen(await seed.save());

    const paragraphs = doc.getParagraphs(0);
    const a = paragraphs.findIndex(p => p.text === 'A');
    expect(a).toBeGreaterThanOrEqual(0);
    expect(doc.copyParagraph(0, a, 0, a)).toBe(true);
    doc.updateParagraphText(0, a + 1, 0, 'A-복제본');

    // 실측: 수정 전 저장본은 [A-복제본, A, B] — 원본 자리에 새 글이 들어갔다.
    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['A', 'A-복제본', 'B']);
  });

  it('다시 열기 → 앞에 문단 삽입 → 뒤 문단 수정', async () => {
    const seed = HwpxDocument.createNew('c2', 'insert-then-edit');
    seed.insertParagraph(0, -1, 'A');
    seed.insertParagraph(0, 0, 'B');
    const doc = await reopen(await seed.save());

    const b = doc.getParagraphs(0).findIndex(p => p.text === 'B');
    doc.insertParagraph(0, -1, '맨 앞');
    doc.updateParagraphText(0, b + 1, 0, 'B-수정');

    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['맨 앞', 'A', 'B-수정']);
  });

  it('다시 열기 → 문단 삭제 → 뒤 문단 수정', async () => {
    const seed = HwpxDocument.createNew('c3', 'delete-then-edit');
    seed.insertParagraph(0, -1, 'A');
    seed.insertParagraph(0, 0, 'B');
    seed.insertParagraph(0, 1, 'C');
    const doc = await reopen(await seed.save());

    const texts = () => doc.getParagraphs(0).map(p => p.text);
    const a = texts().indexOf('A');
    expect(doc.deleteParagraph(0, a)).toBe(true);
    const c = texts().indexOf('C');
    doc.updateParagraphText(0, c, 0, 'C-수정');

    const saved = topLevel(await sectionXml(await doc.save())).filter(t => t !== '');
    expect(saved).toEqual(['B', 'C-수정']);
  });
});

describe('표 안 글자가 셀 테두리에 붙지 않는다', () => {
  it('새로 만든 표는 셀 여백이 실제로 적용되도록 저장된다', async () => {
    const doc = HwpxDocument.createNew('m1', 'margin');
    doc.insertTable(0, 0, 2, 2);
    const xml = await sectionXml(await doc.save());

    const table = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>'));
    const inMargin = table.match(/<hp:inMargin left="(\d+)" right="(\d+)" top="(\d+)" bottom="(\d+)"/)!;
    const cells = [...table.matchAll(/<hp:tc\b[^>]*\bhasMargin="(\d)"/g)].map(m => m[1]);

    // 한/글은 hasMargin="0" 인 셀에 표의 inMargin 을 쓴다. 실측: 수정 전 inMargin 0 이라
    // cellMargin 141 이 적혀 있어도 글자가 셀 선에서 0pt 였다.
    // 한/글 저장 원본 표 1,868개 중 inMargin 0 은 322개(17%)뿐이고 가장 흔한 값은 510/510/141/141.
    expect(cells.every(v => v === '0')).toBe(true);
    const [, left, right, top, bottom] = inMargin.map(Number);
    expect(left).toBeGreaterThan(0);
    expect(right).toBeGreaterThan(0);
    expect(top).toBeGreaterThan(0);
    expect(bottom).toBeGreaterThan(0);
  });
});

describe('행 삽입 시 cell_texts 는 칸마다 하나씩 들어간다', () => {
  it('템플릿 행에 여러 줄 셀이 있어도 새 행의 각 칸에 한 값씩', async () => {
    const doc = HwpxDocument.createNew('r1', 'row-texts');
    doc.insertTable(0, 0, 1, 3);
    doc.updateTableCell(0, 0, 0, 0, '원본 1');
    doc.updateTableCell(0, 0, 0, 1, '○ 첫째\n○ 둘째\n- 셋째'); // 문단 3개
    doc.updateTableCell(0, 0, 0, 2, '원본 3');
    doc.insertTableRow(0, 0, 0, ['새1', '새2', '새3']);

    const xml = await sectionXml(await doc.save());
    const rows = xml.slice(xml.indexOf('<hp:tbl')).split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    const cellTexts = rows[1]
      .split('</hp:tc>')
      .filter(c => c.includes('<hp:tc'))
      .map(c => [...c.matchAll(/<hp:t>([^<]*)<\/hp:t>/g)].map(m => m[1]).join(''));

    // 실측: 수정 전 ["새1", "새2새3", ""] — 빈 <hp:t> 를 칸 구분 없이 순서대로 채워
    // 둘째 칸의 문단 세 개에 새2·새3 이 들어가고 셋째 칸은 비었다.
    expect(cellTexts).toEqual(['새1', '새2', '새3']);
  });

  it('새 행의 셀은 템플릿 셀의 추가 문단을 물려받지 않는다', async () => {
    const doc = HwpxDocument.createNew('r2', 'row-paras');
    doc.insertTable(0, 0, 1, 1);
    doc.updateTableCell(0, 0, 0, 0, '가\n나\n다');
    doc.insertTableRow(0, 0, 0, ['새']);

    const xml = await sectionXml(await doc.save());
    const rows = xml.slice(xml.indexOf('<hp:tbl')).split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    // 빈 문단 두 개가 남으면 한/글이 셀 높이를 세 줄로 잡아 글자가 위로 몰려 보인다.
    expect((rows[1].match(/<hp:p\b/g) || []).length).toBe(1);
  });
});
