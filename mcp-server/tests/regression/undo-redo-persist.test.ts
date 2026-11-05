/**
 * 회귀: undo/redo 뒤 저장하면 남아야 할 수정이 파일에서 빠지던 문제 (2026-10-14 발견)
 *
 * 0.3.5 의 undo/redo 는 메모리 모델만 되돌리고 저장 대기 목록(_pending*)을 통째로 비웠다.
 * 수정은 그 목록을 거쳐야만 XML 에 쓰이므로, 화면(get_table_cell)에는 값이 보이는데
 * 저장본에는 없었다.
 *   - 수정 A, 수정 B, undo, 저장 → A 도 B 도 파일에 없음 (A 는 남아야 한다)
 *   - 수정, undo, redo, 저장 → 수정이 파일에 없음
 */
import { describe, it, expect } from 'vitest';
import { HwpxDocument } from '../../src/HwpxDocument';
import { roundTrip, cellText, paragraphIndexOf } from '../helpers/hwpx';

const texts = (doc: HwpxDocument) => doc.getParagraphs(0).map(p => p.text);

/** 저장·다시 연 2×2 표 문서 — 표가 이미 XML 에 있는 상태(사용자가 연 양식과 같다). */
async function tableDoc(): Promise<HwpxDocument> {
  const seed = HwpxDocument.createNew('u', 'undo');
  seed.insertParagraph(0, -1, '머리');
  seed.insertTable(0, 0, 2, 2);
  for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) seed.updateTableCell(0, 0, r, c, `원래${r}${c}`);
  return (await roundTrip(seed)).doc;
}

describe('undo 뒤 저장', () => {
  it('두 번 고치고 한 번 되돌리면 첫 수정은 파일에 남는다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, '수정A');
    doc.updateTableCell(0, 0, 1, 1, '수정B');
    expect(doc.undo()).toBe(true);

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 0)).toBe('수정A');
    expect(cellText(back, 0, 0, 1, 1)).toBe('원래11');
  });

  it('하나뿐인 수정을 되돌리면 파일은 원래 값 그대로다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 1, '지울 값');
    doc.undo();

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 1)).toBe('원래01');
  });

  it('문단 넣기를 되돌려도 앞선 표 수정은 남는다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 1, 0, '표 수정');
    doc.insertParagraph(0, 0, '되돌릴 문단');
    doc.undo();

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 1, 0)).toBe('표 수정');
    expect(texts(back)).not.toContain('되돌릴 문단');
  });
});

describe('문단 글 바꾸기(update_paragraph_text)', () => {
  it('두 문단을 고치고 하나를 되돌리면 나머지는 파일에 남는다', async () => {
    const seed = HwpxDocument.createNew('p', 'undo');
    seed.insertParagraph(0, -1, '첫 문단');
    seed.insertParagraph(0, 0, '둘째 문단');
    const doc = (await roundTrip(seed)).doc;
    doc.updateParagraphText(0, paragraphIndexOf(doc, 0, '첫 문단'), 0, '첫 수정');
    doc.updateParagraphText(0, paragraphIndexOf(doc, 0, '둘째 문단'), 0, '둘째 수정');
    doc.undo();

    const { doc: back } = await roundTrip(doc);
    expect(texts(back)).toContain('첫 수정');
    expect(texts(back)).toContain('둘째 문단');
    expect(texts(back)).not.toContain('둘째 수정');
  });

  it('되돌렸다가 다시 한 문단 수정이 파일에 남는다', async () => {
    const seed = HwpxDocument.createNew('p2', 'undo');
    seed.insertParagraph(0, -1, '원래 문단');
    const doc = (await roundTrip(seed)).doc;
    doc.updateParagraphText(0, paragraphIndexOf(doc, 0, '원래 문단'), 0, '다시 한 문단');
    doc.undo();
    doc.redo();

    const { doc: back } = await roundTrip(doc);
    expect(texts(back)).toContain('다시 한 문단');
  });
});

describe('redo 뒤 저장', () => {
  it('되돌렸다가 다시 한 수정이 파일에 남는다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, '다시 한 값');
    doc.undo();
    expect(doc.redo()).toBe(true);
    expect(cellText(doc, 0, 0, 0, 0)).toBe('다시 한 값');

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 0)).toBe('다시 한 값');
  });

  it('여러 수정을 모두 되돌렸다가 모두 다시 하면 전부 남는다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, 'A');
    doc.updateTableCell(0, 0, 0, 1, 'B');
    doc.insertParagraph(0, 0, '새 문단');
    doc.undo(); doc.undo(); doc.undo();
    doc.redo(); doc.redo(); doc.redo();

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 0)).toBe('A');
    expect(cellText(back, 0, 0, 0, 1)).toBe('B');
    expect(paragraphIndexOf(back, 0, '새 문단')).toBeGreaterThanOrEqual(0);
  });
});

describe('되돌리기 기록은 뒤의 수정에 바뀌지 않는다', () => {
  // 대기 항목을 기록과 문서가 같이 쓰면, 뒤의 수정이 항목을 그 자리에서 고칠 때
  // 기록까지 바뀌어 undo 뒤 메모리와 저장본이 달라졌다(Codex 검증 2026-10-14 재현).
  it('내어쓰기를 두 번 바꾸고 한 번 되돌리면 첫 값이 저장된다', async () => {
    const seed = HwpxDocument.createNew('h', 'undo');
    seed.insertParagraph(0, -1, '내어쓰기 문단');
    const doc = (await roundTrip(seed)).doc;
    const at = paragraphIndexOf(doc, 0, '내어쓰기 문단');
    expect(doc.setHangingIndent(0, at, 10)).toBe(true);
    expect(doc.setHangingIndent(0, at, 20)).toBe(true);
    doc.undo();
    expect(doc.getHangingIndent(0, at)).toBe(10);

    const { doc: back } = await roundTrip(doc);
    expect(back.getHangingIndent(0, paragraphIndexOf(back, 0, '내어쓰기 문단'))).toBe(10);
  });

  it('표 수정 뒤 앞에 구역을 넣었다가 되돌리면 표 수정은 원래 구역에 저장된다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, '남을 값');
    doc.insertSection(-1);
    doc.undo();
    expect(cellText(doc, 0, 0, 0, 0)).toBe('남을 값');

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 0)).toBe('남을 값');
  });
});

describe('저장은 되돌리기 경계', () => {
  it('저장 뒤에는 저장 전 수정을 되돌릴 수 없다(파일과 메모리가 어긋나지 않게)', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, '저장된 값');
    await doc.save();
    expect(doc.canUndo()).toBe(false);
    expect(doc.undo()).toBe(false);
    expect(cellText(doc, 0, 0, 0, 0)).toBe('저장된 값');
  });

  it('저장 뒤의 수정은 평소처럼 되돌린다', async () => {
    const doc = await tableDoc();
    doc.updateTableCell(0, 0, 0, 0, '저장된 값');
    await doc.save();
    doc.updateTableCell(0, 0, 1, 1, '저장 뒤 값');
    doc.undo();

    const { doc: back } = await roundTrip(doc);
    expect(cellText(back, 0, 0, 0, 0)).toBe('저장된 값');
    expect(cellText(back, 0, 0, 1, 1)).toBe('원래11');
  });
});
