/**
 * 저장 단계가 없는 편집 — src/unsavedTools.ts 가 MCP 도구를 막는 근거.
 *
 * 각 문서 API 를 부른 뒤 저장한 파일(ZIP 안 XML·hpf 전부)에서 넣은 글·도형을 찾는다.
 * 지금은 모두 없다. 어느 하나라도 파일에 남게 되면 이 시험이 실패한다 — 그때
 * unsavedTools.ts 목록에서 그 도구를 빼고 이 시험의 해당 줄도 지운다.
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from '../../src/HwpxDocument';
import { UNSAVED_TOOLS } from '../../src/unsavedTools';
import { roundTrip } from '../helpers/hwpx';

async function savedText(doc: HwpxDocument): Promise<string> {
  const zip = await JSZip.loadAsync(await doc.save());
  let all = '';
  for (const name of Object.keys(zip.files)) {
    if (/\.(xml|hpf)$/.test(name)) all += await zip.file(name)!.async('string');
  }
  return all;
}

const count = (text: string, marker: string) => text.split(marker).length - 1;

/** 도구 이름 → [문서 API 호출, 파일에서 찾을 표식] */
const cases: Record<string, [(d: HwpxDocument) => unknown, string]> = {
  insert_footnote: [d => d.insertFootnote(0, 1, 'MK_FOOT'), 'MK_FOOT'],
  insert_endnote: [d => d.insertEndnote(0, 1, 'MK_END'), 'MK_END'],
  insert_memo: [d => d.insertMemo(0, 1, 'MK_MEMO', 'a'), 'MK_MEMO'],
  insert_bookmark: [d => d.insertBookmark(0, 1, 'MK_BOOK'), 'MK_BOOK'],
  insert_hyperlink: [d => d.insertHyperlink(0, 1, 'https://example.com/MK_LINK', 'MK_LINKTEXT'), 'MK_LINK'],
  insert_equation: [d => d.insertEquation(0, 1, 'MK_EQ over 2'), 'MK_EQ'],
  insert_line: [d => d.insertLine(0, 0, 0, 200, 0), '<hp:line'],
  insert_rect: [d => d.insertRect(0, 0, 0, 100, 50, { fillColor: '#DCEBF5' }), '<hp:rect'],
  insert_ellipse: [d => d.insertEllipse(0, 50, 50, 40, 20), '<hp:ellipse'],
  set_column_def: [d => d.setColumnDef(0, 2, 10), 'colCount="2"'],
  set_document_metadata: [d => { d.setMetadata({ title: 'MK_TITLE' }); return d.getMetadata().title === 'MK_TITLE'; }, 'MK_TITLE'],
};

async function baseDoc(): Promise<HwpxDocument> {
  const seed = HwpxDocument.createNew('u', 'unsaved');
  for (let i = 0; i < 3; i++) seed.insertParagraph(0, i - 1, `문단 ${i}`);
  return (await roundTrip(seed)).doc;
}

describe('저장 단계가 없는 편집 (도구를 막는 근거)', () => {
  it('막는 도구 목록과 이 시험의 목록이 같다', () => {
    expect([...UNSAVED_TOOLS.keys()].sort()).toEqual(Object.keys(cases).sort());
  });

  for (const [tool, [edit, marker]] of Object.entries(cases)) {
    it(`${tool}: 메모리에는 반영되지만 저장 파일에는 남지 않는다`, async () => {
      const doc = await baseDoc();
      const before = count(await savedText(await baseDoc()), marker);
      expect(edit(doc)).toBeTruthy();
      expect(count(await savedText(doc), marker)).toBe(before);
    });
  }
});

describe('HWP(HWP 5 바이너리) 열기', () => {
  it('빈 문서로 열리지 않고 지원하지 않는다고 알린다', async () => {
    await expect(HwpxDocument.createFromBuffer('h', 'old.hwp', Buffer.from('not parsed')))
      .rejects.toThrow(/not supported/);
  });
});
