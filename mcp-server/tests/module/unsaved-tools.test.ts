/**
 * 저장 단계가 없는 편집 — src/unsavedTools.ts 가 MCP 도구를 막는 근거.
 *
 * 같은 파일을 두 번 열어 한쪽에만 편집을 하고 둘 다 저장한다. 편집이 메모리에는
 * 들어갔는데(API 가 성공) 저장된 XML 이 편집하지 않은 쪽과 글자 하나 다르지 않으면
 * 그 편집은 파일에 닿지 않은 것이다. 어느 하나라도 파일에 남게 되면 이 시험이
 * 실패한다 — 그때 unsavedTools.ts 목록에서 그 도구를 빼고 이 시험의 해당 줄도 지운다.
 */
import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from '../../src/HwpxDocument';
import { UNSAVED_TOOLS } from '../../src/unsavedTools';
import { reopen, roundTrip, withSectionXml } from '../helpers/hwpx';

/** 저장본의 XML·hpf 전부 (그림 등 바이너리는 뺀다). */
async function savedXml(doc: HwpxDocument): Promise<string> {
  const zip = await JSZip.loadAsync(await doc.save());
  const names = Object.keys(zip.files).filter(n => /\.(xml|hpf)$/.test(n)).sort();
  let all = '';
  for (const name of names) all += `--- ${name}\n${await zip.file(name)!.async('string')}\n`;
  return all;
}

// 1×1 PNG
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const MEMO_ID = '777';
const MEMO =
  `<hp:ctrl><hp:fieldBegin id="${MEMO_ID}" type="MEMO" name="" editable="1" dirty="0" zorder="-1" fieldid="${MEMO_ID}">` +
  '<hp:parameters cnt="1" name=""><hp:stringParam name="Author">a</hp:stringParam></hp:parameters>' +
  '<hp:subList><hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">' +
  '<hp:run charPrIDRef="0"><hp:t>MK_MEMO</hp:t></hp:run></hp:p></hp:subList></hp:fieldBegin></hp:ctrl>';

/** 편집할 문서: 문단 셋, 그림 하나, 메모 하나가 이미 파일에 든 상태. */
let fixture: Buffer;
beforeAll(async () => {
  const seed = HwpxDocument.createNew('u', 'unsaved');
  for (let i = 0; i < 3; i++) seed.insertParagraph(0, i - 1, `문단 ${i}`);
  seed.insertImage(0, 2, { data: PNG, mimeType: 'image/png', width: 20, height: 20 });
  const withMemo = await withSectionXml((await roundTrip(seed)).doc, xml =>
    xml.replace('<hp:t>문단 1</hp:t>', `${MEMO}<hp:t>문단 1</hp:t><hp:ctrl><hp:fieldEnd beginIDRef="${MEMO_ID}" fieldid="${MEMO_ID}"/></hp:ctrl>`));
  fixture = await withMemo.save();
});

/** 도구 이름 → 문서 API 호출 (메모리에 들어갔는지까지 확인해 true 를 돌려준다) */
const cases: Record<string, (d: HwpxDocument) => boolean> = {
  insert_footnote: d => !!d.insertFootnote(0, 1, 'MK_FOOT'),
  insert_endnote: d => !!d.insertEndnote(0, 1, 'MK_END'),
  insert_memo: d => !!d.insertMemo(0, 1, 'MK_NEWMEMO', 'a'),
  delete_memo: d => d.getMemos().length === 1 && d.deleteMemo(MEMO_ID) && d.getMemos().length === 0,
  insert_bookmark: d => d.insertBookmark(0, 1, 'MK_BOOK'),
  insert_hyperlink: d => d.insertHyperlink(0, 1, 'https://example.com/MK_LINK', 'MK_LINKTEXT'),
  insert_equation: d => !!d.insertEquation(0, 1, 'MK_EQ over 2'),
  insert_line: d => !!d.insertLine(0, 0, 0, 200, 0),
  insert_rect: d => !!d.insertRect(0, 0, 0, 100, 50, { fillColor: '#DCEBF5' }),
  insert_ellipse: d => !!d.insertEllipse(0, 50, 50, 40, 20),
  update_image_size: d => {
    const img = d.getImages()[0];
    return !!img && d.updateImageSize(img.id, 40, 60) && d.getImages()[0].width === 40;
  },
  set_column_def: d => d.setColumnDef(0, 2, 10),
  set_page_settings: d => d.setPageSettings(0, { width: 300, height: 400, marginTop: 10 }) &&
    d.getPageSettings(0)?.width === 300,
  apply_style: d => {
    const style = d.getStyles()[0];
    const target = d.content.sections[0].elements[1];
    return !!style && target.type === 'paragraph' && d.applyStyle(0, 1, style.id) &&
      (target.data as { style?: number }).style === style.id;
  },
};

describe('저장 단계가 없는 편집 (도구를 막는 근거)', () => {
  it('막는 도구 목록과 이 시험의 목록이 같다', () => {
    expect([...UNSAVED_TOOLS.keys()].sort()).toEqual(Object.keys(cases).sort());
  });

  it('재료 문서에 메모와 그림이 들어 있다', async () => {
    const doc = await reopen(fixture);
    expect(doc.getMemos().map(m => m.id)).toEqual([MEMO_ID]);
    expect(doc.getImages()).toHaveLength(1);
  });

  for (const [tool, edit] of Object.entries(cases)) {
    it(`${tool}: 메모리에는 반영되지만 저장 파일은 편집하지 않은 것과 같다`, async () => {
      const untouched = await savedXml(await reopen(fixture));
      const doc = await reopen(fixture);
      expect(edit(doc)).toBe(true);
      expect(await savedXml(doc)).toBe(untouched);
    });
  }
});

describe('문서 정보(set_document_metadata)', () => {
  it('header.xml 에 자리가 없는 문서(새 문서·한컴 저장본)는 저장되지 않는다고 미리 알려 준다', async () => {
    const doc = await reopen(fixture);
    expect(await doc.unsavableMetadataFields(['title', 'creator'])).toEqual(['title', 'creator']);
  });

  it('header.xml 에 자리가 있는 필드는 저장된다', async () => {
    const zip = await JSZip.loadAsync(fixture);
    const header = (await zip.file('Contents/header.xml')!.async('string'))
      .replace(/(<hh:head\b[^>]*>)/, '$1<hh:title>old</hh:title>');
    zip.file('Contents/header.xml', header);
    const doc = await reopen(await zip.generateAsync({ type: 'nodebuffer' }));
    expect(await doc.unsavableMetadataFields(['title', 'creator'])).toEqual(['creator']);

    doc.setMetadata({ title: 'MK_TITLE' });
    expect(await savedXml(doc)).toContain('<hh:title>MK_TITLE</hh:title>');
  });
});

describe('HWP(HWP 5 바이너리) 열기', () => {
  it('빈 문서로 열리지 않고 지원하지 않는다고 알린다', async () => {
    await expect(HwpxDocument.createFromBuffer('h', 'old.hwp', Buffer.from('not parsed')))
      .rejects.toThrow(/not supported/);
  });
});
