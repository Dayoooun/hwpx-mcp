/**
 * 표 셀에 그림이 있는 문서에서 글자·문단 모양을 바꾸면 그 문단에 들어가야 한다.
 * 메모리에는 그림이 'image' 요소로 따로 잡혀 XML 최상위 요소 수와 어긋난다.
 * 순번으로 XML 을 찾으면 그림 수만큼 아래 문단이 바뀌었다(예비창업패키지 양식, 그림 2장 → 2칸 아래).
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function formWithPictures(): Promise<HwpxDocument> {
  const doc = HwpxDocument.createNew('f', 'f');
  doc.insertTable(0, 0, 1, 2);
  doc.insertImageInCell(0, 0, 0, 0, { data: PNG, mimeType: 'image/png', width: 40, height: 30 });
  doc.insertImageInCell(0, 0, 0, 1, { data: PNG, mimeType: 'image/png', width: 40, height: 30 });
  const els = doc.getSerializableContent() as { sections: Array<{ elements: unknown[] }> };
  let last = els.sections[0].elements.length - 1;
  for (const t of ['첫째 문단', '둘째 문단', '셋째 문단', '넷째 문단']) last = doc.insertParagraph(0, last, t);
  const saved = await doc.save();
  return HwpxDocument.createFromBuffer('r', 'r.hwpx', saved);
}

function indexOf(doc: HwpxDocument, text: string): number {
  const c = doc.getSerializableContent() as { sections: Array<{ elements: Array<{ type: string; data: { runs?: Array<{ text?: string }> } }> }> };
  return c.sections[0].elements.findIndex(e => e.type === 'paragraph' && (e.data.runs ?? []).map(r => r.text ?? '').join('') === text);
}

async function sectionXml(doc: HwpxDocument): Promise<string> {
  const z = await JSZip.loadAsync(await doc.save());
  return z.file('Contents/section0.xml')!.async('string');
}

const attrOf = (xml: string, text: string, attr: RegExp) => {
  const at = xml.indexOf(`<hp:t>${text}</hp:t>`);
  const open = xml.lastIndexOf('<hp:p ', at);
  return xml.slice(open, at).match(attr)?.[1];
};

describe('셀 그림 뒤 문단의 모양 변경', () => {
  it('set_text_style 이 고른 문단의 글자 모양만 바꾼다', async () => {
    const doc = await formWithPictures();
    const before = await sectionXml(await formWithPictures());
    const run = /<hp:run charPrIDRef="(\d+)"/;
    doc.applyCharacterStyle(0, indexOf(doc, '둘째 문단'), 0, { fontName: '함초롬바탕', fontSize: 12 });
    const xml = await sectionXml(doc);
    expect(attrOf(xml, '둘째 문단', run)).not.toBe(attrOf(before, '둘째 문단', run));
    expect(attrOf(xml, '넷째 문단', run)).toBe(attrOf(before, '넷째 문단', run));
  });

  it('set_paragraph_style 이 고른 문단의 문단 모양만 바꾼다', async () => {
    const doc = await formWithPictures();
    const before = await sectionXml(await formWithPictures());
    const para = /paraPrIDRef="(\d+)"/;
    doc.applyParagraphStyle(0, indexOf(doc, '첫째 문단'), { align: 'center' });
    const xml = await sectionXml(doc);
    expect(attrOf(xml, '첫째 문단', para)).not.toBe(attrOf(before, '첫째 문단', para));
    expect(attrOf(xml, '셋째 문단', para)).toBe(attrOf(before, '셋째 문단', para));
  });
});
