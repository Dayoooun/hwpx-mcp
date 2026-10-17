/**
 * export_to_pdf 가 열린 문서를 직렬화해도 문서 상태는 그대로다 (docs/specs/export-pdf.md AC-6).
 * save() 는 대기 편집을 zip 에 반영하고 비운다. 내보내기 뒤 undo 하고 저장하면, undo 가
 * 메모리만 되돌려서 내보낼 때 반영된 편집이 파일에 남았다(CodeRabbit #26).
 */
import { describe, it, expect } from 'vitest';
import { HwpxDocument } from '../../src/HwpxDocument';

async function blank(): Promise<HwpxDocument> {
  const base = await HwpxDocument.createNew('seed', 't', 'c').save();
  return HwpxDocument.createFromBuffer('d', '/tmp/x.hwpx', base);
}

async function paragraphs(buf: Buffer): Promise<string[]> {
  const doc = await HwpxDocument.createFromBuffer('r', '/tmp/r.hwpx', buf);
  return doc.content.sections[0].elements
    .filter(e => e.type === 'paragraph')
    .map(e => (e.data as { runs?: { text?: string }[] }).runs?.map(r => r.text ?? '').join('') ?? '');
}

describe('serializeForExport', () => {
  it('export-pdf#AC-6 내보낸 편집을 undo 하고 저장하면 그 편집은 파일에 없다', async () => {
    const doc = await blank();
    doc.insertParagraph(0, -1, '내보내기 직전 편집');
    expect(await paragraphs(await doc.serializeForExport())).toContain('내보내기 직전 편집');
    doc.undo();
    expect(await paragraphs(await doc.save())).not.toContain('내보내기 직전 편집');
  });

  it('export-pdf#AC-6 내보낸 뒤에도 저장 안 함 상태이고, 저장하면 내보낸 것과 같은 내용이다', async () => {
    const doc = await blank();
    doc.insertParagraph(0, -1, '첫째');
    doc.insertParagraph(0, 0, '둘째');
    const exported = await paragraphs(await doc.serializeForExport());
    expect(doc.isDirty).toBe(true);
    doc.insertParagraph(0, 1, '셋째');
    const saved = await paragraphs(await doc.save());
    expect(saved).toEqual([...exported, '셋째']);
    expect(doc.isDirty).toBe(false);
  });

  it('export-pdf#AC-6 대기 중인 표 편집(순서가 결과를 바꾸는)도 저장과 같은 내용으로 나온다', async () => {
    const build = async () => {
      const base = await blank();
      base.insertTable(0, 0, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('t', '/tmp/t.hwpx', await base.save());
      const t = doc.content.sections[0].elements.filter(e => e.type === 'table').length - 1;
      doc.updateTableCell(0, t, 0, 0, '전');
      doc.insertTableRow(0, t, 0, ['새1', '새2']);
      doc.updateTableCell(0, t, 1, 1, '후');
      doc.insertTableColumn(0, t, 0);
      return doc;
    };
    // 표 id 는 불러올 때마다 새로 정해진다(저장 두 번도 다르다) — id 를 빼고 비교한다.
    const shape = (d: HwpxDocument) => JSON.stringify(d.content.sections, (k, v) => (k === 'id' ? undefined : v));
    const exportedDoc = await build();
    const viaExport = await HwpxDocument.createFromBuffer('e', '/tmp/e.hwpx', await exportedDoc.serializeForExport());
    const viaSave = await HwpxDocument.createFromBuffer('s', '/tmp/s.hwpx', await (await build()).save());
    expect(shape(viaExport)).toContain('새1');
    expect(shape(viaExport)).toBe(shape(viaSave));
    const after = await HwpxDocument.createFromBuffer('a', '/tmp/a.hwpx', await exportedDoc.save());
    expect(shape(after)).toBe(shape(viaSave));
  });
});
