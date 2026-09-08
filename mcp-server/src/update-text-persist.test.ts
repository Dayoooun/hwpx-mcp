import { describe, it, expect } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

describe('Update paragraph text persistence', () => {
  it('should persist paragraph text update after save and reload', async () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-text-persist-'));
    try {
      const testFile = path.join(directory, '사업계획서 양식.hwpx');
      const fixture = HwpxDocument.createNew('fixture');
      fixture.insertParagraph(0, 0, '앞 문단 보존');
      fixture.insertParagraph(0, 1, ' ◦ 수정 대상 항목');
      fixture.insertParagraph(0, 2, '뒤 문단 보존');
      fs.writeFileSync(testFile, await fixture.save());

      const doc = await HwpxDocument.createFromBuffer('test-id', testFile, fs.readFileSync(testFile));
      const paragraphs = doc.getParagraphs(0);
      const target = paragraphs.find(paragraph => paragraph.text.includes(' ◦ '));
      expect(target).toBeDefined();
      const newText = '수정 완료: 매출 <계획> & 검토';
      doc.updateParagraphText(0, target!.index, 0, newText);
      expect(doc.getParagraph(0, target!.index)?.text).toBe(newText);
      fs.writeFileSync(testFile, await doc.save());

      const reopened = await HwpxDocument.createFromBuffer('reopened', testFile, fs.readFileSync(testFile));
      expect(reopened.getParagraphs(0).map(paragraph => paragraph.text)).toEqual(
        paragraphs.map(paragraph => paragraph.index === target!.index ? newText : paragraph.text)
      );
    } finally {
      fs.rmSync(directory, { recursive: true, force: true });
    }
  });
});
