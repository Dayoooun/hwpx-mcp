/**
 * Complex bug reproduction test for text duplication issue
 *
 * Simulates the exact user scenario:
 * 1. Open document
 * 2. Update paragraph
 * 3. Save
 * 4. Close and reopen (new session)
 * 5. Update same paragraph again
 * 6. Save
 * 7. Check for duplication
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

describe('Complex Text Duplication Bug', () => {
  const testOutputDir = path.join(__dirname, '..', 'test-output');
  const testFilePath = path.join(testOutputDir, 'complex-duplication-test.hwpx');

  async function createComplexTestHwpx(): Promise<Buffer> {
    const zip = new JSZip();

    const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Complex Test</hh:title>
  </hh:docInfo>
</hh:head>`;

    // More complex section with lineseg and various elements
    const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="800" spacing="160" horzpos="0" horzsize="14000" flags="393216"/>
    <hp:run charPrIDRef="0">
      <hp:t> ◦ 대표자 핵심 역량: 원본 텍스트</hp:t>
    </hp:run>
  </hp:p>
  <hp:p id="para1">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="800" spacing="160" horzpos="0" horzsize="14000" flags="393216"/>
    <hp:run charPrIDRef="0">
      <hp:t>   - 세부 내용 1</hp:t>
    </hp:run>
    <hp:run charPrIDRef="1">
      <hp:t>세부 내용 2</hp:t>
    </hp:run>
  </hp:p>
  <hp:p id="para2">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="800" spacing="160" horzpos="0" horzsize="14000" flags="393216"/>
    <hp:run charPrIDRef="0">
      <hp:t> ◦ 빈 불릿</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

    zip.file('Contents/header.xml', headerXml);
    zip.file('Contents/section0.xml', sectionXml);
    zip.file('[Content_Types].xml', contentTypes);

    return await zip.generateAsync({ type: 'nodebuffer' });
  }

  beforeEach(async () => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterEach(() => {
    if (fs.existsSync(testFilePath)) {
      fs.unlinkSync(testFilePath);
    }
  });

  describe('Multi-session update scenario', () => {
    it('should not duplicate text across multiple save/reload cycles', async () => {
      // === Session 1: Initial update ===
      console.log('\n=== Session 1 ===');
      const buffer1 = await createComplexTestHwpx();
      const doc1 = await HwpxDocument.createFromBuffer('session1', testFilePath, buffer1);

      const initial = doc1.getParagraph(0, 0);
      console.log('Initial text:', initial?.text);

      // First update
      doc1.updateParagraphText(0, 0, 0, ' ◦ 첫 번째 업데이트');

      const afterUpdate1 = doc1.getParagraph(0, 0);
      console.log('After update 1:', afterUpdate1?.text);

      // Save session 1
      const savedBuffer1 = await doc1.save();
      fs.writeFileSync(testFilePath, savedBuffer1);

      // Check saved XML
      const savedZip1 = await JSZip.loadAsync(savedBuffer1);
      const savedXml1 = await savedZip1.file('Contents/section0.xml')?.async('string');
      console.log('Session 1 saved XML para0:', extractPara(savedXml1!, 'para0'));

      // === Session 2: Reopen and update again ===
      console.log('\n=== Session 2 ===');
      const reloadedBuffer = fs.readFileSync(testFilePath);
      const doc2 = await HwpxDocument.createFromBuffer('session2', testFilePath, reloadedBuffer);

      const reloaded = doc2.getParagraph(0, 0);
      console.log('Reloaded text:', reloaded?.text);

      // CRITICAL CHECK: Should not be duplicated
      expect(reloaded?.text).toBe(' ◦ 첫 번째 업데이트');
      expect(reloaded?.text).not.toContain('원본');

      // Second update
      doc2.updateParagraphText(0, 0, 0, ' ◦ 두 번째 업데이트');

      const afterUpdate2 = doc2.getParagraph(0, 0);
      console.log('After update 2:', afterUpdate2?.text);

      // Save session 2
      const savedBuffer2 = await doc2.save();
      fs.writeFileSync(testFilePath, savedBuffer2);

      // Check saved XML
      const savedZip2 = await JSZip.loadAsync(savedBuffer2);
      const savedXml2 = await savedZip2.file('Contents/section0.xml')?.async('string');
      console.log('Session 2 saved XML para0:', extractPara(savedXml2!, 'para0'));

      // === Session 3: Final verification ===
      console.log('\n=== Session 3 (Final) ===');
      const finalBuffer = fs.readFileSync(testFilePath);
      const doc3 = await HwpxDocument.createFromBuffer('session3', testFilePath, finalBuffer);

      const final = doc3.getParagraph(0, 0);
      console.log('Final text:', final?.text);

      // CRITICAL CHECKS
      expect(final?.text).toBe(' ◦ 두 번째 업데이트');
      expect(final?.text).not.toContain('첫 번째');
      expect(final?.text).not.toContain('원본');

      // Check for duplication pattern (same text appearing twice)
      const occurrences = (final?.text?.match(/업데이트/g) || []).length;
      expect(occurrences).toBe(1);
    });

    it('should handle multiple updates in single session without duplication', async () => {
      console.log('\n=== Multiple updates in single session ===');
      const buffer = await createComplexTestHwpx();
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial:', doc.getParagraph(0, 0)?.text);

      // Multiple updates WITHOUT saving between them
      doc.updateParagraphText(0, 0, 0, ' ◦ Update A');
      console.log('After A:', doc.getParagraph(0, 0)?.text);

      doc.updateParagraphText(0, 0, 0, ' ◦ Update B');
      console.log('After B:', doc.getParagraph(0, 0)?.text);

      doc.updateParagraphText(0, 0, 0, ' ◦ Update C');
      console.log('After C:', doc.getParagraph(0, 0)?.text);

      // Now save once
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Check XML
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML para0:', extractPara(savedXml!, 'para0'));

      // Reload and verify
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const final = reloadedDoc.getParagraph(0, 0);
      console.log('Final after reload:', final?.text);

      expect(final?.text).toBe(' ◦ Update C');
      expect(final?.text).not.toContain('Update A');
      expect(final?.text).not.toContain('Update B');
      expect(final?.text).not.toContain('원본');
    });

    it('should handle replace_text followed by update_paragraph_text', async () => {
      console.log('\n=== replace_text + update_paragraph_text ===');
      const buffer = await createComplexTestHwpx();
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial para0:', doc.getParagraph(0, 0)?.text);
      console.log('Initial para2:', doc.getParagraph(0, 2)?.text);

      // First: replace_text (global)
      const count = doc.replaceText('빈 불릿', '교체된 불릿');
      console.log('Replaced count:', count);
      console.log('After replace para2:', doc.getParagraph(0, 2)?.text);

      // Then: update_paragraph_text on different paragraph
      doc.updateParagraphText(0, 0, 0, ' ◦ 업데이트된 텍스트');
      console.log('After update para0:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Check XML
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:\n', savedXml);

      // Reload and verify
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);

      const para0 = reloadedDoc.getParagraph(0, 0);
      const para2 = reloadedDoc.getParagraph(0, 2);

      console.log('Final para0:', para0?.text);
      console.log('Final para2:', para2?.text);

      expect(para0?.text).toBe(' ◦ 업데이트된 텍스트');
      expect(para0?.text).not.toContain('원본');

      expect(para2?.text).toContain('교체된 불릿');
      expect(para2?.text).not.toContain('빈 불릿');
    });
  });

  describe('Edge cases from user report', () => {
    it('should handle updateParagraphTextPreserveStyles on multi-run paragraph', async () => {
      console.log('\n=== updateParagraphTextPreserveStyles multi-run ===');
      const buffer = await createComplexTestHwpx();
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      // Para 1 has 2 runs
      const initial = doc.getParagraph(0, 1);
      console.log('Initial (2 runs):', initial?.text);
      console.log('Runs:', initial?.runs);

      // Update with preserve styles
      doc.updateParagraphTextPreserveStyles(0, 1, '새로운 내용으로 완전히 교체');
      console.log('After update:', doc.getParagraph(0, 1)?.text);

      // Save and reload
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Check XML
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML para1:', extractPara(savedXml!, 'para1'));

      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const final = reloadedDoc.getParagraph(0, 1);
      console.log('Final:', final?.text);

      expect(final?.text).toBe('새로운 내용으로 완전히 교체');
      expect(final?.text).not.toContain('세부 내용');
    });

    it('should handle empty paragraph update', async () => {
      console.log('\n=== Empty paragraph handling ===');
      const buffer = await createComplexTestHwpx();
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      // First make paragraph empty
      doc.updateParagraphText(0, 0, 0, '');
      console.log('After emptying:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer1 = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer1);

      // Reload
      const doc2 = await HwpxDocument.createFromBuffer('session2', testFilePath, savedBuffer1);
      console.log('After reload empty:', doc2.getParagraph(0, 0)?.text);

      // Update with new content
      doc2.updateParagraphText(0, 0, 0, ' ◦ 새로운 내용');
      console.log('After new content:', doc2.getParagraph(0, 0)?.text);

      // Save again
      const savedBuffer2 = await doc2.save();

      // Final check
      const doc3 = await HwpxDocument.createFromBuffer('session3', testFilePath, savedBuffer2);
      const final = doc3.getParagraph(0, 0);
      console.log('Final:', final?.text);

      expect(final?.text).toBe(' ◦ 새로운 내용');
    });
  });
});

// Helper function to extract a specific paragraph from XML
function extractPara(xml: string, paraId: string): string {
  const regex = new RegExp(`<hp:p[^>]*id="${paraId}"[^>]*>[\\s\\S]*?<\\/hp:p>`);
  const match = xml.match(regex);
  return match ? match[0] : 'NOT FOUND';
}
