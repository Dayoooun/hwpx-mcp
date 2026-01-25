/**
 * Bug test for unusual run structures that might cause text duplication
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

describe('Run Structure Edge Cases', () => {
  const testOutputDir = path.join(__dirname, '..', 'test-output');
  const testFilePath = path.join(testOutputDir, 'run-structure-test.hwpx');

  beforeEach(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterEach(() => {
    if (fs.existsSync(testFilePath)) {
      fs.unlinkSync(testFilePath);
    }
  });

  // Test: Run without <hp:t> tags
  describe('Run without hp:t tags', () => {
    it('should handle run with only charPr element', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:run charPrIDRef="0">
      <hp:charPr fontRef="0"/>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial (no hp:t):', doc.getParagraph(0, 0)?.text);
      console.log('Initial runs:', doc.getParagraph(0, 0)?.runs);

      // Update
      doc.updateParagraphText(0, 0, 0, 'New text');
      console.log('After update:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:', savedXml);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      console.log('Reloaded:', reloadedDoc.getParagraph(0, 0)?.text);

      expect(reloadedDoc.getParagraph(0, 0)?.text).toBe('New text');
    });
  });

  // Test: Run with empty <hp:t> tag
  describe('Run with empty hp:t tag', () => {
    it('should handle run with empty hp:t', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial (empty hp:t):', doc.getParagraph(0, 0)?.text);

      // Update
      doc.updateParagraphText(0, 0, 0, 'Filled text');
      console.log('After update:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:', savedXml);

      // Check for duplication
      const tMatches = savedXml?.match(/<hp:t[^>]*>Filled text<\/hp:t>/g) || [];
      console.log('hp:t matches:', tMatches.length);
      expect(tMatches.length).toBe(1);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      expect(reloadedDoc.getParagraph(0, 0)?.text).toBe('Filled text');
    });
  });

  // Test: Run with self-closing hp:t tag
  describe('Run with self-closing hp:t', () => {
    it('should handle self-closing hp:t tag', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:run charPrIDRef="0">
      <hp:t/>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial (self-closing hp:t):', doc.getParagraph(0, 0)?.text);

      // Update
      doc.updateParagraphText(0, 0, 0, 'New content');
      console.log('After update:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:', savedXml);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      console.log('Reloaded:', reloadedDoc.getParagraph(0, 0)?.text);

      expect(reloadedDoc.getParagraph(0, 0)?.text).toBe('New content');
    });
  });

  // Test: Multiple hp:t tags in single run (malformed but might exist)
  describe('Multiple hp:t tags in single run', () => {
    it('should handle multiple hp:t tags properly', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:run charPrIDRef="0">
      <hp:t>Part 1</hp:t>
      <hp:t>Part 2</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial (multiple hp:t):', doc.getParagraph(0, 0)?.text);
      console.log('Initial runs:', doc.getParagraph(0, 0)?.runs);

      // Update
      doc.updateParagraphText(0, 0, 0, 'Single text');
      console.log('After update:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:', savedXml);

      // Check: should NOT duplicate the new text
      const occurrences = (savedXml?.match(/Single text/g) || []).length;
      console.log('Occurrences of "Single text":', occurrences);

      // This might be a bug - with multiple hp:t tags, they'd all get the same text
      expect(occurrences).toBe(1); // Should be exactly 1

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const finalText = reloadedDoc.getParagraph(0, 0)?.text;
      console.log('Reloaded:', finalText);

      // Should NOT have duplicated text
      expect(finalText).toBe('Single text');
      expect(finalText).not.toContain('Single textSingle text');
    });
  });

  // Test: Run with lineseg between runs
  describe('Run with lineseg elements', () => {
    it('should handle lineseg between text elements', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para0">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000"/>
    <hp:run charPrIDRef="0">
      <hp:t>Line 1 text</hp:t>
    </hp:run>
    <hp:lineseg textpos="12" vertpos="1000" vertsize="1000"/>
    <hp:run charPrIDRef="0">
      <hp:t>Line 2 text</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial:', doc.getParagraph(0, 0)?.text);
      console.log('Runs:', doc.getParagraph(0, 0)?.runs);

      // Update run 0
      doc.updateParagraphText(0, 0, 0, 'Updated line 1');
      console.log('After update run 0:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('Saved XML:', savedXml);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      console.log('Reloaded:', reloadedDoc.getParagraph(0, 0)?.text);

      // Should contain updated text
      expect(reloadedDoc.getParagraph(0, 0)?.text).toContain('Updated line 1');
      // Should NOT contain original line 1
      expect(reloadedDoc.getParagraph(0, 0)?.text).not.toContain('Line 1 text');
    });
  });

  // Test: Complex Korean HWPX structure similar to real documents
  describe('Complex Korean HWPX structure', () => {
    it('should handle real-world Korean document structure', async () => {
      const zip = new JSZip();
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hp:p id="0" paraPrIDRef="1" styleIDRef="0">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="880" baseline="800" spacing="160" horzpos="0" horzsize="14000" flags="393216"/>
    <hp:run charPrIDRef="2">
      <hp:secPr textDirection="HORIZONTAL" spaceColumns="1134"/>
      <hp:t> ◦ 대표자 핵심 역량: 원본 텍스트입니다</hp:t>
    </hp:run>
  </hp:p>
  <hp:p id="1" paraPrIDRef="1" styleIDRef="0">
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="880" baseline="800" spacing="160" horzpos="0" horzsize="14000" flags="393216"/>
    <hp:run charPrIDRef="2">
      <hp:t>   - 세부 내용</hp:t>
    </hp:run>
    <hp:run charPrIDRef="3">
      <hp:t>: 추가 설명 텍스트</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`;
      zip.file('Contents/header.xml', createMinimalHeader());
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('[Content_Types].xml', createContentTypes());

      const buffer = await zip.generateAsync({ type: 'nodebuffer' });
      const doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);

      console.log('Initial para 0:', doc.getParagraph(0, 0)?.text);
      console.log('Initial para 1:', doc.getParagraph(0, 1)?.text);

      // Update para 0
      doc.updateParagraphText(0, 0, 0, ' ◦ 새로운 핵심 역량 텍스트');
      console.log('After update para 0:', doc.getParagraph(0, 0)?.text);

      // Save
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('\n=== Saved XML ===\n', savedXml);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const para0 = reloadedDoc.getParagraph(0, 0);
      const para1 = reloadedDoc.getParagraph(0, 1);

      console.log('Reloaded para 0:', para0?.text);
      console.log('Reloaded para 1:', para1?.text);

      // Assertions
      expect(para0?.text).toBe(' ◦ 새로운 핵심 역량 텍스트');
      expect(para0?.text).not.toContain('원본');
      expect(para0?.text).not.toContain('새로운 핵심 역량 텍스트새로운 핵심 역량 텍스트');

      // Para 1 should be unchanged
      expect(para1?.text).toContain('세부 내용');
    });
  });
});

function createMinimalHeader(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Test</hh:title></hh:docInfo>
</hh:head>`;
}

function createContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;
}
