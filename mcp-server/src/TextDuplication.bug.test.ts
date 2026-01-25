/**
 * Bug reproduction test for text duplication issue
 *
 * Bug report:
 * - update_paragraph_text, replace_text are appending text instead of replacing
 * - After save → reload, text appears duplicated
 * - Memory shows correct value, but file has duplicated content
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

describe('Text Duplication Bug', () => {
  let doc: HwpxDocument;
  const testOutputDir = path.join(__dirname, '..', 'test-output');
  const testFilePath = path.join(testOutputDir, 'text-duplication-test.hwpx');

  // Create a test HWPX with multiple runs per paragraph
  async function createTestHwpx(): Promise<Buffer> {
    const zip = new JSZip();

    const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

    // Section with multiple runs in a paragraph (simulates real HWPX documents)
    const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="para1">
    <hp:run charPrIDRef="0">
      <hp:t>Original text in run 0</hp:t>
    </hp:run>
  </hp:p>
  <hp:p id="para2">
    <hp:run charPrIDRef="0">
      <hp:t>Part 1 of paragraph 2</hp:t>
    </hp:run>
    <hp:run charPrIDRef="1">
      <hp:t> - Part 2 of paragraph 2</hp:t>
    </hp:run>
    <hp:run charPrIDRef="2">
      <hp:t> - Part 3 of paragraph 2</hp:t>
    </hp:run>
  </hp:p>
  <hp:p id="para3">
    <hp:run charPrIDRef="0">
      <hp:t> ◦ Original bullet point text</hp:t>
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
    // Ensure test output directory exists
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }

    const buffer = await createTestHwpx();
    doc = await HwpxDocument.createFromBuffer('test', testFilePath, buffer);
  });

  afterEach(() => {
    // Cleanup
    if (fs.existsSync(testFilePath)) {
      fs.unlinkSync(testFilePath);
    }
  });

  describe('Single run paragraph update', () => {
    it('should replace text in run 0, not append', async () => {
      // Initial state
      const initialPara = doc.getParagraph(0, 0);
      expect(initialPara?.text).toBe('Original text in run 0');

      // Update
      doc.updateParagraphText(0, 0, 0, 'New text for run 0');

      // Memory should show new text
      const memoryPara = doc.getParagraph(0, 0);
      expect(memoryPara?.text).toBe('New text for run 0');

      // Save
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Reload
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedPara = reloadedDoc.getParagraph(0, 0);

      // Should NOT be duplicated
      expect(reloadedPara?.text).toBe('New text for run 0');
      expect(reloadedPara?.text).not.toContain('Original');
    });

    it('should handle multiple updates to same paragraph', async () => {
      // First update
      doc.updateParagraphText(0, 0, 0, 'First update');

      // Second update
      doc.updateParagraphText(0, 0, 0, 'Second update');

      // Memory should show latest
      expect(doc.getParagraph(0, 0)?.text).toBe('Second update');

      // Save and reload
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedPara = reloadedDoc.getParagraph(0, 0);

      // Should show final value only
      expect(reloadedPara?.text).toBe('Second update');
      expect(reloadedPara?.text).not.toContain('First update');
      expect(reloadedPara?.text).not.toContain('Original');
    });
  });

  describe('Multi-run paragraph update', () => {
    it('should replace text in run 0 and clear other runs', async () => {
      // Paragraph 1 (index 1) has 3 runs
      const initialPara = doc.getParagraph(0, 1);
      console.log('Initial multi-run para:', initialPara?.text);
      expect(initialPara?.text).toContain('Part 1');
      expect(initialPara?.text).toContain('Part 2');
      expect(initialPara?.text).toContain('Part 3');

      // Update run 0 (should clear runs 1 and 2)
      doc.updateParagraphText(0, 1, 0, 'Completely new text');

      // Memory check
      const memoryPara = doc.getParagraph(0, 1);
      console.log('Memory after update:', memoryPara?.text);
      expect(memoryPara?.text).toBe('Completely new text');
      expect(memoryPara?.text).not.toContain('Part 2');
      expect(memoryPara?.text).not.toContain('Part 3');

      // Save and reload
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedPara = reloadedDoc.getParagraph(0, 1);

      console.log('Reloaded para:', reloadedPara?.text);

      // Should NOT contain old content
      expect(reloadedPara?.text).toBe('Completely new text');
      expect(reloadedPara?.text).not.toContain('Part 1');
      expect(reloadedPara?.text).not.toContain('Part 2');
      expect(reloadedPara?.text).not.toContain('Part 3');
    });
  });

  describe('updateParagraphTextPreserveStyles', () => {
    it('should replace text while preserving style structure', async () => {
      // Paragraph 1 has 3 runs
      const initialPara = doc.getParagraph(0, 1);
      console.log('Initial:', initialPara?.text);

      // Update with preserve styles
      const success = doc.updateParagraphTextPreserveStyles(0, 1, 'New styled text');
      expect(success).toBe(true);

      // Memory check
      const memoryPara = doc.getParagraph(0, 1);
      console.log('Memory:', memoryPara?.text);
      expect(memoryPara?.text).toBe('New styled text');

      // Save and reload
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedPara = reloadedDoc.getParagraph(0, 1);

      console.log('Reloaded:', reloadedPara?.text);

      // Should NOT be duplicated
      expect(reloadedPara?.text).toBe('New styled text');
      expect(reloadedPara?.text).not.toContain('Part 1');
    });
  });

  describe('replaceText', () => {
    it('should replace text globally, not append', async () => {
      // Initial
      const para2 = doc.getParagraph(0, 2);
      console.log('Initial para2:', para2?.text);
      expect(para2?.text).toContain('Original bullet point text');

      // Replace
      const count = doc.replaceText('Original bullet point text', 'Replaced bullet text');
      expect(count).toBe(1);

      // Memory check
      const memoryPara = doc.getParagraph(0, 2);
      console.log('Memory para2:', memoryPara?.text);
      expect(memoryPara?.text).toContain('Replaced bullet text');
      expect(memoryPara?.text).not.toContain('Original bullet point text');

      // Save and reload
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      const reloadedDoc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedPara = reloadedDoc.getParagraph(0, 2);

      console.log('Reloaded para2:', reloadedPara?.text);

      // Should be replaced, not duplicated
      expect(reloadedPara?.text).toContain('Replaced bullet text');
      expect(reloadedPara?.text).not.toContain('Original bullet point text');
    });
  });

  describe('XML inspection', () => {
    it('should show XML before and after save', async () => {
      // Get original XML
      const zip = await JSZip.loadAsync(await createTestHwpx());
      const originalXml = await zip.file('Contents/section0.xml')?.async('string');
      console.log('\n=== Original XML ===\n', originalXml);

      // Update
      doc.updateParagraphText(0, 0, 0, 'UPDATED TEXT');

      // Save
      const savedBuffer = await doc.save();

      // Get saved XML
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
      console.log('\n=== Saved XML ===\n', savedXml);

      // Check for duplication in XML
      const duplicateCheck = (savedXml?.match(/UPDATED TEXT/g) || []).length;
      console.log('Occurrences of "UPDATED TEXT":', duplicateCheck);

      expect(duplicateCheck).toBe(1); // Should appear exactly once

      // Check for original text
      expect(savedXml).not.toContain('Original text in run 0');
    });
  });
});
