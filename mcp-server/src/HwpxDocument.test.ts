import { describe, it, expect, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * Create a minimal HWPX file for testing.
 * HWPX is a ZIP archive with XML files inside.
 */
async function createTestHwpxBuffer(): Promise<Buffer> {
  const zip = new JSZip();

  // Minimal header.xml
  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

  // Section with a table containing cells, plus multiple paragraphs with same marker text
  // This tests that updateParagraphText only affects target location, not all matching text
  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Hello World</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10">
            <hp:run><hp:t>Cell 0,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="11">
            <hp:run><hp:t></hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="20">
            <hp:run><hp:t>Cell 1,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="21">
            <hp:run><hp:t>Cell 1,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p id="2">
    <hp:run><hp:t>Hello World</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run><hp:t>Hello World</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

  // Content types
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', sectionXml);
  zip.file('[Content_Types].xml', contentTypesXml);

  return await zip.generateAsync({ type: 'nodebuffer' });
}

describe('HwpxDocument - Table Cell Update', () => {
  let testFilePath: string;

  beforeEach(async () => {
    // Create a test HWPX file
    const buffer = await createTestHwpxBuffer();
    testFilePath = path.join(__dirname, '..', 'test-temp.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should update table cell and persist after save', async () => {
    // 1. Open document
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // 2. Get initial cell value (getTable returns { rows, cols, data: [[{text, style}]] })
    const tableBefore = doc.getTable(0, 0);
    console.log('Table before:', JSON.stringify(tableBefore, null, 2));
    expect(tableBefore?.data[0][0].text).toBe('Cell 0,0');

    // 3. Update cell (0,0) with new text
    const updateResult = doc.updateTableCell(0, 0, 0, 0, 'Updated Text');
    expect(updateResult).toBe(true);

    // 4. Verify memory update
    const tableAfterMemory = doc.getTable(0, 0);
    console.log('Table after memory update:', JSON.stringify(tableAfterMemory, null, 2));
    expect(tableAfterMemory?.data[0][0].text).toBe('Updated Text');

    // 5. Save document
    const savedBuffer = await doc.save();
    fs.writeFileSync(testFilePath, savedBuffer);

    // 6. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedSection = await savedZip.file('Contents/section0.xml')?.async('string');
    console.log('Saved section XML:', savedSection);

    // 7. Reload document from saved buffer
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 8. Verify persisted value
    const tableAfterReload = doc2.getTable(0, 0);
    console.log('Table after reload:', JSON.stringify(tableAfterReload, null, 2));

    expect(tableAfterReload?.data[0][0].text).toBe('Updated Text');
  });

  it('should update empty cell and persist after save', async () => {
    // 1. Open document
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // 2. Verify cell (0,1) is initially empty
    const tableBefore = doc.getTable(0, 0);
    console.log('Empty cell before:', tableBefore?.data[0][1].text);
    expect(tableBefore?.data[0][1].text).toBe('');

    // 3. Update empty cell (0,1) with new text
    const updateResult = doc.updateTableCell(0, 0, 0, 1, 'New Text in Empty Cell');
    expect(updateResult).toBe(true);

    // 4. Verify memory update
    const tableAfterMemory = doc.getTable(0, 0);
    expect(tableAfterMemory?.data[0][1].text).toBe('New Text in Empty Cell');

    // 5. Save document
    const savedBuffer = await doc.save();

    // 6. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedSection = await savedZip.file('Contents/section0.xml')?.async('string');
    console.log('Saved section XML (empty cell test):', savedSection);

    // 7. Reload document
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 8. Verify persisted value
    const tableAfterReload = doc2.getTable(0, 0);
    console.log('Table after reload (empty cell):', JSON.stringify(tableAfterReload, null, 2));

    expect(tableAfterReload?.data[0][1].text).toBe('New Text in Empty Cell');
  });

  it('should preserve original XML structure after save', async () => {
    // 1. Load original XML
    const originalBuffer = fs.readFileSync(testFilePath);
    const originalZip = await JSZip.loadAsync(originalBuffer);
    const originalXml = await originalZip.file('Contents/section0.xml')?.async('string');

    // 2. Open, update, save
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, originalBuffer);
    doc.updateTableCell(0, 0, 0, 0, 'Modified');
    const savedBuffer = await doc.save();

    // 3. Check saved XML preserves structure
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    console.log('Original XML:', originalXml);
    console.log('Saved XML:', savedXml);

    // Key attributes should be preserved
    expect(savedXml).toContain('hp:tbl');
    expect(savedXml).toContain('id="100"');
    expect(savedXml).toContain('rowCnt="2"');
    expect(savedXml).toContain('hp:subList');
    expect(savedXml).toContain('Modified'); // New text
  });
});

describe('HwpxDocument - updateParagraphTextPreserveStyles', () => {
  let testFilePath: string;

  /**
   * Create a test HWPX with paragraphs containing multiple runs with different styles
   */
  async function createMultiRunTestHwpx(): Promise<Buffer> {
    const zip = new JSZip();

    const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Multi-Run Test</hh:title>
  </hh:docInfo>
</hh:head>`;

    // Create section with various paragraph structures
    const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run charPrIDRef="0"><hp:t>Hello</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run charPrIDRef="0"><hp:t>Hello</hp:t></hp:run>
    <hp:run charPrIDRef="1"><hp:t> World</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run charPrIDRef="0"><hp:t>ABC</hp:t></hp:run>
    <hp:run charPrIDRef="1"><hp:t>DEF</hp:t></hp:run>
    <hp:run charPrIDRef="2"><hp:t>GHI</hp:t></hp:run>
  </hp:p>
  <hp:p id="4">
    <hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>
  </hp:p>
  <hp:p id="5">
    <hp:run charPrIDRef="0"><hp:t>Short</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

    zip.file('Contents/header.xml', headerXml);
    zip.file('Contents/section0.xml', sectionXml);
    zip.file('[Content_Types].xml', contentTypesXml);

    return await zip.generateAsync({ type: 'nodebuffer' });
  }

  beforeEach(async () => {
    testFilePath = path.join(__dirname, 'test-preserve-styles.hwpx');
    const buffer = await createMultiRunTestHwpx();
    fs.writeFileSync(testFilePath, buffer);
  });

  describe('Basic Operations', () => {
    it('should update single run paragraph', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 0: single run "Hello"
      const result = doc.updateParagraphTextPreserveStyles(0, 0, 'Updated');
      expect(result).toBe(true);

      const para = doc.getParagraph(0, 0);
      expect(para?.text).toBe('Updated');
      expect(para?.runs.length).toBe(1);
      expect(para?.runs[0].text).toBe('Updated');
    });

    it('should distribute text proportionally across multiple runs', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 1: "Hello" (5 chars) + " World" (6 chars) = 11 chars total
      // New text: "ABCDEFGHIJK" (11 chars)
      // Expected: ~5 chars in first run, ~6 chars in second run
      const result = doc.updateParagraphTextPreserveStyles(0, 1, 'ABCDEFGHIJK');
      expect(result).toBe(true);

      const para = doc.getParagraph(0, 1);
      expect(para?.text).toBe('ABCDEFGHIJK');
      expect(para?.runs.length).toBe(2);

      // Check proportional distribution
      const run1Length = para!.runs[0].text!.length;
      const run2Length = para!.runs[1].text!.length;

      // Original proportion: 5/11 and 6/11
      // With rounding, first run should be ~5 chars (could be 4-5)
      expect(run1Length).toBeGreaterThanOrEqual(4);
      expect(run1Length).toBeLessThanOrEqual(5);
      expect(run2Length).toBeGreaterThanOrEqual(6);
      expect(run2Length).toBeLessThanOrEqual(7);
    });

    it('should preserve run count', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 2: 3 runs "ABC", "DEF", "GHI"
      const originalPara = doc.getParagraph(0, 2);
      const originalRunCount = originalPara!.runs.length;

      doc.updateParagraphTextPreserveStyles(0, 2, 'NewText123');

      const updatedPara = doc.getParagraph(0, 2);
      expect(updatedPara?.runs.length).toBe(originalRunCount);
    });
  });

  describe('Style Preservation', () => {
    it('should preserve charPrIDRef for all runs', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 2: has different charPrIDRef for each run
      const originalPara = doc.getParagraph(0, 2);
      const originalCharPrRefs = originalPara!.runs.map(r => r.charPrIDRef);

      doc.updateParagraphTextPreserveStyles(0, 2, '123456789');

      const updatedPara = doc.getParagraph(0, 2);
      const updatedCharPrRefs = updatedPara!.runs.map(r => r.charPrIDRef);

      expect(updatedCharPrRefs).toEqual(originalCharPrRefs);
    });

    it('should persist style preservation after save', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Get original charPrIDRefs
      const originalPara = doc.getParagraph(0, 1);
      const originalCharPrRefs = originalPara!.runs.map(r => r.charPrIDRef);

      // Update text
      doc.updateParagraphTextPreserveStyles(0, 1, 'NewContent');

      // Save and reload
      const savedBuffer = await doc.save();
      const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

      // Verify styles persisted
      const reloadedPara = doc2.getParagraph(0, 1);
      expect(reloadedPara?.text).toBe('NewContent');

      const reloadedCharPrRefs = reloadedPara!.runs.map(r => r.charPrIDRef);
      expect(reloadedCharPrRefs).toEqual(originalCharPrRefs);
    });

    it('should verify XML structure contains charPrIDRef after save', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      doc.updateParagraphTextPreserveStyles(0, 1, 'StyleTest');
      const savedBuffer = await doc.save();

      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // Should have multiple charPrIDRef attributes in the paragraph
      const charPrRefMatches = savedXml!.match(/charPrIDRef="\d+"/g);
      expect(charPrRefMatches).toBeTruthy();
      expect(charPrRefMatches!.length).toBeGreaterThanOrEqual(2);
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty paragraph', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 3: has empty text
      const result = doc.updateParagraphTextPreserveStyles(0, 3, 'Now has content');
      expect(result).toBe(true);

      const para = doc.getParagraph(0, 3);
      expect(para?.text).toBe('Now has content');
    });

    it('should handle empty new text', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const result = doc.updateParagraphTextPreserveStyles(0, 0, '');
      expect(result).toBe(true);

      const para = doc.getParagraph(0, 0);
      expect(para?.text).toBe('');
    });

    it('should handle very long text', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const longText = 'A'.repeat(10000);
      const result = doc.updateParagraphTextPreserveStyles(0, 1, longText);
      expect(result).toBe(true);

      const para = doc.getParagraph(0, 1);
      expect(para?.text).toBe(longText);
      expect(para?.text.length).toBe(10000);
    });

    it('should return false for non-existent paragraph', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const result = doc.updateParagraphTextPreserveStyles(0, 999, 'Text');
      expect(result).toBe(false);
    });

    it('should return false for invalid section', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const result = doc.updateParagraphTextPreserveStyles(999, 0, 'Text');
      expect(result).toBe(false);
    });
  });

  describe('Proportional Distribution Validation', () => {
    it('should verify exact proportional distribution for 2 runs', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 1: "Hello" (5) + " World" (6) = 11 total
      // Original proportions: 5/11 ≈ 0.4545, 6/11 ≈ 0.5455
      // New text length 22 chars
      // Expected: 10 chars (5/11 * 22 = 10), 12 chars (remainder)
      const newText = 'A'.repeat(22);
      doc.updateParagraphTextPreserveStyles(0, 1, newText);

      const para = doc.getParagraph(0, 1);
      const run1Len = para!.runs[0].text!.length;
      const run2Len = para!.runs[1].text!.length;

      expect(run1Len + run2Len).toBe(22);
      expect(run1Len).toBe(10); // Math.round(0.4545 * 22) = 10
      expect(run2Len).toBe(12); // Remainder
    });

    it('should verify proportional distribution for 3 runs', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 2: "ABC" (3) + "DEF" (3) + "GHI" (3) = 9 total
      // Each run is 1/3 proportion
      // New text: 18 chars → 6, 6, 6
      const newText = '123456789012345678';
      doc.updateParagraphTextPreserveStyles(0, 2, newText);

      const para = doc.getParagraph(0, 2);
      expect(para!.runs[0].text!.length).toBe(6);
      expect(para!.runs[1].text!.length).toBe(6);
      expect(para!.runs[2].text!.length).toBe(6);
      expect(para!.text).toBe(newText);
    });

    it('should handle text shorter than run count', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 2: 3 runs, but only 2 chars of new text
      doc.updateParagraphTextPreserveStyles(0, 2, 'AB');

      const para = doc.getParagraph(0, 2);
      expect(para?.text).toBe('AB');

      // All runs should still exist
      expect(para?.runs.length).toBe(3);

      // Some runs may be empty, but total should equal 'AB'
      const totalText = para!.runs.map(r => r.text || '').join('');
      expect(totalText).toBe('AB');
    });

    it('should handle text much longer than original', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 4: "Short" (5 chars), update to 100 chars
      const longText = 'X'.repeat(100);
      doc.updateParagraphTextPreserveStyles(0, 4, longText);

      const para = doc.getParagraph(0, 4);
      expect(para?.text).toBe(longText);
      expect(para?.text.length).toBe(100);
    });
  });

  describe('Persistence After Save', () => {
    it('should persist text changes after save and reload', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      doc.updateParagraphTextPreserveStyles(0, 1, 'Persisted Text');

      const savedBuffer = await doc.save();
      const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

      const para = doc2.getParagraph(0, 1);
      expect(para?.text).toBe('Persisted Text');
    });

    it('should verify XML contains updated text', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      doc.updateParagraphTextPreserveStyles(0, 0, 'XMLVerify');

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

      expect(savedXml).toContain('XMLVerify');
    });

    it('should maintain all runs in XML after save', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      // Element 2 has 3 runs
      doc.updateParagraphTextPreserveStyles(0, 2, 'ThreeRuns');

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // Extract paragraph with id="3" (element 2)
      const paraMatch = savedXml!.match(/<hp:p id="3"[^>]*>([\s\S]*?)<\/hp:p>/);
      expect(paraMatch).toBeTruthy();

      const paraContent = paraMatch![1];
      const runCount = (paraContent.match(/<hp:run/g) || []).length;
      expect(runCount).toBe(3);
    });
  });

  describe('Unicode and Special Characters', () => {
    it('should handle Korean text', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const koreanText = '안녕하세요 세계';
      doc.updateParagraphTextPreserveStyles(0, 1, koreanText);

      const para = doc.getParagraph(0, 1);
      expect(para?.text).toBe(koreanText);

      // Verify after save
      const savedBuffer = await doc.save();
      const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);
      const para2 = doc2.getParagraph(0, 1);
      expect(para2?.text).toBe(koreanText);
    });

    it('should handle emoji and special Unicode', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const emojiText = 'Hello 🌍 世界 🎉';
      doc.updateParagraphTextPreserveStyles(0, 0, emojiText);

      const para = doc.getParagraph(0, 0);
      expect(para?.text).toBe(emojiText);
    });

    it('should handle XML special characters', async () => {
      const buffer = fs.readFileSync(testFilePath);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);

      const xmlChars = '<tag> & "quotes" \'apostrophe\'';
      doc.updateParagraphTextPreserveStyles(0, 0, xmlChars);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // Should be properly escaped in XML
      expect(savedXml).toContain('&lt;tag&gt;');
      expect(savedXml).toContain('&amp;');
      expect(savedXml).toContain('&quot;');

      // Should decode correctly on reload
      const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);
      const para = doc2.getParagraph(0, 0);
      expect(para?.text).toBe(xmlChars);
    });
  });
});

describe('HwpxDocument - Paragraph Insert', () => {
  let testFilePath: string;

  beforeEach(async () => {
    // Create test HWPX file
    testFilePath = path.join(__dirname, 'test-para-insert.hwpx');
    const buffer = await createTestHwpxBuffer();
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should insert paragraph and persist after save', async () => {
    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, originalBuffer);

    // 2. Insert paragraph after first element (index 0)
    const newIndex = doc.insertParagraph(0, 0, '새로운 문단 텍스트');
    expect(newIndex).toBe(1);

    // 3. Verify in memory
    const paragraphs = doc.getParagraphs(0);
    expect(paragraphs?.length).toBeGreaterThan(1);

    // 4. Save
    const savedBuffer = await doc.save();

    // 5. Reload and verify
    const reloadedDoc = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);
    const reloadedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await reloadedZip.file('Contents/section0.xml')?.async('string');

    console.log('Saved XML with inserted paragraph:', savedXml);

    // Verify the inserted text is in XML
    expect(savedXml).toContain('새로운 문단 텍스트');

    // Verify document text includes the new paragraph
    const docText = reloadedDoc.getAllText();
    expect(docText).toContain('새로운 문단 텍스트');
  });

  it('should insert paragraph with special characters and escape properly', async () => {
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, originalBuffer);

    // Insert paragraph with special XML characters
    doc.insertParagraph(0, 0, '특수문자 테스트: <tag> & "quotes" \'apostrophe\'');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // XML should be properly escaped
    expect(savedXml).toContain('&lt;tag&gt;');
    expect(savedXml).toContain('&amp;');
    expect(savedXml).toContain('&quot;quotes&quot;');

    // Reload and verify text is decoded correctly
    const reloadedDoc = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);
    const docText = reloadedDoc.getAllText();
    expect(docText).toContain('<tag>');
    expect(docText).toContain('&');
  });

  it('should insert multiple paragraphs and preserve order', async () => {
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, originalBuffer);

    // Insert multiple paragraphs
    doc.insertParagraph(0, 0, '첫 번째 삽입');
    doc.insertParagraph(0, 1, '두 번째 삽입');
    doc.insertParagraph(0, 2, '세 번째 삽입');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // All paragraphs should be in the XML
    expect(savedXml).toContain('첫 번째 삽입');
    expect(savedXml).toContain('두 번째 삽입');
    expect(savedXml).toContain('세 번째 삽입');

    // Verify order (first should come before second)
    const firstIdx = savedXml!.indexOf('첫 번째 삽입');
    const secondIdx = savedXml!.indexOf('두 번째 삽입');
    const thirdIdx = savedXml!.indexOf('세 번째 삽입');
    expect(firstIdx).toBeLessThan(secondIdx);
    expect(secondIdx).toBeLessThan(thirdIdx);
  });

  it('should create new document, insert paragraph, and save valid HWPX', async () => {
    // 1. Create new document using createNew()
    const doc = HwpxDocument.createNew('test-new-doc', 'Test Document', 'Test Author');

    // 2. Insert a paragraph
    const newIndex = doc.insertParagraph(0, 0, 'Hello World - 새 문서 테스트');
    expect(newIndex).toBe(1);

    // 3. Save
    const savedBuffer = await doc.save();
    expect(savedBuffer).toBeDefined();
    expect(savedBuffer.byteLength).toBeGreaterThan(0);

    // 4. Verify ZIP structure
    const savedZip = await JSZip.loadAsync(savedBuffer);

    // Check required files exist
    expect(savedZip.file('mimetype')).not.toBeNull();
    expect(savedZip.file('version.xml')).not.toBeNull();
    expect(savedZip.file('Contents/content.hpf')).not.toBeNull();
    expect(savedZip.file('Contents/header.xml')).not.toBeNull();
    expect(savedZip.file('Contents/section0.xml')).not.toBeNull();

    // 5. Verify XML validity
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();
    console.log('New document section XML:', sectionXml);

    // Check XML is well-formed (no orphan tags)
    const openSec = (sectionXml!.match(/<hs:sec/g) || []).length;
    const closeSec = (sectionXml!.match(/<\/hs:sec>/g) || []).length;
    expect(openSec).toBe(closeSec);

    const openP = (sectionXml!.match(/<hp:p[ >]/g) || []).length;
    const closeP = (sectionXml!.match(/<\/hp:p>/g) || []).length;
    expect(openP).toBe(closeP);

    // Verify inserted text is in XML
    expect(sectionXml).toContain('Hello World - 새 문서 테스트');

    // 6. Reload and verify content persists
    const reloadedDoc = await HwpxDocument.createFromBuffer('test-reload', 'test.hwpx', savedBuffer);
    const docText = reloadedDoc.getAllText();
    expect(docText).toContain('Hello World - 새 문서 테스트');
  });

  it('should validate header.xml in newly created document', async () => {
    const doc = HwpxDocument.createNew('test-header', 'Header Test', 'Author');
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
    expect(headerXml).toBeDefined();
    console.log('Header XML:', headerXml);

    // Check for balanced tags (use word boundary to avoid matching <hh:heading as <hh:head)
    const openHead = (headerXml!.match(/<hh:head[ >]/g) || []).length;
    const closeHead = (headerXml!.match(/<\/hh:head>/g) || []).length;
    expect(openHead).toBe(closeHead);

    // Check that there are no broken closing tags (like <\tag> instead of </tag>)
    expect(headerXml).not.toMatch(/<\\[a-zA-Z]/);
  });

  it('should only update target paragraph when multiple paragraphs have same text', async () => {
    // The test file now has 3 paragraphs with "Hello World" (element 0, 2, 3)
    // When updating element 0, only that paragraph should change

    // 1. Open document
    const buffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-dup', testFilePath, buffer);

    // 2. Get paragraphs and verify initial state
    const paragraphs = doc.getParagraphs(0);
    console.log('Initial paragraphs:', paragraphs);

    // Element indices: 0=para, 1=table, 2=para, 3=para
    // Paragraph indices in getParagraphs: 0=element0, 1=element2, 2=element3
    expect(paragraphs.length).toBe(3); // 3 top-level paragraphs
    expect(paragraphs[0].text).toBe('Hello World');
    expect(paragraphs[1].text).toBe('Hello World');
    expect(paragraphs[2].text).toBe('Hello World');

    // 3. Update only the first paragraph (elementIndex 0, runIndex 0)
    doc.updateParagraphText(0, 0, 0, 'Updated First');

    // 4. Save and check XML
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    console.log('Saved XML:', savedXml);

    // 5. Verify: only the first paragraph was updated
    expect(savedXml).toContain('Updated First');

    // Count remaining "Hello World" - should be 2 (element 2 and 3)
    const helloCount = (savedXml!.match(/Hello World/g) || []).length;
    console.log('Hello World count:', helloCount);
    expect(helloCount).toBe(2); // NOT 0 or 3

    // Also verify table content wasn't affected
    expect(savedXml).toContain('Cell 0,0');
    expect(savedXml).toContain('Cell 1,0');
    expect(savedXml).toContain('Cell 1,1');
  });
});
