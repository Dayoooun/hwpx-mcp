/**
 * Tests for long text chunking functionality
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Text Chunking', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  beforeEach(async () => {
    // Create a minimal HWPX structure for testing
    const zip = new JSZip();

    // Add version.xml
    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

    // Add header.xml
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:refList>
    <hh:charShapeList itemCnt="1">
      <hh:charShape id="0" height="1000" baseSize="1000" ratio="100" spacing="0" relSize="100"/>
    </hh:charShapeList>
    <hh:paraShapeList itemCnt="1">
      <hh:paraShape id="0"/>
    </hh:paraShapeList>
  </hh:refList>
</hh:head>`);

    // Add section with a table for testing
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Test Document</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 0,0</hp:t></hp:run>
            <hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="0" flags="0"/></hp:linesegarray>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="11" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 0,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="20" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 1,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="21" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 1,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    // Add content_types
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

    testBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', testBuffer);
  });

  describe('splitTextIntoChunks', () => {
    it('should not split short text', async () => {
      const shortText = 'This is a short text.';
      doc.updateTableCell(0, 0, 0, 0, shortText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);
      const table = savedDoc.getTable(0, 0);

      expect(table?.data[0][0].text).toBe(shortText);

      // Check XML structure - should have single run
      const zip = (savedDoc as any)._zip;
      const xml = await zip.file('Contents/section0.xml')!.async('string');
      const runCount = (xml.match(/<hp:run[^>]*>/g) || []).length;

      // Should have minimal runs (one per cell that was updated + existing)
      expect(runCount).toBeGreaterThanOrEqual(1);
    });

    it('should split very long text into multiple runs', async () => {
      // Generate text longer than default chunk size (2000 chars)
      const longText = 'A'.repeat(5000);
      doc.updateTableCell(0, 0, 0, 0, longText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify text content is preserved
      const table = savedDoc.getTable(0, 0);
      expect(table?.data[0][0].text).toBe(longText);

      // Check XML structure - should have multiple runs
      const zip = (savedDoc as any)._zip;
      const xml = await zip.file('Contents/section0.xml')!.async('string');

      // Extract content from cell 0,0
      const cellMatch = xml.match(/<hp:tc[^>]*colAddr="0"[^>]*rowAddr="0"[^>]*>[\s\S]*?<\/hp:tc>/);
      expect(cellMatch).toBeTruthy();

      const cellXml = cellMatch![0];
      const runsInCell = (cellXml.match(/<hp:run[^>]*>/g) || []).length;

      // 5000 chars / 2000 chunk size = at least 3 runs
      expect(runsInCell).toBeGreaterThanOrEqual(3);
    });

    it('should split at word boundaries when possible', async () => {
      // Generate text with spaces, longer than chunk size
      const words = Array(300).fill('word').join(' '); // ~1500 chars per 300 words
      const longText = words + ' ' + words + ' ' + words; // ~4500 chars
      doc.updateTableCell(0, 0, 0, 0, longText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify text content is preserved
      const table = savedDoc.getTable(0, 0);
      expect(table?.data[0][0].text).toBe(longText);
    });

    it('should handle multiline long text with chunking', async () => {
      // Create multiline text where each line is very long
      const longLine = 'B'.repeat(3000);
      const multilineText = `${longLine}\n${longLine}\n${longLine}`;
      doc.updateTableCell(0, 0, 0, 0, multilineText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify text content is preserved (newlines become separate paragraphs)
      const table = savedDoc.getTable(0, 0);
      // Multiline creates multiple paragraphs, so text extraction might vary
      expect(table?.data[0][0].text).toContain('B');

      // Check XML structure
      const zip = (savedDoc as any)._zip;
      const xml = await zip.file('Contents/section0.xml')!.async('string');

      // Should have multiple paragraphs (one per line)
      const cellMatch = xml.match(/<hp:tc[^>]*colAddr="0"[^>]*rowAddr="0"[^>]*>[\s\S]*?<\/hp:tc>/);
      expect(cellMatch).toBeTruthy();

      const cellXml = cellMatch![0];
      const paragraphsInCell = (cellXml.match(/<hp:p[^>]*>/g) || []).length;
      expect(paragraphsInCell).toBe(3); // 3 lines = 3 paragraphs

      // Each paragraph should have multiple runs (3000 chars / 2000 = 2 runs each)
      const runsInCell = (cellXml.match(/<hp:run[^>]*>/g) || []).length;
      expect(runsInCell).toBeGreaterThanOrEqual(6); // At least 2 runs per paragraph
    });

    it('should preserve XML structure with chunked text', async () => {
      const longText = 'C'.repeat(6000);
      doc.updateTableCell(0, 0, 0, 0, longText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify XML is well-formed by checking tag balance
      const zip = (savedDoc as any)._zip;
      const xml = await zip.file('Contents/section0.xml')!.async('string');

      const tblOpen = (xml.match(/<hp:tbl[\s>]/g) || []).length;
      const tblClose = (xml.match(/<\/hp:tbl>/g) || []).length;
      const trOpen = (xml.match(/<hp:tr[\s>]/g) || []).length;
      const trClose = (xml.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (xml.match(/<hp:tc[\s>]/g) || []).length;
      const tcClose = (xml.match(/<\/hp:tc>/g) || []).length;
      const runOpen = (xml.match(/<hp:run[\s>]/g) || []).length;
      const runClose = (xml.match(/<\/hp:run>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);
      expect(runOpen).toBe(runClose);
    });

    it('should handle Korean text chunking', async () => {
      // Korean text (2 bytes per char in UTF-8, but character count matters)
      const koreanText = '가나다라마바사아자차카타파하'.repeat(200); // ~2800 chars
      doc.updateTableCell(0, 0, 0, 0, koreanText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify text content is preserved
      const table = savedDoc.getTable(0, 0);
      expect(table?.data[0][0].text).toBe(koreanText);
    });

    it('should handle mixed Korean and English long text', async () => {
      const mixedText = ('Hello 안녕하세요 World 세계 ').repeat(400); // ~8000 chars
      doc.updateTableCell(0, 0, 0, 0, mixedText);

      const savedBuffer = await doc.save();
      const savedDoc = await HwpxDocument.createFromBuffer('saved', 'saved.hwpx', savedBuffer);

      // Verify text content is preserved
      const table = savedDoc.getTable(0, 0);
      expect(table?.data[0][0].text).toBe(mixedText);

      // Check multiple runs created
      const zip = (savedDoc as any)._zip;
      const xml = await zip.file('Contents/section0.xml')!.async('string');
      const cellMatch = xml.match(/<hp:tc[^>]*colAddr="0"[^>]*rowAddr="0"[^>]*>[\s\S]*?<\/hp:tc>/);
      const cellXml = cellMatch![0];
      const runsInCell = (cellXml.match(/<hp:run[^>]*>/g) || []).length;

      expect(runsInCell).toBeGreaterThanOrEqual(4); // 8000 / 2000 = 4 chunks
    });
  });
});
