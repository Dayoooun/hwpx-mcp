/**
 * Tests for hanging indent applied to multiple paragraphs in a cell
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Multiple Paragraph Hanging Indent', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  beforeEach(async () => {
    const zip = new JSZip();

    // version.xml
    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

    // header.xml with initial paraPr definitions
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/common">
  <hh:refList>
    <hh:charShapeList itemCnt="1">
      <hh:charShape id="0" height="1000" baseSize="1000" ratio="100" spacing="0" relSize="100"/>
    </hh:charShapeList>
    <hh:paraProperties itemCnt="1">
      <hh:paraPr id="0" tabPrIDRef="0">
        <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
        <hh:margin>
          <hc:intent value="0" unit="HWPUNIT"/>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="0" unit="HWPUNIT"/>
        </hh:margin>
      </hh:paraPr>
    </hh:paraProperties>
    <hh:paraShapeList itemCnt="1">
      <hh:paraShape id="0"/>
    </hh:paraShapeList>
  </hh:refList>
</hh:head>`);

    // section0.xml with a simple table
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Original content</hp:t></hp:run>
            <hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="0" flags="0"/></hp:linesegarray>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    // content.hpf
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="urn:hancom:hwp:file">
  <hpf:manifest>
    <hpf:item href="Contents/section0.xml" type="application/xml"/>
    <hpf:item href="Contents/header.xml" type="application/xml"/>
  </hpf:manifest>
</hpf:package>`);

    // Content_Types
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

    testBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', testBuffer);
  });

  it('should apply hanging indent to ALL paragraphs when updating cell with multi-line text', async () => {
    // Update cell with 5 lines, each with a marker
    const multiLineText = [
      '○ 첫 번째 항목입니다.',
      '○ 두 번째 항목입니다.',
      '○ 세 번째 항목입니다.',
      '○ 네 번째 항목입니다.',
      '○ 다섯 번째 항목입니다.',
    ].join('\n');

    // First update the cell content (this creates 5 paragraphs in XML)
    doc.updateTableCell(0, 0, 0, 0, multiLineText);

    // Then apply hanging indent to each paragraph
    for (let i = 0; i < 5; i++) {
      doc.setTableCellHangingIndent(0, 0, 0, 0, i, 12); // 12pt indent
    }

    // Save the document
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    // Check header.xml for new paraPr
    const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
    expect(headerXml).toBeDefined();

    // Should have at least 2 paraPr (original + new with hanging indent)
    const paraPrCount = (headerXml?.match(/<hh:paraPr\b/g) || []).length;
    expect(paraPrCount).toBeGreaterThanOrEqual(2);

    // Check section0.xml
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // Find all paragraphs in the cell
    const cellMatch = sectionXml?.match(/<hp:tc[^>]*>([\s\S]*?)<\/hp:tc>/);
    expect(cellMatch).toBeTruthy();

    const cellContent = cellMatch![1];

    // Count paragraphs
    const paragraphs = cellContent.match(/<hp:p\b[^>]*>/g) || [];
    expect(paragraphs.length).toBe(5);

    // Extract paraPrIDRef from each paragraph
    const paraPrIds: string[] = [];
    for (const p of paragraphs) {
      const match = p.match(/paraPrIDRef="(\d+)"/);
      if (match) {
        paraPrIds.push(match[1]);
      }
    }

    // All paragraphs should have the same non-zero paraPrIDRef

    // ALL paragraphs should have the NEW paraPrIDRef (not 0)
    // This is the key assertion - if the bug exists, only the first will be updated
    for (let i = 0; i < paraPrIds.length; i++) {
      expect(paraPrIds[i]).not.toBe('0');
    }

    // All paragraphs should have the same paraPrIDRef (same indent)
    const uniqueIds = new Set(paraPrIds);
    expect(uniqueIds.size).toBe(1);
  });

  it('should apply different hanging indents to paragraphs with different markers', async () => {
    // Different markers have different indent widths
    const multiLineText = [
      '○ 항목1',  // Circle marker
      '1. 항목2', // Number marker
      '가. 항목3', // Korean marker
    ].join('\n');

    doc.updateTableCell(0, 0, 0, 0, multiLineText);

    // Apply different indents
    doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10); // 10pt for ○
    doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 12); // 12pt for 1.
    doc.setTableCellHangingIndent(0, 0, 0, 0, 2, 14); // 14pt for 가.

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    // Check header.xml - should have 3 new paraPr (one for each indent size)
    const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
    const paraPrCount = (headerXml?.match(/<hh:paraPr\b/g) || []).length;
    expect(paraPrCount).toBeGreaterThanOrEqual(4); // 1 original + 3 new

    // Check section0.xml
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    const cellMatch = sectionXml?.match(/<hp:tc[^>]*>([\s\S]*?)<\/hp:tc>/);
    const cellContent = cellMatch![1];

    // Extract paraPrIDRef from each paragraph
    const paragraphs = cellContent.match(/<hp:p\b[^>]*>/g) || [];
    const paraPrIds: string[] = [];
    for (const p of paragraphs) {
      const match = p.match(/paraPrIDRef="(\d+)"/);
      if (match) {
        paraPrIds.push(match[1]);
      }
    }

    // Each paragraph should have a different paraPrIDRef (different indent sizes)

    // All should be non-zero
    for (const id of paraPrIds) {
      expect(id).not.toBe('0');
    }

    // All should be different (different indent sizes)
    expect(new Set(paraPrIds).size).toBe(3);
  });
});
