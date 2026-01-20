/**
 * Tests for hanging indent applied to multiple tables (stale position bug fix)
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Multiple Table Hanging Indent (Stale Position Fix)', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  beforeEach(async () => {
    const zip = new JSZip();

    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

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

    // Create section with 3 tables to test multiple table handling
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>Before Table 0</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table 0 content</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p id="2" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>Between Table 0 and 1</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="200" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="20" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table 1 content</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p id="3" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>Between Table 1 and 2</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="300" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="30" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table 2 content</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="urn:hancom:hwp:file">
  <hpf:manifest>
    <hpf:item href="Contents/section0.xml" type="application/xml"/>
    <hpf:item href="Contents/header.xml" type="application/xml"/>
  </hpf:manifest>
</hpf:package>`);

    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`);

    testBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', testBuffer);
  });

  it('should apply hanging indent to multiple tables without XML corruption', async () => {
    // Update all 3 tables with multi-line text
    const texts = [
      '○ 테이블0 항목1\n○ 테이블0 항목2',
      '○ 테이블1 항목1\n○ 테이블1 항목2\n○ 테이블1 항목3',
      '○ 테이블2 항목1\n○ 테이블2 항목2',
    ];

    // Update cells
    doc.updateTableCell(0, 0, 0, 0, texts[0]);
    doc.updateTableCell(0, 1, 0, 0, texts[1]);
    doc.updateTableCell(0, 2, 0, 0, texts[2]);

    // Apply hanging indent to all paragraphs in all tables
    // Table 0: 2 paragraphs
    doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 12);
    doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 12);

    // Table 1: 3 paragraphs
    doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
    doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 12);
    doc.setTableCellHangingIndent(0, 1, 0, 0, 2, 12);

    // Table 2: 2 paragraphs
    doc.setTableCellHangingIndent(0, 2, 0, 0, 0, 12);
    doc.setTableCellHangingIndent(0, 2, 0, 0, 1, 12);

    // Save the document - THIS SHOULD NOT FAIL
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    // Verify XML structure is valid
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // Count opening and closing tags to ensure balance
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;
    const trOpen = (sectionXml?.match(/<hp:tr\b/g) || []).length;
    const trClose = (sectionXml?.match(/<\/hp:tr>/g) || []).length;
    const tcOpen = (sectionXml?.match(/<hp:tc\b/g) || []).length;
    const tcClose = (sectionXml?.match(/<\/hp:tc>/g) || []).length;

    // Tags should be balanced
    expect(tblOpen).toBe(tblClose);
    expect(trOpen).toBe(trClose);
    expect(tcOpen).toBe(tcClose);

    // Should still have 3 tables
    expect(tblOpen).toBe(3);
  });

  it('should handle non-consecutive table indices correctly', async () => {
    // Only update tables 0 and 2 (skip table 1)
    doc.updateTableCell(0, 0, 0, 0, '○ 항목A');
    doc.updateTableCell(0, 2, 0, 0, '○ 항목B');

    doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10);
    doc.setTableCellHangingIndent(0, 2, 0, 0, 0, 14);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // Verify XML is not corrupted
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;

    expect(tblOpen).toBe(3);
    expect(tblOpen).toBe(tblClose);

    // Verify header has new paraPr entries with different indent values
    const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
    expect(headerXml).toBeDefined();

    // Should have at least 3 paraPr (original + 2 new)
    const paraPrCount = (headerXml?.match(/<hh:paraPr\b/g) || []).length;
    expect(paraPrCount).toBeGreaterThanOrEqual(3);
  });

  it('should process tables in descending order to avoid stale positions', async () => {
    // This test verifies that modifying table 2 first doesn't corrupt tables 0 and 1
    doc.updateTableCell(0, 0, 0, 0, '○ 작은텍스트');
    doc.updateTableCell(0, 1, 0, 0, '○ 중간텍스트입니다아아아아');
    doc.updateTableCell(0, 2, 0, 0, '○ 매우긴텍스트입니다아아아아아아아아아아아아아아');

    // Apply hanging indent (different sizes to create different paraPr)
    doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10);
    doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
    doc.setTableCellHangingIndent(0, 2, 0, 0, 0, 14);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // All tables should be present and valid
    expect(sectionXml).toContain('작은텍스트');
    expect(sectionXml).toContain('중간텍스트');
    expect(sectionXml).toContain('매우긴텍스트');

    // Tag balance check
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;

    expect(tblOpen).toBe(3);
    expect(tblOpen).toBe(tblClose);
  });
});
