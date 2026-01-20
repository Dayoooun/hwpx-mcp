/**
 * Tests for parallel table updates (reproducing real-world bug)
 * Bug scenario: update_table_cell called for Table 17 + Table 19 in parallel
 * then save() → "Broken tag structure" error
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Parallel Table Update (Real-world Bug Fix)', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  beforeEach(async () => {
    const zip = new JSZip();

    // version.xml
    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

    // header.xml with paraPr definitions
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

    // Create section with 20 tables to simulate real document (indices 0-19)
    let tablesXml = '';
    for (let i = 0; i < 20; i++) {
      tablesXml += `
  <hp:p id="${1000 + i}" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>Before Table ${i}</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="${100 + i}" rowCnt="3" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10000 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 0,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10001 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 0,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10010 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 1,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10011 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 1,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="2" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10020 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 2,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="2" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${10021 + i * 100}" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Table ${i} Cell 2,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>`;
    }

    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
${tablesXml}
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

  it('should handle parallel updates to Table 17 and Table 19', async () => {
    // Simulate parallel updates (like real-world scenario)
    // Update Table 17 Cell (0, 1)
    doc.updateTableCell(0, 17, 0, 1, '○ 테이블17 항목1\n○ 테이블17 항목2');

    // Update Table 19 Cell (0, 1)
    doc.updateTableCell(0, 19, 0, 1, '○ 테이블19 항목1\n○ 테이블19 항목2\n○ 테이블19 항목3');

    // Apply hanging indent to both tables (auto_hanging_indent simulation)
    doc.setTableCellHangingIndent(0, 17, 0, 1, 0, 12);
    doc.setTableCellHangingIndent(0, 17, 0, 1, 1, 12);
    doc.setTableCellHangingIndent(0, 19, 0, 1, 0, 12);
    doc.setTableCellHangingIndent(0, 19, 0, 1, 1, 12);
    doc.setTableCellHangingIndent(0, 19, 0, 1, 2, 12);

    // Save should NOT fail with "Broken tag structure" error
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    // Verify XML structure is valid
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // Check for broken tag structure pattern (this was the bug)
    expect(sectionXml).not.toMatch(/<[^>]*</);

    // Verify tag balance
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;
    const trOpen = (sectionXml?.match(/<hp:tr\b/g) || []).length;
    const trClose = (sectionXml?.match(/<\/hp:tr>/g) || []).length;
    const tcOpen = (sectionXml?.match(/<hp:tc\b/g) || []).length;
    const tcClose = (sectionXml?.match(/<\/hp:tc>/g) || []).length;

    expect(tblOpen).toBe(tblClose);
    expect(trOpen).toBe(trClose);
    expect(tcOpen).toBe(tcClose);

    // Should still have 20 tables
    expect(tblOpen).toBe(20);

    // Verify the updated content is present
    expect(sectionXml).toContain('테이블17');
    expect(sectionXml).toContain('테이블19');
  });

  it('should handle sequential updates followed by single save', async () => {
    // Simulate sequential updates like real-world test
    // Table 1
    doc.updateTableCell(0, 1, 0, 0, '○ 첫 번째 테이블');
    doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 10);

    // Table 4
    doc.updateTableCell(0, 4, 1, 0, '○ 네 번째 테이블');
    doc.setTableCellHangingIndent(0, 4, 1, 0, 0, 10);

    // Table 7
    doc.updateTableCell(0, 7, 2, 1, '○ 일곱 번째 테이블');
    doc.setTableCellHangingIndent(0, 7, 2, 1, 0, 10);

    // Table 10
    doc.updateTableCell(0, 10, 0, 1, '○ 열 번째 테이블');
    doc.setTableCellHangingIndent(0, 10, 0, 1, 0, 10);

    // Table 12
    doc.updateTableCell(0, 12, 1, 1, '○ 열두 번째 테이블');
    doc.setTableCellHangingIndent(0, 12, 1, 1, 0, 10);

    // Table 14
    doc.updateTableCell(0, 14, 2, 0, '○ 열네 번째 테이블');
    doc.setTableCellHangingIndent(0, 14, 2, 0, 0, 10);

    // Save should work
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();
    expect(sectionXml).not.toMatch(/<[^>]*</);

    // All updates should be present
    expect(sectionXml).toContain('첫 번째 테이블');
    expect(sectionXml).toContain('네 번째 테이블');
    expect(sectionXml).toContain('일곱 번째 테이블');
    expect(sectionXml).toContain('열 번째 테이블');
    expect(sectionXml).toContain('열두 번째 테이블');
    expect(sectionXml).toContain('열네 번째 테이블');
  });

  it('should handle updates to non-consecutive tables', async () => {
    // Update tables 0, 5, 10, 15, 19 (spread across the document)
    const tableIndices = [0, 5, 10, 15, 19];

    for (const idx of tableIndices) {
      doc.updateTableCell(0, idx, 0, 0, `○ 테이블${idx} 수정됨`);
      doc.setTableCellHangingIndent(0, idx, 0, 0, 0, 12);
    }

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();

    // Verify all updates
    for (const idx of tableIndices) {
      expect(sectionXml).toContain(`테이블${idx} 수정됨`);
    }

    // Verify no corruption
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;
    expect(tblOpen).toBe(20);
    expect(tblOpen).toBe(tblClose);
  });

  it('should handle multi-line text updates to multiple tables', async () => {
    // Update multiple tables with multi-line text
    doc.updateTableCell(0, 3, 0, 0, '1. 항목A\n2. 항목B\n3. 항목C');
    doc.updateTableCell(0, 8, 1, 1, '가. 내용1\n나. 내용2');
    doc.updateTableCell(0, 13, 2, 0, '(1) 첫째\n(2) 둘째\n(3) 셋째\n(4) 넷째');

    // Apply hanging indent to all paragraphs
    for (let i = 0; i < 3; i++) {
      doc.setTableCellHangingIndent(0, 3, 0, 0, i, 15);
    }
    for (let i = 0; i < 2; i++) {
      doc.setTableCellHangingIndent(0, 8, 1, 1, i, 18);
    }
    for (let i = 0; i < 4; i++) {
      doc.setTableCellHangingIndent(0, 13, 2, 0, i, 20);
    }

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);

    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toBeDefined();
    expect(sectionXml).not.toMatch(/<[^>]*</);

    // Verify content
    expect(sectionXml).toContain('항목A');
    expect(sectionXml).toContain('내용1');
    expect(sectionXml).toContain('첫째');

    // Tag balance
    const tblOpen = (sectionXml?.match(/<hp:tbl\b/g) || []).length;
    const tblClose = (sectionXml?.match(/<\/hp:tbl>/g) || []).length;
    expect(tblOpen).toBe(tblClose);
  });

  it('should handle update/save cycle multiple times', async () => {
    // First update cycle
    doc.updateTableCell(0, 5, 0, 0, '○ 첫번째 수정');
    doc.setTableCellHangingIndent(0, 5, 0, 0, 0, 10);

    let savedBuffer = await doc.save();
    let savedZip = await JSZip.loadAsync(savedBuffer);
    let sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toContain('첫번째 수정');
    expect(sectionXml).not.toMatch(/<[^>]*</);

    // Reload document
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', savedBuffer);

    // Second update cycle - different table
    doc.updateTableCell(0, 10, 1, 1, '○ 두번째 수정');
    doc.setTableCellHangingIndent(0, 10, 1, 1, 0, 12);

    savedBuffer = await doc.save();
    savedZip = await JSZip.loadAsync(savedBuffer);
    sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).toContain('첫번째 수정');
    expect(sectionXml).toContain('두번째 수정');
    expect(sectionXml).not.toMatch(/<[^>]*</);

    // Reload again
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', savedBuffer);

    // Third cycle - update previously modified table
    doc.updateTableCell(0, 5, 0, 0, '○ 세번째 수정 (덮어쓰기)');
    doc.setTableCellHangingIndent(0, 5, 0, 0, 0, 14);

    savedBuffer = await doc.save();
    savedZip = await JSZip.loadAsync(savedBuffer);
    sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
    expect(sectionXml).not.toContain('첫번째 수정'); // Overwritten
    expect(sectionXml).toContain('세번째 수정');
    expect(sectionXml).toContain('두번째 수정');
    expect(sectionXml).not.toMatch(/<[^>]*</);
  });
});
