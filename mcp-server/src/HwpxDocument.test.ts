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

  // Section with a table containing cells
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
