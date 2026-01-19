import { describe, it, expect, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * Create a 3x3 table for cell merge testing.
 */
async function createTestHwpxWithTable(rows: number = 3, cols: number = 3): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

  // Generate table rows dynamically
  let tableRows = '';
  for (let r = 0; r < rows; r++) {
    tableRows += `    <hp:tr>\n`;
    for (let c = 0; c < cols; c++) {
      tableRows += `      <hp:tc colAddr="${c}" rowAddr="${r}" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${r * 10 + c}">
            <hp:run><hp:t>Cell ${r},${c}</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>\n`;
    }
    tableRows += `    </hp:tr>\n`;
  }

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Document with table</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="${rows}" colCnt="${cols}">
${tableRows}  </hp:tbl>
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

describe('HwpxDocument - Cell Merge (가로 병합)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-merge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should merge two cells horizontally (colSpan)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0) and (0,1) horizontally
    const result = doc.mergeCells(0, 0, 0, 0, 0, 1);
    expect(result).toBe(true);

    // Save and reload
    const savedBuffer = await doc.save();
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // Check XML structure
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    console.log('Merged XML:', savedXml);

    // Verify: First cell should have colSpan="2"
    expect(savedXml).toContain('colSpan="2"');

    // Verify: Row 0 should have only 2 cells (one merged covering 2 columns, one remaining)
    const row0Matches = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    expect(row0Matches).toBeDefined();
    if (row0Matches) {
      const row0 = row0Matches[0];
      const cellsInRow0 = (row0.match(/<hp:tc\b/g) || []).length;
      expect(cellsInRow0).toBe(2); // 2 cells in first row after merge
    }
  });

  it('should merge three cells horizontally', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge all cells in row 1: (1,0), (1,1), (1,2)
    const result = doc.mergeCells(0, 0, 1, 0, 1, 2);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Verify: Merged cell should have colSpan="3"
    expect(savedXml).toContain('colSpan="3"');

    // Row 1 should have only 1 cell
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows && rows[1]) {
      const cellsInRow1 = (rows[1].match(/<hp:tc\b/g) || []).length;
      expect(cellsInRow1).toBe(1);
    }
  });
});

describe('HwpxDocument - Cell Merge (세로 병합)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-merge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should merge two cells vertically (rowSpan)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0) and (1,0) vertically
    const result = doc.mergeCells(0, 0, 0, 0, 1, 0);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    console.log('Vertically merged XML:', savedXml);

    // Verify: First cell should have rowSpan="2"
    expect(savedXml).toContain('rowSpan="2"');

    // Row 0 should have 3 cells, Row 1 should have 2 cells (missing col 0)
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      const cellsInRow0 = (rows[0].match(/<hp:tc\b/g) || []).length;
      const cellsInRow1 = (rows[1].match(/<hp:tc\b/g) || []).length;
      expect(cellsInRow0).toBe(3);
      expect(cellsInRow1).toBe(2); // One cell removed due to rowSpan
    }
  });

  it('should merge three cells vertically', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,2), (1,2), (2,2) vertically
    const result = doc.mergeCells(0, 0, 0, 2, 2, 2);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Verify: Merged cell should have rowSpan="3"
    expect(savedXml).toContain('rowSpan="3"');

    // Row 0: 3 cells, Row 1: 2 cells, Row 2: 2 cells
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(3);
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(2);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(2);
    }
  });
});

describe('HwpxDocument - Cell Merge (블록 병합)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-merge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should merge 2x2 block of cells', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge 2x2 block: (0,0), (0,1), (1,0), (1,1)
    const result = doc.mergeCells(0, 0, 0, 0, 1, 1);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    console.log('Block merged XML:', savedXml);

    // Verify: First cell should have colSpan="2" and rowSpan="2"
    expect(savedXml).toContain('colSpan="2"');
    expect(savedXml).toContain('rowSpan="2"');

    // Row 0: 2 cells (merged + col 2), Row 1: 1 cell (col 2), Row 2: 3 cells
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(2);
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(1);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(3);
    }
  });

  it('should merge entire 3x3 table into one cell', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge entire table
    const result = doc.mergeCells(0, 0, 0, 0, 2, 2);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Verify: First cell should have colSpan="3" and rowSpan="3"
    expect(savedXml).toContain('colSpan="3"');
    expect(savedXml).toContain('rowSpan="3"');

    // Row 0: 1 cell, Row 1: 0 cells, Row 2: 0 cells
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(1);
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(0);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(0);
    }
  });
});

describe('HwpxDocument - Cell Merge (경계 조건)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-merge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should reject invalid range (startRow > endRow)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 0, 2, 0, 0, 0);
    expect(result).toBe(false);
  });

  it('should reject invalid range (startCol > endCol)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 0, 0, 2, 0, 0);
    expect(result).toBe(false);
  });

  it('should reject out of bounds row', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 0, 0, 0, 5, 0);
    expect(result).toBe(false);
  });

  it('should reject out of bounds column', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 0, 0, 0, 0, 5);
    expect(result).toBe(false);
  });

  it('should reject single cell merge (no change needed)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 0, 0, 0, 0, 0);
    expect(result).toBe(false);
  });

  it('should reject invalid section index', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(99, 0, 0, 0, 0, 1);
    expect(result).toBe(false);
  });

  it('should reject invalid table index', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.mergeCells(0, 99, 0, 0, 0, 1);
    expect(result).toBe(false);
  });
});

describe('HwpxDocument - Cell Merge (내용 보존)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-merge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should preserve master cell content after merge', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0) and (0,1)
    doc.mergeCells(0, 0, 0, 0, 0, 1);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Master cell (0,0) content should be preserved
    expect(savedXml).toContain('Cell 0,0');
  });

  it('should preserve XML tag balance after merge', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge 2x2 block
    doc.mergeCells(0, 0, 0, 0, 1, 1);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    if (savedXml) {
      // Check tag balance
      const tblOpen = (savedXml.match(/<hp:tbl\b/g) || []).length;
      const tblClose = (savedXml.match(/<\/hp:tbl>/g) || []).length;
      const trOpen = (savedXml.match(/<hp:tr\b/g) || []).length;
      const trClose = (savedXml.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (savedXml.match(/<hp:tc\b/g) || []).length;
      const tcClose = (savedXml.match(/<\/hp:tc>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);
    }
  });
});

/**
 * Create a 3x3 table with a pre-merged cell for split testing.
 */
async function createMergedCellHwpx(
  mergeStartRow: number,
  mergeStartCol: number,
  mergeEndRow: number,
  mergeEndCol: number
): Promise<Buffer> {
  const zip = new JSZip();
  const rows = 3;
  const cols = 3;

  const colSpan = mergeEndCol - mergeStartCol + 1;
  const rowSpan = mergeEndRow - mergeStartRow + 1;

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

  // Generate table rows with merged cell
  let tableRows = '';
  for (let r = 0; r < rows; r++) {
    tableRows += `    <hp:tr>\n`;
    for (let c = 0; c < cols; c++) {
      // Skip cells that are covered by merge (except master cell)
      const isMasterCell = r === mergeStartRow && c === mergeStartCol;
      const isMergedCell = r >= mergeStartRow && r <= mergeEndRow && c >= mergeStartCol && c <= mergeEndCol;

      if (isMergedCell && !isMasterCell) {
        // Skip this cell - it's covered by merge
        continue;
      }

      if (isMasterCell) {
        tableRows += `      <hp:tc colAddr="${c}" rowAddr="${r}" colSpan="${colSpan}" rowSpan="${rowSpan}">
        <hp:subList>
          <hp:p id="${r * 10 + c}">
            <hp:run><hp:t>Merged Cell</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>\n`;
      } else {
        tableRows += `      <hp:tc colAddr="${c}" rowAddr="${r}" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="${r * 10 + c}">
            <hp:run><hp:t>Cell ${r},${c}</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>\n`;
      }
    }
    tableRows += `    </hp:tr>\n`;
  }

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Document with merged cell</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="${rows}" colCnt="${cols}">
${tableRows}  </hp:tbl>
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

describe('HwpxDocument - Cell Split (가로 분할)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    // Create a table with cells (0,0)-(0,1) merged horizontally
    const buffer = await createMergedCellHwpx(0, 0, 0, 1);
    testFilePath = path.join(__dirname, '..', 'test-split.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should split horizontally merged cell', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Verify merged state
    const savedZip = await JSZip.loadAsync(fs.readFileSync(testFilePath));
    const beforeXml = await savedZip.file('Contents/section0.xml')?.async('string');
    console.log('Before split XML:', beforeXml);
    expect(beforeXml).toContain('colSpan="2"');

    // Split the merged cell at (0,0)
    const result = doc.splitCell(0, 0, 0, 0);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const afterZip = await JSZip.loadAsync(savedBuffer);
    const afterXml = await afterZip.file('Contents/section0.xml')?.async('string');

    console.log('After split XML:', afterXml);

    // Verify: No more colSpan="2" (should be colSpan="1" or no colSpan)
    expect(afterXml).not.toContain('colSpan="2"');

    // Row 0 should now have 3 cells again
    const rows = afterXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      const cellsInRow0 = (rows[0].match(/<hp:tc\b/g) || []).length;
      expect(cellsInRow0).toBe(3);
    }
  });
});

describe('HwpxDocument - Cell Split (세로 분할)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    // Create a table with cells (0,0)-(1,0) merged vertically
    const buffer = await createMergedCellHwpx(0, 0, 1, 0);
    testFilePath = path.join(__dirname, '..', 'test-split.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should split vertically merged cell', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Split the merged cell at (0,0)
    const result = doc.splitCell(0, 0, 0, 0);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const afterZip = await JSZip.loadAsync(savedBuffer);
    const afterXml = await afterZip.file('Contents/section0.xml')?.async('string');

    console.log('After vertical split XML:', afterXml);

    // Verify: No more rowSpan="2"
    expect(afterXml).not.toContain('rowSpan="2"');

    // Row 0 and Row 1 should both have 3 cells
    const rows = afterXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      const cellsInRow0 = (rows[0].match(/<hp:tc\b/g) || []).length;
      const cellsInRow1 = (rows[1].match(/<hp:tc\b/g) || []).length;
      expect(cellsInRow0).toBe(3);
      expect(cellsInRow1).toBe(3);
    }
  });
});

describe('HwpxDocument - Cell Split (블록 분할)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    // Create a table with 2x2 block merged
    const buffer = await createMergedCellHwpx(0, 0, 1, 1);
    testFilePath = path.join(__dirname, '..', 'test-split.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should split 2x2 merged block', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Split the merged cell at (0,0)
    const result = doc.splitCell(0, 0, 0, 0);
    expect(result).toBe(true);

    const savedBuffer = await doc.save();
    const afterZip = await JSZip.loadAsync(savedBuffer);
    const afterXml = await afterZip.file('Contents/section0.xml')?.async('string');

    // Verify: No more colSpan="2" or rowSpan="2"
    expect(afterXml).not.toContain('colSpan="2"');
    expect(afterXml).not.toContain('rowSpan="2"');

    // All rows should have 3 cells
    const rows = afterXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    if (rows) {
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(3);
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(3);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(3);
    }
  });
});

describe('HwpxDocument - Cell Split (경계 조건)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(3, 3);
    testFilePath = path.join(__dirname, '..', 'test-split.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should reject split on non-merged cell', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Cell (0,0) is not merged in this test file
    const result = doc.splitCell(0, 0, 0, 0);
    expect(result).toBe(false);
  });

  it('should reject invalid section index', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.splitCell(99, 0, 0, 0);
    expect(result).toBe(false);
  });

  it('should reject invalid table index', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.splitCell(0, 99, 0, 0);
    expect(result).toBe(false);
  });

  it('should preserve XML tag balance after split', async () => {
    // Use merged cell for this test
    const buffer = await createMergedCellHwpx(0, 0, 1, 1);
    fs.writeFileSync(testFilePath, buffer);

    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));
    doc.splitCell(0, 0, 0, 0);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    if (savedXml) {
      const tblOpen = (savedXml.match(/<hp:tbl\b/g) || []).length;
      const tblClose = (savedXml.match(/<\/hp:tbl>/g) || []).length;
      const trOpen = (savedXml.match(/<hp:tr\b/g) || []).length;
      const trClose = (savedXml.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (savedXml.match(/<hp:tc\b/g) || []).length;
      const tcClose = (savedXml.match(/<\/hp:tc>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);
    }
  });
});

describe('HwpxDocument - Cell Merge/Split Integrity Checks', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTable(4, 4);
    testFilePath = path.join(__dirname, '..', 'test-integrity.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should maintain correct cell count after horizontal merge', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0)-(0,1) horizontally
    doc.mergeCells(0, 0, 0, 0, 0, 1);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Row 0 should have 3 cells (1 merged covering 2 cols + 2 remaining)
    // Row 1-3 should have 4 cells each
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    expect(rows).toHaveLength(4);

    if (rows) {
      // Row 0: 3 cells
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(3);
      // Row 1-3: 4 cells each
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(4);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(4);
      expect((rows[3].match(/<hp:tc\b/g) || []).length).toBe(4);
    }
  });

  it('should maintain correct cell count after vertical merge', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0)-(2,0) vertically (3 rows)
    doc.mergeCells(0, 0, 0, 0, 2, 0);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    expect(rows).toHaveLength(4);

    if (rows) {
      // Row 0: 4 cells (master cell + 3 others)
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(4);
      // Row 1: 3 cells (col 0 is covered by rowSpan)
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(3);
      // Row 2: 3 cells (col 0 is covered by rowSpan)
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(3);
      // Row 3: 4 cells (normal)
      expect((rows[3].match(/<hp:tc\b/g) || []).length).toBe(4);
    }
  });

  it('should preserve colAddr/rowAddr consistency after merge', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge 2x2 block at (1,1)
    doc.mergeCells(0, 0, 1, 1, 2, 2);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    if (savedXml) {
      // Extract all colAddr values
      const colAddrs = [...savedXml.matchAll(/colAddr="(\d+)"/g)].map(m => parseInt(m[1]));
      const rowAddrs = [...savedXml.matchAll(/rowAddr="(\d+)"/g)].map(m => parseInt(m[1]));

      // All colAddr should be valid (0-3)
      expect(colAddrs.every(addr => addr >= 0 && addr <= 3)).toBe(true);
      // All rowAddr should be valid (0-3)
      expect(rowAddrs.every(addr => addr >= 0 && addr <= 3)).toBe(true);
    }
  });

  it('should restore correct cell count after split', async () => {
    // First merge, then split
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge cells (0,0)-(1,1) (2x2 block)
    doc.mergeCells(0, 0, 0, 0, 1, 1);
    await doc.save();

    // Split the merged cell
    doc.splitCell(0, 0, 0, 0);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // All rows should have 4 cells again
    const rows = savedXml?.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
    expect(rows).toHaveLength(4);

    if (rows) {
      expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(4);
      expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(4);
      expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(4);
      expect((rows[3].match(/<hp:tc\b/g) || []).length).toBe(4);
    }
  });

  it('should not corrupt table when merging edge cells', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Merge last row cells (3,0)-(3,3) - entire bottom row
    doc.mergeCells(0, 0, 3, 0, 3, 3);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    if (savedXml) {
      // Verify tag balance
      const tblOpen = (savedXml.match(/<hp:tbl\b/g) || []).length;
      const tblClose = (savedXml.match(/<\/hp:tbl>/g) || []).length;
      const trOpen = (savedXml.match(/<hp:tr\b/g) || []).length;
      const trClose = (savedXml.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (savedXml.match(/<hp:tc\b/g) || []).length;
      const tcClose = (savedXml.match(/<\/hp:tc>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);

      // Row 3 should have only 1 cell with colSpan="4"
      const rows = savedXml.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
      if (rows) {
        expect((rows[3].match(/<hp:tc\b/g) || []).length).toBe(1);
        expect(rows[3]).toContain('colSpan="4"');
      }
    }
  });

  it('should handle multiple merge operations without corruption', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Multiple merges
    doc.mergeCells(0, 0, 0, 0, 0, 1); // Row 0, cols 0-1
    doc.mergeCells(0, 0, 0, 2, 0, 3); // Row 0, cols 2-3
    doc.mergeCells(0, 0, 2, 0, 3, 0); // Rows 2-3, col 0

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    if (savedXml) {
      // Verify tag balance
      const tblOpen = (savedXml.match(/<hp:tbl\b/g) || []).length;
      const tblClose = (savedXml.match(/<\/hp:tbl>/g) || []).length;
      const trOpen = (savedXml.match(/<hp:tr\b/g) || []).length;
      const trClose = (savedXml.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (savedXml.match(/<hp:tc\b/g) || []).length;
      const tcClose = (savedXml.match(/<\/hp:tc>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);

      // Expected cell counts:
      // Row 0: 2 cells (2 merged cells, each covering 2 cols)
      // Row 1: 4 cells (normal)
      // Row 2: 4 cells (col 0 has rowSpan=2)
      // Row 3: 3 cells (col 0 is covered)
      const rows = savedXml.match(/<hp:tr>[\s\S]*?<\/hp:tr>/g);
      if (rows) {
        expect((rows[0].match(/<hp:tc\b/g) || []).length).toBe(2);
        expect((rows[1].match(/<hp:tc\b/g) || []).length).toBe(4);
        expect((rows[2].match(/<hp:tc\b/g) || []).length).toBe(4);
        expect((rows[3].match(/<hp:tc\b/g) || []).length).toBe(3);
      }
    }
  });
});
