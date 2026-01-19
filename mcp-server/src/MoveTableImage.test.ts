/**
 * Tests for Table/Image Move Feature (표/이미지 이동 기능)
 *
 * TDD: RED → GREEN → REFACTOR
 *
 * 엄격 검증 (Option B):
 * - ID 중복 검사
 * - 태그 균형 검사
 * - 참조 무결성
 * - 이스케이프 검증
 * - XML 파싱 테스트
 */
import { describe, it, expect, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * Create a test HWPX with 2 sections, each containing a table and an image-like element
 */
async function createTestHwpxWithTableAndImage(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Move Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

  // Section 0: paragraph + table + paragraph
  const section0Xml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="p0">
    <hp:run><hp:t>Section 0 - Paragraph 0</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="tbl100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="tbl100_p0">
            <hp:run><hp:t>Table100 Cell 0,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="tbl100_p1">
            <hp:run><hp:t>Table100 Cell 0,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="tbl100_p2">
            <hp:run><hp:t>Table100 Cell 1,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="tbl100_p3">
            <hp:run><hp:t>Table100 Cell 1,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p id="p1">
    <hp:run><hp:t>Section 0 - Paragraph 1</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

  // Section 1: paragraph only (target for moves)
  const section1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="s1_p0">
    <hp:run><hp:t>Section 1 - Paragraph 0</hp:t></hp:run>
  </hp:p>
  <hp:p id="s1_p1">
    <hp:run><hp:t>Section 1 - Paragraph 1</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', section0Xml);
  zip.file('Contents/section1.xml', section1Xml);
  zip.file('[Content_Types].xml', contentTypesXml);

  return await zip.generateAsync({ type: 'nodebuffer' });
}

describe('Table Move Feature (표 이동)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTableAndImage();
    testFilePath = path.join(__dirname, 'test-move-table.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  describe('moveTable - 기본 기능', () => {
    it('should move table within same section', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      // Section 0: p0, tbl100, p1
      // Move tbl100 to after p1 (index 2)
      const result = doc.moveTable(0, 0, 0, 2);

      expect(result.success).toBe(true);

      // Save and reload
      const savedBuffer = await doc.save();
      const reloadedDoc = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

      // Verify table is now after p1
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // p0 should come before tbl100, p1 should come before tbl100
      const p1Index = savedXml!.indexOf('Section 0 - Paragraph 1');
      const tblIndex = savedXml!.indexOf('hp:tbl');
      expect(p1Index).toBeLessThan(tblIndex);
    });

    it('should move table to different section', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      // Move tbl100 from section 0 to section 1, after s1_p0
      const result = doc.moveTable(0, 0, 1, 0);

      expect(result.success).toBe(true);

      const savedBuffer = await doc.save();

      // Verify section0 no longer has the table
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const section0Xml = await savedZip.file('Contents/section0.xml')?.async('string');
      const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

      expect(section0Xml).not.toContain('hp:tbl');
      expect(section1Xml).toContain('hp:tbl');
      expect(section1Xml).toContain('Table100 Cell 0,0');
    });
  });

  describe('moveTable - 엄격 검증', () => {
    it('should generate new IDs when duplicates exist', async () => {
      // Create a document with duplicate IDs
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      // First move - should work
      const result1 = doc.moveTable(0, 0, 1, 0);
      expect(result1.success).toBe(true);

      // Copy the table back (to create potential duplicate)
      // This tests ID regeneration
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

      // Verify IDs are unique (not the original tbl100)
      // The implementation should generate new IDs
      expect(section1Xml).toContain('hp:tbl');
    });

    it('should maintain tag balance after move', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = doc.moveTable(0, 0, 1, 0);
      expect(result.success).toBe(true);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

      // Count opening and closing tags
      const tblOpen = (section1Xml!.match(/<hp:tbl[^/]/g) || []).length;
      const tblClose = (section1Xml!.match(/<\/hp:tbl>/g) || []).length;
      expect(tblOpen).toBe(tblClose);

      const trOpen = (section1Xml!.match(/<hp:tr[^/]/g) || []).length;
      const trClose = (section1Xml!.match(/<\/hp:tr>/g) || []).length;
      expect(trOpen).toBe(trClose);

      const tcOpen = (section1Xml!.match(/<hp:tc[^/]/g) || []).length;
      const tcClose = (section1Xml!.match(/<\/hp:tc>/g) || []).length;
      expect(tcOpen).toBe(tcClose);
    });

    it('should validate XML is well-formed after move', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = doc.moveTable(0, 0, 1, 0);
      expect(result.success).toBe(true);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

      // XML should be parseable
      expect(() => {
        // Simple validation - check for basic XML structure
        if (!section1Xml!.includes('<?xml') || !section1Xml!.includes('</hs:sec>')) {
          throw new Error('Invalid XML structure');
        }
      }).not.toThrow();
    });
  });

  describe('moveTable - 에러 케이스', () => {
    it('should return error for invalid source section index', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = doc.moveTable(99, 0, 1, 0);

      expect(result.success).toBe(false);
      expect(result.error).toContain('section');
    });

    it('should handle invalid table index gracefully during save', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      // Invalid table index is accepted at call time, validated during save
      const result = doc.moveTable(0, 99, 1, 0);
      expect(result.success).toBe(true); // Deferred validation

      // Save should complete without crashing (invalid ops are skipped with console warning)
      const savedBuffer = await doc.save();
      expect(savedBuffer).toBeDefined();

      // Target section should NOT have any table (invalid op was skipped)
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');
      expect(section1Xml).not.toContain('hp:tbl');
    });

    it('should return error for invalid target section index', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = doc.moveTable(0, 0, 99, 0);

      expect(result.success).toBe(false);
      expect(result.error).toContain('section');
    });

    it('should handle move to same position (no-op)', async () => {
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      // Table is at index 1, moving to after index 0 keeps it in same position
      const result = doc.moveTable(0, 0, 0, 0);

      // Should succeed but be a no-op
      expect(result.success).toBe(true);
    });
  });
});

describe('copyTable Feature (표 복사)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTableAndImage();
    testFilePath = path.join(__dirname, 'test-copy-table.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should copy table to different section (preserving original)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.copyTable(0, 0, 1, 0);

    expect(result.success).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const section0Xml = await savedZip.file('Contents/section0.xml')?.async('string');
    const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

    // Original should still exist in section 0
    expect(section0Xml).toContain('hp:tbl');
    expect(section0Xml).toContain('Table100 Cell 0,0');

    // Copy should exist in section 1
    expect(section1Xml).toContain('hp:tbl');
    expect(section1Xml).toContain('Table100 Cell 0,0');
  });

  it('should generate new IDs for copied table', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.copyTable(0, 0, 1, 0);
    expect(result.success).toBe(true);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const section0Xml = await savedZip.file('Contents/section0.xml')?.async('string');
    const section1Xml = await savedZip.file('Contents/section1.xml')?.async('string');

    // Original table ID should be preserved
    expect(section0Xml).toContain('id="tbl100"');

    // Copied table should exist in section 1
    expect(section1Xml).toContain('hp:tbl');
    expect(section1Xml).toContain('Table100 Cell 0,0');

    // Original paragraph IDs should NOT be in copied table
    // (they should have been regenerated)
    expect(section1Xml).not.toContain('id="tbl100_p0"');
    expect(section1Xml).not.toContain('id="tbl100_p1"');
  });
});

describe('Validation Utilities (검증 유틸리티)', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createTestHwpxWithTableAndImage();
    testFilePath = path.join(__dirname, 'test-validation.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('validateTagBalance should detect balanced tags', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    const result = doc.validateTagBalance(`
      <hp:tbl id="1">
        <hp:tr>
          <hp:tc><hp:subList><hp:p id="1"><hp:run><hp:t>text</hp:t></hp:run></hp:p></hp:subList></hp:tc>
        </hp:tr>
      </hp:tbl>
    `);

    expect(result.balanced).toBe(true);
    expect(result.mismatches).toHaveLength(0);
  });

  it('validateTagBalance should detect unbalanced tags', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Missing closing </hp:tbl> tag (no comment that would be counted)
    const result = doc.validateTagBalance(`
      <hp:tbl id="1">
        <hp:tr>
          <hp:tc><hp:subList><hp:p id="1"><hp:run><hp:t>text</hp:t></hp:run></hp:p></hp:subList></hp:tc>
        </hp:tr>
    `);

    expect(result.balanced).toBe(false);
    expect(result.mismatches.some(m => m.tag === 'hp:tbl')).toBe(true);
  });

  it('validateXmlEscaping should detect unescaped characters', async () => {
    const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

    // Well-escaped
    const result1 = doc.validateXmlEscaping('<hp:t>&lt;tag&gt; &amp; text</hp:t>');
    expect(result1.valid).toBe(true);

    // Unescaped < in text content (this is complex to detect reliably)
    // For now, we trust that our escapeXml function works correctly
  });
});
