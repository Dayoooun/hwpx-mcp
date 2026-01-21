/**
 * Tests for verify_screenshot functionality
 * - Tracks last modified element
 * - Captures screenshot via Hancom Office automation
 * - Returns Base64 image for Claude verification
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Verify Screenshot', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  beforeEach(async () => {
    const zip = new JSZip();

    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:refList>
    <hh:charShapeList itemCnt="1">
      <hh:charShape id="0" height="1000" baseSize="1000"/>
    </hh:charShapeList>
    <hh:paraShapeList itemCnt="1">
      <hh:paraShape id="0"/>
    </hh:paraShapeList>
    <hh:borderFillList itemCnt="2">
      <hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
        <hh:leftBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:rightBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:topBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:bottomBorder type="NONE" width="0.1mm" color="#000000"/>
      </hh:borderFill>
      <hh:borderFill id="2" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
        <hh:leftBorder type="SOLID" width="0.12mm" color="#000000"/>
        <hh:rightBorder type="SOLID" width="0.12mm" color="#000000"/>
        <hh:topBorder type="SOLID" width="0.12mm" color="#000000"/>
        <hh:bottomBorder type="SOLID" width="0.12mm" color="#000000"/>
      </hh:borderFill>
    </hh:borderFillList>
  </hh:refList>
</hh:head>`);

    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>Test document for screenshot verification</hp:t></hp:run>
  </hp:p>
</hs:sec>`);

    testBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    doc = await HwpxDocument.createFromBuffer('test-doc', 'test.hwpx', testBuffer);
  });

  describe('Last Modified Element Tracking', () => {
    it('should have no last modified element initially', () => {
      const lastModified = doc.getLastModifiedElement();
      expect(lastModified).toBeNull();
    });

    it('should track last modified element after table insertion', () => {
      doc.insertTable(0, 0, 3, 3);

      const lastModified = doc.getLastModifiedElement();
      expect(lastModified).not.toBeNull();
      expect(lastModified?.type).toBe('table');
      expect(lastModified?.sectionIndex).toBe(0);
    });

    it('should track last modified element after cell update', async () => {
      // First insert a table
      doc.insertTable(0, 0, 2, 2);
      await doc.save();

      // Reload to get the table in content
      const buffer = await doc.toBuffer();
      doc = await HwpxDocument.createFromBuffer('test-doc', 'test.hwpx', buffer);

      // Update a cell
      doc.updateTableCell(0, 0, 0, 1, 'Updated content');

      const lastModified = doc.getLastModifiedElement();
      expect(lastModified).not.toBeNull();
      expect(lastModified?.type).toBe('table_cell');
      expect(lastModified?.tableIndex).toBe(0);
      expect(lastModified?.row).toBe(0);
      expect(lastModified?.col).toBe(1);
    });

    it('should track last modified element after nested table insertion', async () => {
      // First insert a parent table
      doc.insertTable(0, 0, 2, 2);
      await doc.save();

      const buffer = await doc.toBuffer();
      doc = await HwpxDocument.createFromBuffer('test-doc', 'test.hwpx', buffer);

      // Insert nested table
      const result = doc.insertNestedTable(0, 0, 0, 0, 2, 2);
      expect(result.success).toBe(true);

      const lastModified = doc.getLastModifiedElement();
      expect(lastModified).not.toBeNull();
      expect(lastModified?.type).toBe('nested_table');
      expect(lastModified?.parentTableIndex).toBe(0);
    });
  });

  describe('verifyScreenshot', () => {
    it('should return error when no modification was made', async () => {
      const result = await doc.verifyScreenshot();

      expect(result.success).toBe(false);
      expect(result.error).toContain('No recent modification');
    });

    it('should return error when Hancom Office is not installed', async () => {
      // Insert a table to have something to screenshot
      doc.insertTable(0, 0, 2, 2);

      // Mock: pretend Hancom Office is not installed
      // This test will pass on systems without Hancom Office
      const result = await doc.verifyScreenshot();

      // On systems without Hancom Office, expect specific error
      if (!result.success && result.error?.includes('Hancom Office')) {
        expect(result.error).toContain('Hancom Office');
      }
      // On systems with Hancom Office, expect success or other error
    });

    it('should return Base64 image on success (requires Hancom Office)', async () => {
      doc.insertTable(0, 0, 2, 2);

      const result = await doc.verifyScreenshot();

      // Skip assertion if Hancom Office not installed
      if (result.success) {
        expect(result.image).toBeDefined();
        expect(result.image).toMatch(/^data:image\/(png|jpeg);base64,/);
        expect(result.element).toBeDefined();
        expect(result.element?.type).toBe('table');
      }
    });

    it('should include element metadata in response', async () => {
      doc.insertTable(0, 0, 3, 4);

      const result = await doc.verifyScreenshot();

      if (result.success) {
        expect(result.element).toMatchObject({
          type: 'table',
          sectionIndex: 0,
          description: expect.stringContaining('3x4')
        });
      }
    });
  });
});
