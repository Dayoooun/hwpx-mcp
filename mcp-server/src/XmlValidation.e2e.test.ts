import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * E2E test for XML validation in HwpxDocument.setSectionXml()
 */

async function createValidTestHwpx(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Test</hh:title></hh:docInfo>
</hh:head>`;

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Valid paragraph</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', sectionXml);
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`);

  return await zip.generateAsync({ type: 'nodebuffer' });
}

describe('XML Validation E2E - setSectionXml', () => {
  let testFilePath: string;
  let doc: HwpxDocument;

  beforeEach(async () => {
    testFilePath = path.join(__dirname, '..', 'test-output', 'xml-validation-test.hwpx');
    const buffer = await createValidTestHwpx();
    fs.writeFileSync(testFilePath, buffer);
    const fileBuffer = fs.readFileSync(testFilePath);
    doc = await HwpxDocument.createFromBuffer('xml-test', testFilePath, fileBuffer);
  });

  afterEach(() => {
    if (fs.existsSync(testFilePath)) {
      fs.unlinkSync(testFilePath);
    }
  });

  describe('Invalid XML Detection', () => {
    it('should reject non-XML content', async () => {
      const result = await doc.setSectionXml(0, 'This is not XML at all');
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Missing section root element/);
    });

    it('should reject XML without HWPML root element', async () => {
      const invalidXml = `<?xml version="1.0"?><root><item>test</item></root>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Missing section root element/);
    });

    it('should reject XML with missing namespace', async () => {
      const invalidXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://example.com/wrong">
  <hp:p id="1"><hp:run><hp:t>Test</hp:t></hp:run></hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Missing required HWPML namespace/);
    });

    it('should reject XML with unbalanced paragraph tags', async () => {
      const invalidXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Unclosed paragraph</hp:t></hp:run>
</hs:sec>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Tag imbalance for <p>/);
    });

    it('should reject XML with unbalanced table tags', async () => {
      const invalidXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl>
    <hp:tr>
      <hp:tc>
        <hp:p id="1"><hp:run><hp:t>Cell</hp:t></hp:run></hp:p>
      </hp:tc>
    </hp:tr>
</hs:sec>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Tag imbalance for <tbl>/);
    });

    it('should reject XML with invalid control characters', async () => {
      const invalidXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Text with\x01control char</hp:t></hp:run>
  </hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Invalid control character/);
    });

    it('should reject XML with mismatched angle brackets', async () => {
      const invalidXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Missing bracket</hp:t></hp:run
  </hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, invalidXml);
      expect(result.success).toBe(false);
      expect(result.error).toMatch(/Mismatched angle brackets/);
    });
  });

  describe('Valid XML Acceptance', () => {
    it('should accept valid HWPML XML with hs:sec root', async () => {
      const validXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Valid text</hp:t></hp:run>
  </hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, validXml);
      expect(result.success).toBe(true);
    });

    it('should accept valid HWPML XML with Korean text', async () => {
      const validXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>한글 테스트 텍스트</hp:t></hp:run>
  </hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, validXml);
      expect(result.success).toBe(true);
    });

    it('should accept valid XML with table structure', async () => {
      const validXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl>
    <hp:tr>
      <hp:tc>
        <hp:p id="1"><hp:run><hp:t>Cell 1</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p id="2"><hp:run><hp:t>Cell 2</hp:t></hp:run></hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;
      const result = await doc.setSectionXml(0, validXml);
      expect(result.success).toBe(true);
    });

    it('should accept valid XML with self-closing tags', async () => {
      const validXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t/></hp:run>
  </hp:p>
</hs:sec>`;
      const result = await doc.setSectionXml(0, validXml);
      expect(result.success).toBe(true);
    });
  });

  describe('Integration - Save and Reload', () => {
    it('should save valid XML and reload successfully', async () => {
      const validXml = `<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Modified content for save test</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const result = await doc.setSectionXml(0, validXml);
      expect(result.success).toBe(true);

      // Save
      const savePath = path.join(__dirname, '..', 'test-output', 'xml-validation-saved.hwpx');

      // Ensure test-output directory exists
      const testOutputDir = path.dirname(savePath);
      if (!fs.existsSync(testOutputDir)) {
        fs.mkdirSync(testOutputDir, { recursive: true });
      }

      try {
        const savedBuffer = await doc.save();
        fs.writeFileSync(savePath, savedBuffer);
      } catch (saveError) {
        throw new Error(`Save failed: ${saveError}`);
      }

      // Verify file was created
      expect(fs.existsSync(savePath)).toBe(true);

      // Reload
      const savedBuffer = fs.readFileSync(savePath);
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload-test', savePath, savedBuffer);

      const para = reloadedDoc.getParagraph(0, 0);
      expect(para?.text).toBe('Modified content for save test');

      // Cleanup
      fs.unlinkSync(savePath);
    });
  });
});
