/**
 * Tests for XML repair functionality
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('HwpxDocument - XML Repair', () => {
  let testFilePath: string;
  let tempDir: string;

  beforeEach(async () => {
    // Create temp directory
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-repair-test-'));

    // Create a test HWPX file with corrupted XML
    const zip = new JSZip();

    // Add required files
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/hwp+zip" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/section0.xml"/>
</manifest:manifest>`);
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/packageList">
  <hpf:fileItem>Contents/section0.xml</hpf:fileItem>
</hpf:package>`);
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
</hh:head>`);

    testFilePath = path.join(tempDir, 'test-repair.hwpx');
  });

  afterEach(() => {
    // Cleanup
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  async function createCorruptedFile(sectionXml: string): Promise<void> {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/hwp+zip" manifest:full-path="/"/>
</manifest:manifest>`);
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/packageList">
  <hpf:fileItem>Contents/section0.xml</hpf:fileItem>
</hpf:package>`);
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"></hh:head>`);
    zip.file('Contents/section0.xml', sectionXml);

    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);
  }

  describe('analyze_xml', () => {
    it('should detect orphan closing tags', async () => {
      // Create file with orphan </hp:tbl> tag
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Text</hp:t></hp:run>
  </hp:p>
  </hp:tbl>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const analysis = await doc.analyzeXml();
      expect(analysis.sections).toHaveLength(1);

      const sectionIssues = analysis.sections[0].issues;
      expect(sectionIssues.length).toBeGreaterThan(0);

      // Should detect tbl tag imbalance
      const tblIssue = sectionIssues.find(i => i.message.includes('hp:tbl'));
      expect(tblIssue).toBeDefined();
    });

    it('should detect tr/tc orphan tags', async () => {
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="0"><hp:run><hp:t>Cell</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const analysis = await doc.analyzeXml();
      const sectionIssues = analysis.sections[0].issues;

      // Should detect tc and tr imbalances
      const tcIssue = sectionIssues.find(i => i.message.includes('hp:tc'));
      const trIssue = sectionIssues.find(i => i.message.includes('hp:tr'));

      expect(tcIssue || trIssue).toBeDefined();
    });
  });

  describe('repair_xml', () => {
    it('should remove orphan tbl closing tag', async () => {
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Before</hp:t></hp:run>
  </hp:p>
  </hp:tbl>
  <hp:p id="2">
    <hp:run><hp:t>After</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = await doc.repairXml(0);

      expect(result.success).toBe(true);
      expect(result.repairsApplied.length).toBeGreaterThan(0);

      // Verify repair
      const buffer = await doc.save();
      const zip = await JSZip.loadAsync(buffer);
      const repairedXml = await zip.file('Contents/section0.xml')?.async('string');

      expect(repairedXml).toBeDefined();
      expect(repairedXml).not.toContain('</hp:tbl>');
      expect(repairedXml).toContain('Before');
      expect(repairedXml).toContain('After');
    });

    it('should remove orphan tr/tc closing tags in table', async () => {
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="0"><hp:run><hp:t>Cell</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = await doc.repairXml(0);

      expect(result.success).toBe(true);

      // Verify the XML is now balanced
      const buffer = await doc.save();
      const zip = await JSZip.loadAsync(buffer);
      const repairedXml = await zip.file('Contents/section0.xml')?.async('string');

      expect(repairedXml).toBeDefined();

      // Count tags
      const trOpen = (repairedXml!.match(/<hp:tr/g) || []).length;
      const trClose = (repairedXml!.match(/<\/hp:tr>/g) || []).length;
      const tcOpen = (repairedXml!.match(/<hp:tc/g) || []).length;
      const tcClose = (repairedXml!.match(/<\/hp:tc>/g) || []).length;

      expect(trOpen).toBe(trClose);
      expect(tcOpen).toBe(tcClose);
    });

    it('should handle multiple orphan tags in sequence', async () => {
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="0"><hp:run><hp:t>Data</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  </hp:tc>
  </hp:tr>
  </hp:tbl>
  </hp:tc>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = await doc.repairXml(0);

      expect(result.success).toBe(true);

      // Verify all tags are balanced
      const buffer = await doc.save();
      const zip = await JSZip.loadAsync(buffer);
      const repairedXml = await zip.file('Contents/section0.xml')?.async('string');

      const tblOpen = (repairedXml!.match(/<hp:tbl/g) || []).length;
      const tblClose = (repairedXml!.match(/<\/hp:tbl>/g) || []).length;

      expect(tblOpen).toBe(tblClose);
    });

    it('should preserve valid table structure while removing orphans', async () => {
      const corruptedXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="0"><hp:run><hp:t>A1</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="1"><hp:run><hp:t>B1</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    </hp:tc>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="2"><hp:run><hp:t>A2</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="3"><hp:run><hp:t>B2</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;

      await createCorruptedFile(corruptedXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = await doc.repairXml(0);

      expect(result.success).toBe(true);

      // Verify content is preserved
      const buffer = await doc.save();
      const zip = await JSZip.loadAsync(buffer);
      const repairedXml = await zip.file('Contents/section0.xml')?.async('string');

      expect(repairedXml).toContain('A1');
      expect(repairedXml).toContain('B1');
      expect(repairedXml).toContain('A2');
      expect(repairedXml).toContain('B2');

      // Verify structure
      const tcOpen = (repairedXml!.match(/<hp:tc/g) || []).length;
      const tcClose = (repairedXml!.match(/<\/hp:tc>/g) || []).length;
      expect(tcOpen).toBe(tcClose);
      expect(tcOpen).toBe(4); // Should have 4 cells
    });

    it('should report no issues when XML is valid', async () => {
      const validXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>Valid document</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="0"><hp:run><hp:t>Cell</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;

      await createCorruptedFile(validXml);
      const doc = await HwpxDocument.createFromBuffer('test-id', testFilePath, fs.readFileSync(testFilePath));

      const result = await doc.repairXml(0);

      expect(result.success).toBe(true);
      expect(result.message).toBe('No issues to repair');
      expect(result.repairsApplied).toHaveLength(0);
    });
  });
});
