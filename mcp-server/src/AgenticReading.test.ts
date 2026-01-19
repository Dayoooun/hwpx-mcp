/**
 * Tests for Agentic Document Reading functionality
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('HwpxDocument - Agentic Reading', () => {
  let testFilePath: string;
  let tempDir: string;

  beforeEach(async () => {
    // Create temp directory
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-agentic-test-'));
  });

  afterEach(() => {
    // Cleanup
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  async function createTestFile(sectionXml: string): Promise<HwpxDocument> {
    const zip = new JSZip();
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
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"></hh:head>`);
    zip.file('Contents/section0.xml', sectionXml);

    testFilePath = path.join(tempDir, 'test-agentic.hwpx');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);

    return HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);
  }

  describe('chunkDocument', () => {
    it('should chunk a simple document', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>1. 첫 번째 제목</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run><hp:t>첫 번째 내용입니다. 이것은 긴 문장입니다. 더 많은 텍스트가 여기에 있습니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run><hp:t>가. 두 번째 제목</hp:t></hp:run>
  </hp:p>
  <hp:p id="4">
    <hp:run><hp:t>두 번째 내용입니다.</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const chunks = doc.chunkDocument(100, 20);

      expect(chunks.length).toBeGreaterThan(0);
      expect(chunks[0].id).toMatch(/^chunk_/);
      expect(chunks[0].sectionIndex).toBe(0);
      expect(chunks[0].metadata.charCount).toBeGreaterThan(0);
    });

    it('should respect chunk size and overlap', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>가나다라마바사아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사아자차카타파하</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const chunks = doc.chunkDocument(50, 10);

      // Check that chunks are within size limits (may be slightly over for last chunk)
      for (let i = 0; i < chunks.length - 1; i++) {
        expect(chunks[i].text.length).toBeLessThanOrEqual(51);
      }
    });
  });

  describe('searchChunks', () => {
    it('should find relevant chunks by keyword', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>안녕하세요 테스트 문서입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run><hp:t>이것은 프로젝트 계획서입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run><hp:t>다른 내용이 있습니다.</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const results = doc.searchChunks('프로젝트 계획', 3);

      expect(results.length).toBeGreaterThan(0);
      expect(results[0].score).toBeGreaterThan(0);
      expect(results[0].matchedTerms.length).toBeGreaterThan(0);
    });
  });

  describe('extractToc', () => {
    it('should extract headings from document', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>1. 개요</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run><hp:t>개요 내용입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run><hp:t>가. 배경</hp:t></hp:run>
  </hp:p>
  <hp:p id="4">
    <hp:run><hp:t>배경 내용입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="5">
    <hp:run><hp:t>2. 목적</hp:t></hp:run>
  </hp:p>
  <hp:p id="6">
    <hp:run><hp:t>① 첫 번째 목적</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const toc = doc.extractToc();

      expect(toc.length).toBeGreaterThanOrEqual(4);

      // Check that "1. 개요" is detected as level 1
      const heading1 = toc.find(t => t.title.includes('1. 개요'));
      expect(heading1).toBeDefined();
      expect(heading1?.level).toBe(1);

      // Check that "가. 배경" is detected as level 2
      const heading2 = toc.find(t => t.title.includes('가. 배경'));
      expect(heading2).toBeDefined();
      expect(heading2?.level).toBe(2);

      // Check that "① 첫 번째" is detected as level 3
      const heading3 = toc.find(t => t.title.includes('①'));
      expect(heading3).toBeDefined();
      expect(heading3?.level).toBe(3);
    });
  });

  describe('positionIndex', () => {
    it('should build and search position index', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>1. 프로젝트 개요</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run><hp:t>이 프로젝트는 문서 처리를 위한 것입니다.</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="3"><hp:run><hp:t>항목</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="4"><hp:run><hp:t>값</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="5"><hp:run><hp:t>테스트</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="6"><hp:run><hp:t>123</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const index = doc.buildPositionIndex();

      expect(index.length).toBeGreaterThan(0);

      // Check that heading is detected
      const heading = index.find(e => e.type === 'heading' && e.text.includes('프로젝트 개요'));
      expect(heading).toBeDefined();
      expect(heading?.level).toBe(1);

      // Check that table is detected
      const table = index.find(e => e.type === 'table');
      expect(table).toBeDefined();
      expect(table?.tableInfo?.rows).toBe(2);
      expect(table?.tableInfo?.cols).toBe(2);

      // Search the index
      const searchResults = doc.searchPositionIndex('프로젝트');
      expect(searchResults.length).toBeGreaterThan(0);

      // Search with type filter
      const headingResults = doc.searchPositionIndex('프로젝트', 'heading');
      expect(headingResults.every(r => r.type === 'heading')).toBe(true);
    });
  });

  describe('getChunkContext', () => {
    it('should return surrounding chunks', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>첫 번째 단락입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="2">
    <hp:run><hp:t>두 번째 단락입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="3">
    <hp:run><hp:t>세 번째 단락입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="4">
    <hp:run><hp:t>네 번째 단락입니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="5">
    <hp:run><hp:t>다섯 번째 단락입니다.</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);
      const chunks = doc.chunkDocument(30, 5);

      if (chunks.length >= 3) {
        const middleChunk = chunks[Math.floor(chunks.length / 2)];
        const context = doc.getChunkContext(middleChunk.id, 1, 1);

        expect(context.chunks.length).toBeGreaterThanOrEqual(1);
        expect(context.centerIndex).toBeGreaterThanOrEqual(0);
      }
    });
  });

  describe('invalidateReadingCache', () => {
    it('should clear cached data', async () => {
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1">
    <hp:run><hp:t>테스트 내용</hp:t></hp:run>
  </hp:p>
</hs:sec>`;

      const doc = await createTestFile(sectionXml);

      // Build caches
      doc.chunkDocument();
      doc.buildPositionIndex();

      // Invalidate
      doc.invalidateReadingCache();

      // After invalidation, next call should rebuild
      const newChunks = doc.chunkDocument();
      expect(newChunks.length).toBeGreaterThan(0);
    });
  });
});
