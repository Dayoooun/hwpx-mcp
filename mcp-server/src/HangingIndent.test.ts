/**
 * Tests for Hanging Indent (내어쓰기) functionality
 *
 * HWPML에서 내어쓰기는 다음과 같이 동작:
 * - intent (음수): 첫 줄을 왼쪽으로 당김
 * - left (양수): 기본 왼쪽 여백 (나머지 줄에 적용)
 *
 * 예: intent=-1312, left=1312 → 첫줄 0pt, 나머지 13.12pt
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('HwpxDocument - Hanging Indent (내어쓰기)', () => {
  let testFilePath: string;
  let tempDir: string;

  beforeEach(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-hanging-indent-test-'));
  });

  afterEach(() => {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  /**
   * 테스트용 HWPX 파일 생성
   * header.xml에 기본 paraPr 정의 포함
   */
  async function createTestFile(): Promise<HwpxDocument> {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');

    // META-INF/manifest.xml
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/hwp+zip" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/section0.xml"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/header.xml"/>
</manifest:manifest>`);

    // Contents/content.hpf
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/packageList">
  <hpf:fileItem>Contents/header.xml</hpf:fileItem>
  <hpf:fileItem>Contents/section0.xml</hpf:fileItem>
</hpf:package>`);

    // Contents/header.xml - 기본 paraPr 정의 포함
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hh:refList>
    <hh:fontfaces itemCnt="1">
      <hh:fontface lang="HANGUL" fontCnt="1">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    <hh:charProperties itemCnt="1">
      <hh:charPr id="0" height="1000" textColor="#000000"/>
    </hh:charProperties>
    <hh:tabProperties itemCnt="1">
      <hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>
    </hh:tabProperties>
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
        <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
      </hh:paraPr>
    </hh:paraProperties>
    <hh:styles itemCnt="1">
      <hh:style id="0" type="PARA" name="바탕글" paraPrIDRef="0" charPrIDRef="0"/>
    </hh:styles>
  </hh:refList>
</hh:head>`);

    // Contents/section0.xml - 테스트용 단락 2개
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>1. 첫 번째 항목입니다. 이것은 긴 텍스트로 여러 줄에 걸쳐 표시될 수 있습니다.</hp:t></hp:run>
  </hp:p>
  <hp:p id="2" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>2. 두 번째 항목입니다.</hp:t></hp:run>
  </hp:p>
</hs:sec>`);

    testFilePath = path.join(tempDir, 'test-hanging-indent.hwpx');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);

    return HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);
  }

  describe('setHangingIndent', () => {
    it('should set hanging indent on a paragraph', async () => {
      const doc = await createTestFile();

      // 내어쓰기 설정: 15pt (1500 HWPUNIT)
      const result = doc.setHangingIndent(0, 0, 15);

      expect(result).toBe(true);

      // 단락 스타일 확인
      const style = doc.getParagraphStyle(0, 0);
      expect(style?.firstLineIndent).toBe(-15); // 음수 (내어쓰기)
      expect(style?.marginLeft).toBe(15);       // 양수 (기본 여백)
    });

    it('should persist hanging indent after save', async () => {
      const doc = await createTestFile();

      // 내어쓰기 설정
      doc.setHangingIndent(0, 0, 20);

      // 저장
      const savedPath = path.join(tempDir, 'saved-hanging-indent.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(savedPath, savedBuffer);

      // 저장된 파일의 header.xml 확인
      const savedData = fs.readFileSync(savedPath);
      const savedZip = await JSZip.loadAsync(savedData);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 새 paraPr이 추가되었는지 확인 (음수 intent)
      expect(headerXml).toContain('intent value="-2000"'); // -20pt = -2000 HWPUNIT
      expect(headerXml).toContain('left value="2000"');    // 20pt = 2000 HWPUNIT
    });

    it('should update paragraph paraPrIDRef after setting hanging indent', async () => {
      const doc = await createTestFile();

      // 내어쓰기 설정
      doc.setHangingIndent(0, 0, 10);

      // 저장
      const savedPath = path.join(tempDir, 'saved-para-ref.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(savedPath, savedBuffer);

      // section XML에서 단락의 paraPrIDRef가 새 ID를 참조하는지 확인
      const savedData = fs.readFileSync(savedPath);
      const savedZip = await JSZip.loadAsync(savedData);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // 첫 번째 단락의 paraPrIDRef가 0이 아닌 새 ID로 변경되어야 함
      expect(sectionXml).toMatch(/<hp:p[^>]*id="1"[^>]*paraPrIDRef="[1-9]\d*"/);
    });

    it('should reject invalid section index', async () => {
      const doc = await createTestFile();

      const result = doc.setHangingIndent(99, 0, 15);

      expect(result).toBe(false);
    });

    it('should reject invalid paragraph index', async () => {
      const doc = await createTestFile();

      const result = doc.setHangingIndent(0, 99, 15);

      expect(result).toBe(false);
    });

    it('should reject zero or negative indent', async () => {
      const doc = await createTestFile();

      expect(doc.setHangingIndent(0, 0, 0)).toBe(false);
      expect(doc.setHangingIndent(0, 0, -5)).toBe(false);
    });
  });

  describe('getHangingIndent', () => {
    it('should return hanging indent value for a paragraph', async () => {
      const doc = await createTestFile();

      // 내어쓰기 설정
      doc.setHangingIndent(0, 0, 25);

      // 내어쓰기 값 조회
      const indent = doc.getHangingIndent(0, 0);

      expect(indent).toBe(25);
    });

    it('should return 0 for paragraph without hanging indent', async () => {
      const doc = await createTestFile();

      const indent = doc.getHangingIndent(0, 0);

      expect(indent).toBe(0);
    });

    it('should return null for invalid indices', async () => {
      const doc = await createTestFile();

      expect(doc.getHangingIndent(99, 0)).toBeNull();
      expect(doc.getHangingIndent(0, 99)).toBeNull();
    });
  });

  describe('removeHangingIndent', () => {
    it('should remove hanging indent from a paragraph', async () => {
      const doc = await createTestFile();

      // 먼저 내어쓰기 설정
      doc.setHangingIndent(0, 0, 15);
      expect(doc.getHangingIndent(0, 0)).toBe(15);

      // 내어쓰기 제거
      const result = doc.removeHangingIndent(0, 0);

      expect(result).toBe(true);
      expect(doc.getHangingIndent(0, 0)).toBe(0);
    });
  });

  describe('multiple paragraphs (다중 단락)', () => {
    it('should apply same indent to multiple paragraphs and reuse paraPr', async () => {
      const doc = await createTestFile();

      // 두 단락에 같은 내어쓰기 적용
      doc.setHangingIndent(0, 0, 15);
      doc.setHangingIndent(0, 1, 15);

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // paraPr ID가 하나만 추가되었는지 확인 (같은 indent 값은 재사용)
      const intentMatches = headerXml?.match(/intent value="-1500"/g);
      expect(intentMatches?.length).toBe(1);

      // 두 단락 모두 같은 paraPrIDRef를 참조해야 함
      const paraMatches = sectionXml?.match(/paraPrIDRef="(\d+)"/g);
      expect(paraMatches?.length).toBeGreaterThanOrEqual(2);
    });

    it('should create separate paraPr for different indent values', async () => {
      const doc = await createTestFile();

      // 두 단락에 다른 내어쓰기 적용
      doc.setHangingIndent(0, 0, 10);
      doc.setHangingIndent(0, 1, 20);

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 두 개의 다른 intent 값이 있어야 함
      expect(headerXml).toContain('intent value="-1000"'); // 10pt
      expect(headerXml).toContain('intent value="-2000"'); // 20pt
    });

    it('should use last value when same paragraph is changed multiple times', async () => {
      const doc = await createTestFile();

      // 같은 단락에 여러 번 내어쓰기 변경
      doc.setHangingIndent(0, 0, 10);
      doc.setHangingIndent(0, 0, 15);
      doc.setHangingIndent(0, 0, 20); // 최종 값

      // 메모리 상태 확인
      expect(doc.getHangingIndent(0, 0)).toBe(20);

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 최종 값인 20pt만 저장되어야 함
      expect(headerXml).toContain('intent value="-2000"');
      // 중간 값들은 저장되지 않아야 함 (최적화)
      expect(headerXml).not.toContain('intent value="-1000"');
      expect(headerXml).not.toContain('intent value="-1500"');
    });
  });
});
