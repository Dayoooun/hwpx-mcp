import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';

/**
 * TDD 테스트: 독립 단락 텍스트 업데이트
 *
 * 문제 상황:
 * - 테이블 셀 업데이트: lineseg 초기화 O → 정상 작동
 * - 독립 단락 업데이트: lineseg 초기화 X → 텍스트 겹침
 *
 * 해결 목표:
 * - 독립 단락 업데이트 시에도 lineseg 초기화
 */
describe('Paragraph Text Update - lineseg reset', () => {
  const testDir = path.join(__dirname, 'test-temp-para');
  const testFile = path.join(testDir, 'test-paragraph.hwpx');

  beforeEach(() => {
    if (!fs.existsSync(testDir)) {
      fs.mkdirSync(testDir, { recursive: true });
    }
  });

  afterEach(() => {
    if (fs.existsSync(testDir)) {
      fs.rmSync(testDir, { recursive: true, force: true });
    }
  });

  it('should reset lineseg when updating paragraph text', async () => {
    // 1. 새 문서 생성 (정적 메서드 사용)
    const doc = HwpxDocument.createNew('test-para-1');

    // 2. 단락 삽입
    const paraIndex = doc.insertParagraph(0, -1, '초기 텍스트');
    expect(paraIndex).toBeGreaterThanOrEqual(0);

    // 3. 단락 텍스트 업데이트
    doc.updateParagraphText(0, paraIndex, 0, '업데이트된 긴 텍스트입니다. 이 텍스트는 lineseg가 초기화되어야 합니다.');

    // 4. 저장
    const buffer = await doc.save();
    fs.writeFileSync(testFile, buffer);

    // 5. 저장된 파일의 XML 확인
    const zipData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(zipData);
    const sectionXml = await zip.file('Contents/section0.xml')?.async('string');

    expect(sectionXml).toBeDefined();

    // lineseg가 기본값으로 초기화되어야 함 (horzsize="0")
    // 한글이 파일을 열 때 자동으로 재계산
    const linesegMatch = sectionXml!.match(/<(?:hp|hs|hc):lineseg[^>]*horzsize="(\d+)"[^>]*>/);
    if (linesegMatch) {
      // horzsize가 0이면 초기화됨
      expect(linesegMatch[1]).toBe('0');
    }
  });

  it('should handle long text with chunking', async () => {
    const doc = HwpxDocument.createNew('test-para-2');

    const paraIndex = doc.insertParagraph(0, -1, '초기');

    // 긴 텍스트 (500자 이상)
    const longText = 'A'.repeat(600) + '끝';
    doc.updateParagraphText(0, paraIndex, 0, longText);

    const buffer = await doc.save();
    fs.writeFileSync(testFile, buffer);

    const zipData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(zipData);
    const sectionXml = await zip.file('Contents/section0.xml')?.async('string');

    expect(sectionXml).toBeDefined();
    // 저장된 텍스트가 원본과 일치해야 함 (청크로 나뉘더라도)
    expect(sectionXml).toContain('AAAA');
    expect(sectionXml).toContain('끝');
  });

  it('should handle text with newlines', async () => {
    const doc = HwpxDocument.createNew('test-para-3');

    const paraIndex = doc.insertParagraph(0, -1, '초기');

    // 줄바꿈이 있는 텍스트
    const multilineText = '첫 번째 줄\n두 번째 줄\n세 번째 줄';
    doc.updateParagraphText(0, paraIndex, 0, multilineText);

    const buffer = await doc.save();
    fs.writeFileSync(testFile, buffer);

    const zipData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(zipData);
    const sectionXml = await zip.file('Contents/section0.xml')?.async('string');

    expect(sectionXml).toBeDefined();
    // 각 줄이 별도의 단락으로 또는 같은 단락 내에 있어야 함
    expect(sectionXml).toContain('첫 번째 줄');
    expect(sectionXml).toContain('두 번째 줄');
    expect(sectionXml).toContain('세 번째 줄');
  });
});

describe('Comparison: Table Cell vs Paragraph Update', () => {
  const testDir = path.join(__dirname, 'test-temp-compare');
  const testFile = path.join(testDir, 'test-compare.hwpx');

  beforeEach(() => {
    if (!fs.existsSync(testDir)) {
      fs.mkdirSync(testDir, { recursive: true });
    }
  });

  afterEach(() => {
    if (fs.existsSync(testDir)) {
      fs.rmSync(testDir, { recursive: true, force: true });
    }
  });

  it('should treat paragraph update same as table cell update for lineseg', async () => {
    const doc = HwpxDocument.createNew('test-compare-1');

    // 테이블 삽입
    const tableResult = doc.insertTable(0, -1, 2, 2);
    expect(tableResult).not.toBeNull();

    // 단락 삽입
    const paraIndex = doc.insertParagraph(0, 0, '단락 텍스트');

    // 테이블 셀 업데이트
    doc.updateTableCell(0, 0, 0, 0, '셀 텍스트 업데이트');

    // 단락 텍스트 업데이트
    doc.updateParagraphText(0, paraIndex, 0, '단락 텍스트 업데이트');

    const buffer = await doc.save();
    fs.writeFileSync(testFile, buffer);

    const zipData = fs.readFileSync(testFile);
    const zip = await JSZip.loadAsync(zipData);
    const sectionXml = await zip.file('Contents/section0.xml')?.async('string');

    expect(sectionXml).toBeDefined();

    // 두 업데이트 모두 lineseg가 초기화되어야 함
    // linesegarray 내의 모든 lineseg의 horzsize가 0이어야 함
    const linesegMatches = sectionXml!.matchAll(/<(?:hp|hs|hc):lineseg[^>]*horzsize="(\d+)"[^>]*>/g);
    for (const match of linesegMatches) {
      // 새로 업데이트된 텍스트의 lineseg는 horzsize=0이어야 함
      expect(match[1]).toBe('0');
    }
  });
});
