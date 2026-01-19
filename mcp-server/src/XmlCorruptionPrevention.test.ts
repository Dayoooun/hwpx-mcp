import { describe, it, expect, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * XML 손상 방지 테스트
 *
 * 테스트 가설:
 * 1. 중첩 테이블이 있는 셀 업데이트 시 중첩 테이블 보존
 * 2. 멀티라인 텍스트 업데이트 시 중첩 테이블 보존
 * 3. 같은 테이블의 여러 행 업데이트 시 태그 균형 유지
 * 4. 같은 행의 여러 셀 업데이트 시 태그 균형 유지
 * 5. 태그 균형 검증이 손상 감지
 */

/**
 * 기본 테스트용 HWPX 파일 생성
 */
async function createBasicTestHwpx(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Test</hh:title></hh:docInfo>
</hh:head>`;

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="3" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="10"><hp:run charPrIDRef="0"><hp:t>R0C0</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="11"><hp:run charPrIDRef="0"><hp:t>R0C1</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="20"><hp:run charPrIDRef="0"><hp:t>R1C0</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="21"><hp:run charPrIDRef="0"><hp:t>R1C1</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="2" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="30"><hp:run charPrIDRef="0"><hp:t>R2C0</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="2" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="31"><hp:run charPrIDRef="0"><hp:t>R2C1</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
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

/**
 * 중첩 테이블이 있는 HWPX 파일 생성
 */
async function createNestedTableHwpx(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Nested Table Test</hh:title></hh:docInfo>
</hh:head>`;

  // 셀 안에 중첩 테이블이 있는 구조
  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10"><hp:run charPrIDRef="0"><hp:t>Text before nested</hp:t></hp:run></hp:p>
          <hp:tbl id="200" rowCnt="2" colCnt="2">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0"><hp:subList><hp:p id="201"><hp:run><hp:t>Nested 0,0</hp:t></hp:run></hp:p></hp:subList></hp:tc>
              <hp:tc colAddr="1" rowAddr="0"><hp:subList><hp:p id="202"><hp:run><hp:t>Nested 0,1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
            </hp:tr>
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="1"><hp:subList><hp:p id="203"><hp:run><hp:t>Nested 1,0</hp:t></hp:run></hp:p></hp:subList></hp:tc>
              <hp:tc colAddr="1" rowAddr="1"><hp:subList><hp:p id="204"><hp:run><hp:t>Nested 1,1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="11"><hp:run charPrIDRef="0"><hp:t>R0C1 Simple</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="20"><hp:run charPrIDRef="0"><hp:t>R1C0 Simple</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList><hp:p id="21"><hp:run charPrIDRef="0"><hp:t>R1C1 Simple</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
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

/**
 * XML에서 태그 개수 세기
 */
function countTags(xml: string, tagName: string): { open: number; close: number } {
  const openPattern = new RegExp(`<(?:hp|hs|hc):${tagName}[\\s>]`, 'g');
  const closePattern = new RegExp(`<\\/(?:hp|hs|hc):${tagName}>`, 'g');
  return {
    open: (xml.match(openPattern) || []).length,
    close: (xml.match(closePattern) || []).length
  };
}

describe('XML 손상 방지 - 태그 균형 검증', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createBasicTestHwpx();
    testFilePath = path.join(__dirname, '..', 'test-corruption-basic.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('기본 테이블의 태그 균형이 맞아야 함', async () => {
    const buffer = fs.readFileSync(testFilePath);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(xml, 'tbl');
    const tr = countTags(xml, 'tr');
    const tc = countTags(xml, 'tc');

    expect(tbl.open).toBe(tbl.close);
    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(3); // 3 rows
    expect(tc.open).toBe(6); // 3 rows x 2 cols
  });

  it('셀 업데이트 후 태그 균형이 유지되어야 함', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 여러 셀 업데이트
    doc.updateTableCell(0, 0, 0, 0, 'Updated R0C0');
    doc.updateTableCell(0, 0, 1, 1, 'Updated R1C1');
    doc.updateTableCell(0, 0, 2, 0, 'Updated R2C0');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(savedXml, 'tbl');
    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tbl.open).toBe(tbl.close);
    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(3);
    expect(tc.open).toBe(6);
  });

  it('같은 행의 여러 셀 업데이트 시 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 같은 행(row 1)의 두 셀 모두 업데이트
    doc.updateTableCell(0, 0, 1, 0, 'Row1 Col0 Updated');
    doc.updateTableCell(0, 0, 1, 1, 'Row1 Col1 Updated');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(3);

    // 업데이트된 내용 확인
    expect(savedXml).toContain('Row1 Col0 Updated');
    expect(savedXml).toContain('Row1 Col1 Updated');
  });
});

describe('XML 손상 방지 - 중첩 테이블 보존', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createNestedTableHwpx();
    testFilePath = path.join(__dirname, '..', 'test-corruption-nested.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('중첩 테이블이 있는 HWPX 로드 시 태그 균형 확인', async () => {
    const buffer = fs.readFileSync(testFilePath);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(xml, 'tbl');
    const tr = countTags(xml, 'tr');
    const tc = countTags(xml, 'tc');

    // 외부 테이블 1개 + 중첩 테이블 1개 = 2개
    expect(tbl.open).toBe(2);
    expect(tbl.close).toBe(2);
    // 외부 테이블 2행 + 중첩 테이블 2행 = 4행
    expect(tr.open).toBe(4);
    expect(tr.close).toBe(4);
    // 외부 테이블 4셀 + 중첩 테이블 4셀 = 8셀
    expect(tc.open).toBe(8);
    expect(tc.close).toBe(8);
  });

  it('단순 텍스트 업데이트 시 중첩 테이블 보존', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 중첩 테이블이 있는 셀(0,0)을 단순 텍스트로 업데이트
    doc.updateTableCell(0, 0, 0, 0, 'Simple update');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(savedXml, 'tbl');
    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    // 중첩 테이블이 보존되어야 함
    expect(tbl.open).toBe(2);
    expect(tbl.close).toBe(2);
    expect(tr.open).toBe(4);
    expect(tr.close).toBe(4);

    // 중첩 테이블 내용도 보존
    expect(savedXml).toContain('Nested 0,0');
    expect(savedXml).toContain('id="200"'); // 중첩 테이블 ID
  });

  it('멀티라인 텍스트 업데이트 시 중첩 테이블 보존 (핵심 테스트)', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 중첩 테이블이 있는 셀(0,0)을 멀티라인 텍스트로 업데이트
    // 이 경우 updateTextInCellMultiline이 호출됨
    doc.updateTableCell(0, 0, 0, 0, 'Line 1\nLine 2\nLine 3');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(savedXml, 'tbl');
    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    // 중첩 테이블이 보존되어야 함 - 이것이 핵심 수정 사항!
    expect(tbl.open).toBe(2);
    expect(tbl.close).toBe(2);
    expect(tr.open).toBe(4);
    expect(tr.close).toBe(4);

    // 업데이트된 텍스트와 중첩 테이블 모두 존재
    expect(savedXml).toContain('Line 1');
    expect(savedXml).toContain('Line 2');
    expect(savedXml).toContain('Line 3');
    expect(savedXml).toContain('Nested 0,0'); // 중첩 테이블 내용 보존
    expect(savedXml).toContain('id="200"'); // 중첩 테이블 ID 보존
  });

  it('다른 셀 업데이트가 중첩 테이블에 영향 없음', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 중첩 테이블이 없는 다른 셀들 업데이트
    doc.updateTableCell(0, 0, 0, 1, 'Update R0C1');
    doc.updateTableCell(0, 0, 1, 0, 'Update R1C0');
    doc.updateTableCell(0, 0, 1, 1, 'Update R1C1');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(savedXml, 'tbl');
    const tr = countTags(savedXml, 'tr');

    // 중첩 테이블 구조 보존
    expect(tbl.open).toBe(2);
    expect(tbl.close).toBe(2);
    expect(tr.open).toBe(4);
    expect(tr.close).toBe(4);

    // 원래 중첩 테이블 내용 보존
    expect(savedXml).toContain('Text before nested');
    expect(savedXml).toContain('Nested 0,0');
  });
});

describe('XML 손상 방지 - 복합 업데이트 시나리오', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createBasicTestHwpx();
    testFilePath = path.join(__dirname, '..', 'test-corruption-complex.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('모든 셀 업데이트 후 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 모든 셀 업데이트 (3x2 = 6셀)
    for (let row = 0; row < 3; row++) {
      for (let col = 0; col < 2; col++) {
        doc.updateTableCell(0, 0, row, col, `Updated [${row},${col}]`);
      }
    }

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tbl = countTags(savedXml, 'tbl');
    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tbl.open).toBe(tbl.close);
    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(3);
    expect(tc.open).toBe(6);

    // 모든 업데이트된 내용 확인
    for (let row = 0; row < 3; row++) {
      for (let col = 0; col < 2; col++) {
        expect(savedXml).toContain(`Updated [${row},${col}]`);
      }
    }
  });

  it('긴 텍스트로 셀 업데이트 시 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 매우 긴 텍스트로 업데이트
    const longText = 'A'.repeat(10000);
    doc.updateTableCell(0, 0, 0, 0, longText);

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(3);
  });

  it('특수 문자가 포함된 텍스트로 업데이트 시 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // XML 특수 문자가 포함된 텍스트
    doc.updateTableCell(0, 0, 0, 0, 'Test <tag> & "quotes" \'apostrophe\'');
    doc.updateTableCell(0, 0, 1, 0, '한글 텍스트 테스트');
    doc.updateTableCell(0, 0, 2, 0, '🎉 이모지 테스트');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);

    // 특수 문자가 이스케이프되어 있어야 함
    expect(savedXml).toContain('&lt;tag&gt;');
    expect(savedXml).toContain('&amp;');
    expect(savedXml).toContain('한글 텍스트');
  });

  it('멀티라인 텍스트 업데이트 시 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 멀티라인 텍스트 (여러 paragraph 생성)
    doc.updateTableCell(0, 0, 0, 0, 'Line 1\nLine 2\nLine 3');
    doc.updateTableCell(0, 0, 1, 1, 'Another\nMultiline\nText');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');
    const p = countTags(savedXml, 'p');

    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
    expect(p.open).toBe(p.close);
    expect(tr.open).toBe(3);

    // 멀티라인 내용 확인
    expect(savedXml).toContain('Line 1');
    expect(savedXml).toContain('Line 2');
    expect(savedXml).toContain('Line 3');
  });
});

describe('XML 손상 방지 - 데이터 무결성', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createBasicTestHwpx();
    testFilePath = path.join(__dirname, '..', 'test-integrity.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('저장 후 다시 로드하면 동일한 데이터', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    doc.updateTableCell(0, 0, 0, 0, 'Test Data');
    doc.updateTableCell(0, 0, 1, 1, 'More Data');

    const savedBuffer = await doc.save();

    // 다시 로드
    const doc2 = await HwpxDocument.createFromBuffer('test2', testFilePath, savedBuffer);
    const table = doc2.getTable(0, 0);

    expect(table?.data[0][0].text).toBe('Test Data');
    expect(table?.data[1][1].text).toBe('More Data');
  });

  it('여러 번 저장-로드 사이클 후에도 데이터 유지', async () => {
    let buffer = fs.readFileSync(testFilePath);

    for (let i = 0; i < 5; i++) {
      const doc = await HwpxDocument.createFromBuffer(`test-${i}`, testFilePath, buffer);
      doc.updateTableCell(0, 0, 0, 0, `Iteration ${i}`);
      buffer = await doc.save();
    }

    // 최종 확인
    const finalDoc = await HwpxDocument.createFromBuffer('final', testFilePath, buffer);
    const table = finalDoc.getTable(0, 0);
    expect(table?.data[0][0].text).toBe('Iteration 4');

    // 태그 균형 확인
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file('Contents/section0.xml')?.async('string') || '';
    const tr = countTags(xml, 'tr');
    expect(tr.open).toBe(tr.close);
  });

  it('빈 문자열로 업데이트 시 태그 균형 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    doc.updateTableCell(0, 0, 0, 0, '');
    doc.updateTableCell(0, 0, 1, 1, '');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    const tc = countTags(savedXml, 'tc');

    expect(tr.open).toBe(tr.close);
    expect(tc.open).toBe(tc.close);
  });
});

describe('XML 손상 방지 - 경계 조건', () => {
  let testFilePath: string;

  beforeEach(async () => {
    const buffer = await createBasicTestHwpx();
    testFilePath = path.join(__dirname, '..', 'test-edge.hwpx');
    fs.writeFileSync(testFilePath, buffer);
  });

  it('존재하지 않는 셀 업데이트 시 에러 없음', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    // 범위를 벗어난 셀 업데이트 시도
    const result = doc.updateTableCell(0, 0, 10, 10, 'Invalid');
    expect(result).toBe(false);

    // 저장은 정상 동작해야 함
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tr = countTags(savedXml, 'tr');
    expect(tr.open).toBe(tr.close);
  });

  it('같은 셀 여러 번 업데이트 시 마지막 값 유지', async () => {
    const doc = await HwpxDocument.createFromBuffer('test', testFilePath, fs.readFileSync(testFilePath));

    doc.updateTableCell(0, 0, 0, 0, 'First');
    doc.updateTableCell(0, 0, 0, 0, 'Second');
    doc.updateTableCell(0, 0, 0, 0, 'Third');

    const savedBuffer = await doc.save();
    const doc2 = await HwpxDocument.createFromBuffer('test2', testFilePath, savedBuffer);
    const table = doc2.getTable(0, 0);

    expect(table?.data[0][0].text).toBe('Third');

    // 태그 균형 확인
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
    const tr = countTags(savedXml, 'tr');
    expect(tr.open).toBe(tr.close);
  });
});

// Cleanup after all tests
import { afterAll } from 'vitest';
afterAll(() => {
  const testFiles = [
    'test-corruption-basic.hwpx',
    'test-corruption-nested.hwpx',
    'test-corruption-complex.hwpx',
    'test-integrity.hwpx',
    'test-edge.hwpx'
  ];

  for (const file of testFiles) {
    const filePath = path.join(__dirname, '..', file);
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  }
});
