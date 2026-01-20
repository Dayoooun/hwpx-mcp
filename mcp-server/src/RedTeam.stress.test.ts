/**
 * RED TEAM STRESS TESTS
 *
 * 극한 상황에서 시스템이 버티는지 테스트
 * - 대량 테이블 (50개+)
 * - 중첩 테이블 + 내어쓰기 동시 적용
 * - 비순차적/무작위 테이블 수정
 * - 동일 셀 반복 수정
 * - 매우 긴 텍스트 + 많은 단락
 * - 복합 작업 (이미지 + 텍스트 + 내어쓰기)
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

// 테스트용 1x1 PNG (base64)
const testImageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

/**
 * 대량 테이블이 포함된 복잡한 문서 생성
 */
async function createComplexDocument(tableCount: number, rowsPerTable: number, colsPerTable: number): Promise<Buffer> {
  const zip = new JSZip();

  zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

  zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/common">
  <hh:refList>
    <hh:charShapeList itemCnt="1">
      <hh:charShape id="0" height="1000" baseSize="1000" ratio="100" spacing="0" relSize="100"/>
    </hh:charShapeList>
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
      </hh:paraPr>
    </hh:paraProperties>
    <hh:paraShapeList itemCnt="1">
      <hh:paraShape id="0"/>
    </hh:paraShapeList>
  </hh:refList>
</hh:head>`);

  // 대량 테이블 생성
  let sectionContent = '';
  let paraId = 1;
  let tableId = 100;

  for (let t = 0; t < tableCount; t++) {
    // 테이블 앞 단락
    sectionContent += `<hp:p id="${paraId++}" paraPrIDRef="0" styleIDRef="0">
      <hp:run charPrIDRef="0"><hp:t>테이블 ${t} 제목</hp:t></hp:run>
    </hp:p>\n`;

    // 테이블 시작
    sectionContent += `<hp:tbl id="${tableId++}" rowCnt="${rowsPerTable}" colCnt="${colsPerTable}">\n`;

    for (let r = 0; r < rowsPerTable; r++) {
      sectionContent += `  <hp:tr>\n`;
      for (let c = 0; c < colsPerTable; c++) {
        sectionContent += `    <hp:tc colAddr="${c}" rowAddr="${r}" colSpan="1" rowSpan="1">
      <hp:subList>
        <hp:p id="${paraId++}" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>T${t}R${r}C${c}</hp:t></hp:run>
        </hp:p>
      </hp:subList>
    </hp:tc>\n`;
      }
      sectionContent += `  </hp:tr>\n`;
    }

    sectionContent += `</hp:tbl>\n`;
  }

  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
${sectionContent}
</hs:sec>`);

  zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="urn:hancom:hwp:file">
  <hpf:manifest>
    <hpf:item href="Contents/section0.xml" type="application/xml"/>
    <hpf:item href="Contents/header.xml" type="application/xml"/>
  </hpf:manifest>
</hpf:package>`);

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`);

  return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * XML 태그 균형 검사
 */
function checkTagBalance(xml: string): { balanced: boolean; details: Record<string, { open: number; close: number }> } {
  const tags = ['hp:tbl', 'hp:tr', 'hp:tc', 'hp:p', 'hp:run', 'hp:subList'];
  const details: Record<string, { open: number; close: number }> = {};
  let balanced = true;

  for (const tag of tags) {
    const openRegex = new RegExp(`<${tag}\\b`, 'g');
    const closeRegex = new RegExp(`</${tag}>`, 'g');
    const open = (xml.match(openRegex) || []).length;
    const close = (xml.match(closeRegex) || []).length;
    details[tag] = { open, close };
    if (open !== close) balanced = false;
  }

  return { balanced, details };
}

describe('RED TEAM - Extreme Stress Tests', () => {

  describe('대량 테이블 스트레스 테스트', () => {

    it('50개 테이블에 무작위 순서로 내어쓰기 적용', async () => {
      const buffer = await createComplexDocument(50, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 무작위 순서로 테이블 선택 (0-49에서 20개)
      const randomTables = [47, 3, 29, 8, 42, 15, 31, 0, 49, 22, 11, 38, 5, 44, 17, 26, 33, 9, 41, 2];

      for (const tableIdx of randomTables) {
        const text = `○ 항목1-${tableIdx}\n○ 항목2-${tableIdx}\n○ 항목3-${tableIdx}`;
        doc.updateTableCell(0, tableIdx, 0, 0, text);

        // 각 단락에 내어쓰기
        for (let p = 0; p < 3; p++) {
          doc.setTableCellHangingIndent(0, tableIdx, 0, 0, p, 12);
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      expect(sectionXml).toBeDefined();
      const { balanced, details } = checkTagBalance(sectionXml!);

      console.log('50개 테이블 무작위 수정 결과:', details);
      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(50);
    });

    it('100개 테이블 중 홀수 인덱스만 수정', async () => {
      const buffer = await createComplexDocument(100, 1, 1);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 홀수 인덱스만 (1, 3, 5, ... 99) = 50개
      for (let i = 1; i < 100; i += 2) {
        doc.updateTableCell(0, i, 0, 0, `■ 홀수테이블${i}`);
        doc.setTableCellHangingIndent(0, i, 0, 0, 0, 10);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('100개 테이블 홀수만 수정 결과:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(100);
    });
  });

  describe('동일 셀 반복 수정 스트레스', () => {

    it('같은 셀을 10번 연속 수정 후 저장', async () => {
      const buffer = await createComplexDocument(5, 3, 3);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 같은 셀을 10번 수정 (마지막 값만 유지되어야 함)
      for (let i = 0; i < 10; i++) {
        const text = `수정${i}: ○ 항목A\n○ 항목B`;
        doc.updateTableCell(0, 2, 1, 1, text);

        doc.setTableCellHangingIndent(0, 2, 1, 1, 0, 10 + i);
        doc.setTableCellHangingIndent(0, 2, 1, 1, 1, 10 + i);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);

      // 마지막 수정 내용 확인
      expect(sectionXml).toContain('수정9');
    });

    it('여러 테이블의 같은 위치 셀을 동시에 수정', async () => {
      const buffer = await createComplexDocument(30, 4, 4);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 모든 테이블의 (2,2) 셀 수정
      for (let t = 0; t < 30; t++) {
        doc.updateTableCell(0, t, 2, 2, `○ 테이블${t} 중앙셀\n○ 두번째줄`);
        doc.setTableCellHangingIndent(0, t, 2, 2, 0, 12);
        doc.setTableCellHangingIndent(0, t, 2, 2, 1, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('30개 테이블 동일위치 수정 결과:', details);

      expect(balanced).toBe(true);
      expect(sectionXml).toContain('테이블0 중앙셀');
      expect(sectionXml).toContain('테이블29 중앙셀');
    });
  });

  describe('매우 긴 텍스트 + 많은 단락', () => {

    it('한 셀에 50줄 텍스트 + 모든 줄에 내어쓰기', async () => {
      const buffer = await createComplexDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 50줄 텍스트 생성
      const lines: string[] = [];
      for (let i = 0; i < 50; i++) {
        lines.push(`○ ${i + 1}번째 항목입니다. 이것은 매우 긴 텍스트 라인입니다.`);
      }
      const longText = lines.join('\n');

      doc.updateTableCell(0, 1, 0, 0, longText);

      // 모든 50개 단락에 내어쓰기 적용
      for (let p = 0; p < 50; p++) {
        doc.setTableCellHangingIndent(0, 1, 0, 0, p, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);

      // 50개 단락 존재 확인 (테이블 1의 셀 0,0 내부)
      const table1Match = sectionXml?.match(/<hp:tbl[^>]*id="101"[^>]*>([\s\S]*?)<\/hp:tbl>/);
      expect(table1Match).toBeTruthy();

      const paragraphCount = (table1Match![1].match(/<hp:p\b/g) || []).length;
      // 2x2 테이블 = 4셀, 첫 셀에 50개 + 나머지 3셀에 1개씩 = 53개
      expect(paragraphCount).toBeGreaterThanOrEqual(50);
    });

    it('여러 테이블에 각각 다른 길이의 텍스트', async () => {
      const buffer = await createComplexDocument(10, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 각 테이블에 다른 길이의 텍스트
      for (let t = 0; t < 10; t++) {
        const lineCount = (t + 1) * 5; // 5, 10, 15, ... 50줄
        const lines: string[] = [];
        for (let i = 0; i < lineCount; i++) {
          lines.push(`○ T${t}L${i}`);
        }

        doc.updateTableCell(0, t, 0, 0, lines.join('\n'));

        for (let p = 0; p < lineCount; p++) {
          doc.setTableCellHangingIndent(0, t, 0, 0, p, 10 + t);
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('10개 테이블 다양한 길이 결과:', details);

      expect(balanced).toBe(true);
    });
  });

  describe('복합 작업 동시 실행', () => {

    it('텍스트 수정 + 내어쓰기 + 이미지 삽입 동시에', async () => {
      const buffer = await createComplexDocument(10, 3, 3);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      for (let t = 0; t < 10; t++) {
        // 텍스트 수정
        doc.updateTableCell(0, t, 0, 0, `○ 테이블${t} 텍스트\n○ 두번째줄`);
        doc.updateTableCell(0, t, 1, 1, `■ 중앙셀 내용`);

        // 내어쓰기
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, 12);
        doc.setTableCellHangingIndent(0, t, 0, 0, 1, 12);
        doc.setTableCellHangingIndent(0, t, 1, 1, 0, 10);

        // 이미지 삽입 (다른 셀에)
        doc.insertImageInCell(0, t, 2, 2, {
          data: testImageBase64,
          mimeType: 'image/png',
          width: 50,
          height: 50,
        });
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('복합 작업 결과:', details);

      expect(balanced).toBe(true);

      // 이미지 10개 확인
      const picCount = (sectionXml?.match(/<hp:pic\b/g) || []).length;
      expect(picCount).toBe(10);

      // BinData에 이미지 10개 (폴더 엔트리 제외, 실제 파일만 카운트)
      const binFiles = Object.keys(savedZip.files).filter(f => f.startsWith('BinData/') && !f.endsWith('/'));
      expect(binFiles.length).toBe(10);
    });

    it('역순으로 테이블 수정 (49 -> 0)', async () => {
      const buffer = await createComplexDocument(50, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 역순으로 수정
      for (let t = 49; t >= 0; t--) {
        doc.updateTableCell(0, t, 0, 0, `○ 역순${t}`);
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
      expect(sectionXml).toContain('역순0');
      expect(sectionXml).toContain('역순49');
    });
  });

  describe('엣지 케이스', () => {

    it('첫 번째와 마지막 테이블만 수정', async () => {
      const buffer = await createComplexDocument(100, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 첫 번째 테이블 (0)
      doc.updateTableCell(0, 0, 0, 0, '○ 첫번째테이블\n○ 두번째줄');
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10);
      doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 10);

      // 마지막 테이블 (99)
      doc.updateTableCell(0, 99, 1, 1, '○ 마지막테이블\n○ 두번째줄\n○ 세번째줄');
      doc.setTableCellHangingIndent(0, 99, 1, 1, 0, 14);
      doc.setTableCellHangingIndent(0, 99, 1, 1, 1, 14);
      doc.setTableCellHangingIndent(0, 99, 1, 1, 2, 14);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('첫/마지막 테이블만 수정 결과:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(100);
      expect(sectionXml).toContain('첫번째테이블');
      expect(sectionXml).toContain('마지막테이블');
    });

    it('모든 테이블의 모든 셀에 내어쓰기 적용', async () => {
      const buffer = await createComplexDocument(20, 3, 3);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 20 테이블 x 3행 x 3열 = 180개 셀
      for (let t = 0; t < 20; t++) {
        for (let r = 0; r < 3; r++) {
          for (let c = 0; c < 3; c++) {
            doc.updateTableCell(0, t, r, c, `○ T${t}R${r}C${c}\n○ 추가줄`);
            doc.setTableCellHangingIndent(0, t, r, c, 0, 12);
            doc.setTableCellHangingIndent(0, t, r, c, 1, 12);
          }
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('모든 셀 수정 결과:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(20);
      // 20 테이블 x 3행 x 3열 = 180 셀
      expect(details['hp:tc'].open).toBe(180);
    });

    it('다양한 내어쓰기 크기 (1pt ~ 50pt)', async () => {
      const buffer = await createComplexDocument(50, 1, 1);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 각 테이블에 다른 내어쓰기 크기
      for (let t = 0; t < 50; t++) {
        const indentSize = t + 1; // 1pt ~ 50pt
        doc.updateTableCell(0, t, 0, 0, `○ 들여쓰기${indentSize}pt`);
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, indentSize);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);

      // header.xml에 50개의 서로 다른 paraPr 생성 확인
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');
      const paraPrCount = (headerXml?.match(/<hh:paraPr\b/g) || []).length;

      console.log(`생성된 paraPr 개수: ${paraPrCount}`);
      // 원본 1개 + 50개 새로 생성 = 51개
      expect(paraPrCount).toBe(51);

      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });
  });

  describe('연속 저장 테스트', () => {

    it('수정 -> 저장 -> 다시 열기 -> 수정 -> 저장 반복', async () => {
      let buffer = await createComplexDocument(10, 2, 2);

      for (let iteration = 0; iteration < 5; iteration++) {
        const doc = await HwpxDocument.createFromBuffer('test', `iter${iteration}.hwpx`, buffer);

        // 매 반복마다 다른 테이블 수정
        const targetTable = iteration * 2;
        doc.updateTableCell(0, targetTable, 0, 0, `○ 반복${iteration}\n○ 추가`);
        doc.setTableCellHangingIndent(0, targetTable, 0, 0, 0, 10 + iteration);
        doc.setTableCellHangingIndent(0, targetTable, 0, 0, 1, 10 + iteration);

        buffer = await doc.save();

        // 저장된 파일 검증
        const savedZip = await JSZip.loadAsync(buffer);
        const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');
        const { balanced } = checkTagBalance(sectionXml!);

        expect(balanced).toBe(true);
        expect(sectionXml).toContain(`반복${iteration}`);
      }

      // 최종 파일에 모든 수정 내용 존재 확인
      const finalZip = await JSZip.loadAsync(buffer);
      const finalXml = await finalZip.file('Contents/section0.xml')?.async('string');

      for (let i = 0; i < 5; i++) {
        expect(finalXml).toContain(`반복${i}`);
      }
    });
  });

  describe('특수 문자 및 긴 마커', () => {

    it('다양한 마커 타입 혼합 사용', async () => {
      const buffer = await createComplexDocument(5, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      const markers = [
        '○ 동그라미 마커',
        '● 검정 동그라미',
        '■ 사각형 마커',
        '□ 빈 사각형',
        '◆ 다이아몬드',
        '▶ 삼각형',
        '※ 특수기호',
        '1. 숫자마커',
        '가. 한글마커',
        '(1) 괄호숫자',
        'A. 영문마커',
        '① 원숫자',
      ];

      // 각 테이블에 다양한 마커 조합
      for (let t = 0; t < 5; t++) {
        const startIdx = t * 2;
        const text = markers.slice(startIdx, startIdx + 4).join('\n');
        doc.updateTableCell(0, t, 0, 0, text);

        for (let p = 0; p < 4 && startIdx + p < markers.length; p++) {
          doc.setTableCellHangingIndent(0, t, 0, 0, p, 12);
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);

      // 모든 마커가 포함되어 있는지 확인
      expect(sectionXml).toContain('동그라미');
      expect(sectionXml).toContain('사각형');
      expect(sectionXml).toContain('숫자마커');
    });

    it('매우 긴 텍스트 라인 (500자 이상)', async () => {
      const buffer = await createComplexDocument(3, 1, 1);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 500자 이상의 긴 텍스트
      const longLine = '○ ' + '가나다라마바사아자차카타파하'.repeat(40); // 약 560자
      const text = `${longLine}\n${longLine}\n${longLine}`;

      doc.updateTableCell(0, 1, 0, 0, text);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 12);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 2, 12);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });
  });
});
