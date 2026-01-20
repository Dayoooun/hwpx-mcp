/**
 * RED TEAM - EXTREME EDGE CASE TESTS
 *
 * 더 극한의 상황 테스트:
 * - 중첩 테이블 + 내어쓰기 동시 처리
 * - 200개+ 테이블
 * - 메모리 스트레스 (대용량 텍스트)
 * - 경계값 테스트
 * - 손상 시나리오 복구
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

const testImageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

/**
 * 복잡한 문서 생성 헬퍼
 */
async function createDocument(tableCount: number, rows: number, cols: number): Promise<Buffer> {
  const zip = new JSZip();

  zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

  zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/common">
  <hh:refList>
    <hh:charShapeList itemCnt="1"><hh:charShape id="0" height="1000" baseSize="1000"/></hh:charShapeList>
    <hh:paraProperties itemCnt="1">
      <hh:paraPr id="0" tabPrIDRef="0"><hh:align horizontal="JUSTIFY" vertical="BASELINE"/><hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/></hh:margin></hh:paraPr>
    </hh:paraProperties>
    <hh:paraShapeList itemCnt="1"><hh:paraShape id="0"/></hh:paraShapeList>
  </hh:refList>
</hh:head>`);

  let sectionContent = '';
  let paraId = 1;
  let tableId = 100;

  for (let t = 0; t < tableCount; t++) {
    sectionContent += `<hp:p id="${paraId++}" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>Table ${t}</hp:t></hp:run></hp:p>\n`;
    sectionContent += `<hp:tbl id="${tableId++}" rowCnt="${rows}" colCnt="${cols}">\n`;
    for (let r = 0; r < rows; r++) {
      sectionContent += `  <hp:tr>\n`;
      for (let c = 0; c < cols; c++) {
        sectionContent += `    <hp:tc colAddr="${c}" rowAddr="${r}" colSpan="1" rowSpan="1"><hp:subList><hp:p id="${paraId++}" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>T${t}R${r}C${c}</hp:t></hp:run></hp:p></hp:subList></hp:tc>\n`;
      }
      sectionContent += `  </hp:tr>\n`;
    }
    sectionContent += `</hp:tbl>\n`;
  }

  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
${sectionContent}
</hs:sec>`);

  zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?><hpf:package xmlns:hpf="urn:hancom:hwp:file"><hpf:manifest><hpf:item href="Contents/section0.xml" type="application/xml"/><hpf:item href="Contents/header.xml" type="application/xml"/></hpf:manifest></hpf:package>`);
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`);

  return zip.generateAsync({ type: 'nodebuffer' });
}

function checkTagBalance(xml: string): { balanced: boolean; details: Record<string, { open: number; close: number }> } {
  const tags = ['hp:tbl', 'hp:tr', 'hp:tc', 'hp:p', 'hp:run', 'hp:subList'];
  const details: Record<string, { open: number; close: number }> = {};
  let balanced = true;

  for (const tag of tags) {
    const open = (xml.match(new RegExp(`<${tag}\\b`, 'g')) || []).length;
    const close = (xml.match(new RegExp(`</${tag}>`, 'g')) || []).length;
    details[tag] = { open, close };
    if (open !== close) balanced = false;
  }

  return { balanced, details };
}

describe('RED TEAM - Extreme Edge Cases', () => {

  describe('대용량 문서 (200+ 테이블)', () => {

    it('200개 테이블 생성 후 랜덤 50개 수정', async () => {
      const buffer = await createDocument(200, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'large.hwpx', buffer);

      // 랜덤하게 50개 테이블 선택
      const randomIndices = Array.from({ length: 50 }, () => Math.floor(Math.random() * 200));
      const uniqueIndices = [...new Set(randomIndices)];

      for (const idx of uniqueIndices) {
        doc.updateTableCell(0, idx, 0, 0, `○ 수정됨-${idx}\n○ 두번째줄`);
        doc.setTableCellHangingIndent(0, idx, 0, 0, 0, 12);
        doc.setTableCellHangingIndent(0, idx, 0, 0, 1, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('200개 테이블 랜덤 수정:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(200);
    }, 30000); // 30초 타임아웃

    it('150개 테이블 전체에 내어쓰기 적용', async () => {
      const buffer = await createDocument(150, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'large.hwpx', buffer);

      // 모든 테이블의 첫 셀에 내어쓰기
      for (let t = 0; t < 150; t++) {
        doc.updateTableCell(0, t, 0, 0, `○ T${t}`);
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    }, 30000);
  });

  describe('중첩 테이블 + 내어쓰기 복합', () => {

    it('테이블에 중첩 테이블 삽입 후 내어쓰기 적용', async () => {
      const buffer = await createDocument(5, 3, 3);
      const doc = await HwpxDocument.createFromBuffer('test', 'nested.hwpx', buffer);

      // 테이블 2에 중첩 테이블 삽입
      doc.insertNestedTable(0, 2, 1, 1, 2, 2);

      // 다른 테이블들에 내어쓰기 적용
      doc.updateTableCell(0, 0, 0, 0, '○ 첫테이블\n○ 두번째줄');
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 12);
      doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 12);

      doc.updateTableCell(0, 4, 2, 2, '■ 마지막테이블\n■ 항목2');
      doc.setTableCellHangingIndent(0, 4, 2, 2, 0, 10);
      doc.setTableCellHangingIndent(0, 4, 2, 2, 1, 10);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('중첩 테이블 + 내어쓰기:', details);

      expect(balanced).toBe(true);
      // 5개 기본 테이블 + 1개 중첩 테이블 = 6개
      expect(details['hp:tbl'].open).toBe(6);
    });

    it('여러 테이블에 중첩 테이블 + 이미지 + 내어쓰기', async () => {
      const buffer = await createDocument(10, 3, 3);
      const doc = await HwpxDocument.createFromBuffer('test', 'complex.hwpx', buffer);

      // 테이블 3, 5, 7에 중첩 테이블
      doc.insertNestedTable(0, 3, 1, 1, 2, 2);
      doc.insertNestedTable(0, 5, 1, 1, 2, 2);
      doc.insertNestedTable(0, 7, 1, 1, 2, 2);

      // 테이블 0, 2, 4, 6, 8에 내어쓰기 + 이미지
      for (const t of [0, 2, 4, 6, 8]) {
        doc.updateTableCell(0, t, 0, 0, `○ T${t} 항목\n○ 두번째`);
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, 12);
        doc.setTableCellHangingIndent(0, t, 0, 0, 1, 12);

        doc.insertImageInCell(0, t, 2, 2, {
          data: testImageBase64,
          mimeType: 'image/png',
          width: 30,
          height: 30,
        });
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('중첩+이미지+내어쓰기:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(13); // 10 + 3 중첩

      const picCount = (sectionXml?.match(/<hp:pic\b/g) || []).length;
      expect(picCount).toBe(5);
    });
  });

  describe('경계값 테스트', () => {

    it('빈 텍스트로 업데이트', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 빈 텍스트
      doc.updateTableCell(0, 1, 0, 0, '');

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });

    it('단일 문자 텍스트', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      doc.updateTableCell(0, 1, 0, 0, '가');
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
      expect(sectionXml).toContain('가');
    });

    it('0pt 내어쓰기 시도 (무시되어야 함)', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      doc.updateTableCell(0, 1, 0, 0, '○ 테스트');
      const result = doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 0);

      expect(result).toBe(false); // 0pt는 거부되어야 함
    });

    it('음수 내어쓰기 시도 (무시되어야 함)', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      doc.updateTableCell(0, 1, 0, 0, '○ 테스트');
      const result = doc.setTableCellHangingIndent(0, 1, 0, 0, 0, -10);

      expect(result).toBe(false);
    });

    it('존재하지 않는 테이블 인덱스', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 테이블 99는 존재하지 않음
      const result = doc.setTableCellHangingIndent(0, 99, 0, 0, 0, 12);
      expect(result).toBe(false);
    });

    it('존재하지 않는 셀 좌표', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 행 99, 열 99는 존재하지 않음
      const result = doc.setTableCellHangingIndent(0, 1, 99, 99, 0, 12);
      expect(result).toBe(false);
    });
  });

  describe('동시성 시뮬레이션', () => {

    it('같은 테이블의 여러 셀을 번갈아 수정', async () => {
      const buffer = await createDocument(3, 4, 4);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 체스판 패턴으로 수정
      for (let r = 0; r < 4; r++) {
        for (let c = 0; c < 4; c++) {
          if ((r + c) % 2 === 0) {
            doc.updateTableCell(0, 1, r, c, `○ R${r}C${c}\n○ 줄2`);
            doc.setTableCellHangingIndent(0, 1, r, c, 0, 12);
            doc.setTableCellHangingIndent(0, 1, r, c, 1, 12);
          }
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });

    it('여러 테이블을 지그재그로 수정', async () => {
      const buffer = await createDocument(20, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 짝수 -> 홀수 -> 짝수 순서로 수정
      const order = [0, 19, 2, 17, 4, 15, 6, 13, 8, 11, 10, 9, 12, 7, 14, 5, 16, 3, 18, 1];

      for (const t of order) {
        doc.updateTableCell(0, t, 0, 0, `○ 순서-${order.indexOf(t)}`);
        doc.setTableCellHangingIndent(0, t, 0, 0, 0, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('지그재그 수정:', details);

      expect(balanced).toBe(true);
      expect(details['hp:tbl'].open).toBe(20);
    });
  });

  describe('메모리/성능 스트레스', () => {

    it('100줄 텍스트 x 10개 테이블', async () => {
      const buffer = await createDocument(10, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 각 테이블에 100줄 텍스트
      for (let t = 0; t < 10; t++) {
        const lines = Array.from({ length: 100 }, (_, i) => `○ T${t}L${i} 긴텍스트입니다`);
        doc.updateTableCell(0, t, 0, 0, lines.join('\n'));

        // 모든 100줄에 내어쓰기
        for (let p = 0; p < 100; p++) {
          doc.setTableCellHangingIndent(0, t, 0, 0, p, 12);
        }
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced, details } = checkTagBalance(sectionXml!);
      console.log('100줄x10테이블:', details);

      expect(balanced).toBe(true);
      // 10개 테이블 x (첫 셀에 100줄 + 나머지 3셀에 1줄) = 10 x 103 + 10(테이블 앞 단락)
      expect(details['hp:p'].open).toBeGreaterThan(1000);
    }, 60000);

    it('1000자 텍스트 라인 x 20줄', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'stress.hwpx', buffer);

      // 각 라인이 1000자
      const longSegment = '가나다라마바사아자차카타파하'.repeat(70); // ~980자
      const lines = Array.from({ length: 20 }, (_, i) => `○ ${i}: ${longSegment}`);

      doc.updateTableCell(0, 1, 0, 0, lines.join('\n'));

      for (let p = 0; p < 20; p++) {
        doc.setTableCellHangingIndent(0, 1, 0, 0, p, 12);
      }

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });
  });

  describe('특수 시나리오', () => {

    it('같은 셀에 텍스트 수정 -> 이미지 삽입 -> 내어쓰기 순서', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 순서: 텍스트 -> 이미지 -> 내어쓰기
      doc.updateTableCell(0, 1, 0, 0, '○ 먼저 텍스트\n○ 두번째');
      doc.insertImageInCell(0, 1, 0, 0, {
        data: testImageBase64,
        mimeType: 'image/png',
        width: 50,
        height: 50,
      });
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 12);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });

    it('이미지 먼저 -> 텍스트 수정 -> 내어쓰기 순서', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 순서: 이미지 -> 텍스트 -> 내어쓰기
      doc.insertImageInCell(0, 1, 1, 1, {
        data: testImageBase64,
        mimeType: 'image/png',
        width: 50,
        height: 50,
      });
      doc.updateTableCell(0, 1, 0, 0, '○ 나중 텍스트\n○ 두번째');
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 12);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);
    });

    it('내어쓰기 적용 후 같은 셀 텍스트 덮어쓰기', async () => {
      const buffer = await createDocument(3, 2, 2);
      const doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', buffer);

      // 첫 번째 수정 + 내어쓰기
      doc.updateTableCell(0, 1, 0, 0, '○ 첫번째\n○ 두번째');
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 12);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 12);

      // 다시 덮어쓰기 + 다른 내어쓰기
      doc.updateTableCell(0, 1, 0, 0, '■ 덮어쓴텍스트\n■ 새두번째\n■ 세번째');
      doc.setTableCellHangingIndent(0, 1, 0, 0, 0, 14);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 1, 14);
      doc.setTableCellHangingIndent(0, 1, 0, 0, 2, 14);

      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      const { balanced } = checkTagBalance(sectionXml!);
      expect(balanced).toBe(true);

      // 덮어쓴 내용만 있어야 함
      expect(sectionXml).toContain('덮어쓴텍스트');
      expect(sectionXml).not.toContain('첫번째');
    });
  });
});
