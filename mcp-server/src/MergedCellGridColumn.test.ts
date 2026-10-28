/**
 * 한/글이 저장한 표는 병합으로 덮인 칸의 <hp:tc> 를 빼고 저장한다.
 * 실측 결함: 예비창업패키지 양식 「명칭 | (칸) | 범주(2칸 병합) | (칸)」 행에서
 * 마지막 칸(배열 3번, 격자 4열)을 「범주」가 덮는다고 잘못 보고 쓰기를 거부했다.
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

const tc = (col: number, span: number, text: string) =>
  `<hp:tc><hp:subList><hp:p id="p${col}" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p></hp:subList>` +
  `<hp:cellAddr colAddr="${col}" rowAddr="0"/><hp:cellSpan colSpan="${span}" rowSpan="1"/><hp:cellSz width="1000" height="1000"/></hp:tc>`;

async function hancomRow(): Promise<HwpxDocument> {
  const zip = new JSZip();
  zip.file('mimetype', 'application/hwp+zip');
  zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');
  zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?><hpf:package xmlns:hpf="urn:hancom:hwp:file"><hpf:manifest><hpf:item href="Contents/section0.xml" type="application/xml"/><hpf:item href="Contents/header.xml" type="application/xml"/></hpf:manifest></hpf:package>`);
  zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?><hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"><hh:refList><hh:charProperties itemCnt="1"><hh:charPr id="0" height="1000"/></hh:charProperties><hh:paraProperties itemCnt="1"><hh:paraPr id="0" tabPrIDRef="0"><hh:align horizontal="LEFT" vertical="BASELINE"/></hh:paraPr></hh:paraProperties></hh:refList></hh:head>`);
  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="b" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:tbl id="9" rowCnt="1" colCnt="5"><hp:tr>${tc(0, 1, '명칭')}${tc(1, 1, '')}${tc(2, 2, '범주')}${tc(4, 1, '')}</hp:tr></hp:tbl></hp:run></hp:p>
</hs:sec>`);
  return HwpxDocument.createFromBuffer('m', 'm.hwpx', await zip.generateAsync({ type: 'nodebuffer' }));
}

describe('병합 칸 판정은 격자 열(colAddr)로 한다', () => {
  it('2칸 병합 뒤 칸(배열 3, 격자 4)에 쓸 수 있고 저장 후에도 남는다', async () => {
    const doc = await hancomRow();
    expect(doc.updateTableCell(0, 0, 0, 3, '연구 데이터 처리 소프트웨어')).toBe(true);
    const z = await JSZip.loadAsync(await doc.save());
    const xml = await z.file('Contents/section0.xml')!.async('string');
    const last = xml.split('<hp:tc>')[4];
    expect(last).toContain('연구 데이터 처리 소프트웨어');
    expect(last).toContain('colAddr="4"');
  });

  it('병합 칸 자신(배열 2)에는 그대로 쓸 수 있다', async () => {
    const doc = await hancomRow();
    expect(doc.updateTableCell(0, 0, 0, 2, '범주')).toBe(true);
  });
});
