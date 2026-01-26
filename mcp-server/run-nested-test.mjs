/**
 * 중첩 테이블 삭제 테스트
 */
import { HwpxDocument } from './dist/HwpxDocument.js';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

const testOutputDir = path.join(process.cwd(), 'test-output');
if (!fs.existsSync(testOutputDir)) {
  fs.mkdirSync(testOutputDir, { recursive: true });
}

// 중첩 테이블이 포함된 문서 생성
async function createNestedTableDocument() {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Nested Table Test</hh:title></hh:docInfo>
</hh:head>`;

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="p1"><hp:run><hp:t>문서 시작</hp:t></hp:run></hp:p>

  <!-- 외부 테이블 1 (중첩 테이블 2개 포함) -->
  <hp:tbl id="tbl_outer1" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0">
        <hp:subList>
          <hp:p id="outer1_cell00"><hp:run><hp:t>외부1-셀00</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 1 -->
          <hp:tbl id="tbl_inner1" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner1"><hp:run><hp:t>내부테이블1</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0">
        <hp:subList>
          <hp:p id="outer1_cell01"><hp:run><hp:t>외부1-셀01</hp:t></hp:run></hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1">
        <hp:subList>
          <hp:p id="outer1_cell10"><hp:run><hp:t>외부1-셀10</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 2 -->
          <hp:tbl id="tbl_inner2" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner2"><hp:run><hp:t>내부테이블2</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1">
        <hp:subList><hp:p id="outer1_cell11"><hp:run><hp:t>외부1-셀11</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p2"><hp:run><hp:t>중간 문단</hp:t></hp:run></hp:p>

  <!-- 단순 테이블 (중첩 없음) -->
  <hp:tbl id="tbl_simple" rowCnt="1" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0"><hp:subList><hp:p id="simple_cell0"><hp:run><hp:t>단순셀0</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="1" rowAddr="0"><hp:subList><hp:p id="simple_cell1"><hp:run><hp:t>단순셀1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p3"><hp:run><hp:t>또 다른 중간 문단</hp:t></hp:run></hp:p>

  <!-- 외부 테이블 2 (중첩 테이블 1개 포함) -->
  <hp:tbl id="tbl_outer2" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0">
        <hp:subList>
          <hp:p id="outer2_cell"><hp:run><hp:t>외부2-셀</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 3 -->
          <hp:tbl id="tbl_inner3" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner3"><hp:run><hp:t>내부테이블3</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p4"><hp:run><hp:t>문서 끝</hp:t></hp:run></hp:p>
</hs:sec>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', sectionXml);
  zip.file('Contents/content.hpf', '<?xml version="1.0"?><pkg:package xmlns:pkg="http://www.hancom.co.kr/hwpml/2011/package"><pkg:manifest><pkg:item id="section0" href="section0.xml"/></pkg:manifest></pkg:package>');
  zip.file('version.xml', '<?xml version="1.0"?><hwpml version="1.0"/>');
  zip.file('mimetype', 'application/hwp+zip');

  return await zip.generateAsync({ type: 'nodebuffer' });
}

// 태그 균형 검사
function checkTagBalance(xml, tagName) {
  const openRegex = new RegExp(`<hp:${tagName}\\b`, 'g');
  const closeRegex = new RegExp(`</hp:${tagName}>`, 'g');
  const openCount = (xml.match(openRegex) || []).length;
  const closeCount = (xml.match(closeRegex) || []).length;
  return { open: openCount, close: closeCount, balanced: openCount === closeCount };
}

async function runTests() {
  let passed = 0;
  let failed = 0;

  console.log('\n========================================');
  console.log('  중첩 테이블 삭제 E2E 테스트');
  console.log('========================================\n');

  // 테스트 1: 중첩 테이블 포함 테이블 삭제 후 태그 균형
  console.log('테스트 1: 중첩 테이블 포함 테이블 삭제');
  console.log('----------------------------------------');
  try {
    const buffer = await createNestedTableDocument();
    const testPath = path.join(testOutputDir, 'nested-test-1.hwpx');
    fs.writeFileSync(testPath, buffer);

    let doc = await HwpxDocument.createFromBuffer('test1', testPath, buffer);

    // 원본 확인
    const origZip = await JSZip.loadAsync(buffer);
    const origXml = await origZip.file('Contents/section0.xml')?.async('string') || '';
    const origBalance = checkTagBalance(origXml, 'tbl');
    console.log(`  원본: <hp:tbl> ${origBalance.open}개, </hp:tbl> ${origBalance.close}개`);

    if (origBalance.open !== 6 || !origBalance.balanced) {
      throw new Error('원본 태그 불균형');
    }

    // 외부 테이블 1 삭제 (중첩 2개 포함)
    doc.deleteTable(0, 0);
    console.log('  외부 테이블 1 삭제 (중첩 테이블 2개 포함)');

    // 저장
    const savedBuffer = await doc.save();
    fs.writeFileSync(testPath, savedBuffer);

    // 검증
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
    const savedBalance = checkTagBalance(savedXml, 'tbl');

    console.log(`  저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);

    if (!savedBalance.balanced) {
      console.log('  ❌ FAIL: 태그 불균형!');
      failed++;
    } else if (savedBalance.open !== 3) {
      console.log(`  ❌ FAIL: 예상 3개, 실제 ${savedBalance.open}개`);
      failed++;
    } else {
      console.log('  ✅ PASS');
      passed++;
    }

    fs.unlinkSync(testPath);
  } catch (e) {
    console.log('  ❌ FAIL:', e.message);
    failed++;
  }

  // 테스트 2: 연속 삭제 (앞에서부터)
  console.log('\n테스트 2: 연속 삭제 (index 0 반복)');
  console.log('----------------------------------------');
  try {
    const buffer = await createNestedTableDocument();
    const testPath = path.join(testOutputDir, 'nested-test-2.hwpx');
    fs.writeFileSync(testPath, buffer);

    let doc = await HwpxDocument.createFromBuffer('test2', testPath, buffer);

    // 앞에서부터 3번 삭제
    doc.deleteTable(0, 0);
    console.log('  Step 1: index 0 삭제');
    doc.deleteTable(0, 0);
    console.log('  Step 2: index 0 삭제');
    doc.deleteTable(0, 0);
    console.log('  Step 3: index 0 삭제');

    const savedBuffer = await doc.save();
    fs.writeFileSync(testPath, savedBuffer);

    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
    const savedBalance = checkTagBalance(savedXml, 'tbl');

    console.log(`  저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);

    if (!savedBalance.balanced || savedBalance.open !== 0) {
      console.log('  ❌ FAIL');
      failed++;
    } else {
      console.log('  ✅ PASS');
      passed++;
    }

    fs.unlinkSync(testPath);
  } catch (e) {
    console.log('  ❌ FAIL:', e.message);
    failed++;
  }

  // 테스트 3: 역순 삭제
  console.log('\n테스트 3: 역순 삭제 (index 2, 1, 0)');
  console.log('----------------------------------------');
  try {
    const buffer = await createNestedTableDocument();
    const testPath = path.join(testOutputDir, 'nested-test-3.hwpx');
    fs.writeFileSync(testPath, buffer);

    let doc = await HwpxDocument.createFromBuffer('test3', testPath, buffer);

    doc.deleteTable(0, 2);
    console.log('  index 2 삭제');
    doc.deleteTable(0, 1);
    console.log('  index 1 삭제');
    doc.deleteTable(0, 0);
    console.log('  index 0 삭제');

    const savedBuffer = await doc.save();
    fs.writeFileSync(testPath, savedBuffer);

    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
    const savedBalance = checkTagBalance(savedXml, 'tbl');

    console.log(`  저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);

    if (!savedBalance.balanced || savedBalance.open !== 0) {
      console.log('  ❌ FAIL');
      failed++;
    } else {
      console.log('  ✅ PASS');
      passed++;
    }

    fs.unlinkSync(testPath);
  } catch (e) {
    console.log('  ❌ FAIL:', e.message);
    failed++;
  }

  // 테스트 4: XML 무결성 검사
  console.log('\n테스트 4: XML 구조 무결성');
  console.log('----------------------------------------');
  try {
    const buffer = await createNestedTableDocument();
    const testPath = path.join(testOutputDir, 'nested-test-4.hwpx');
    fs.writeFileSync(testPath, buffer);

    let doc = await HwpxDocument.createFromBuffer('test4', testPath, buffer);

    doc.deleteTable(0, 0); // 중첩 테이블 포함 삭제

    const savedBuffer = await doc.save();
    fs.writeFileSync(testPath, savedBuffer);

    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

    const tagsToCheck = ['tbl', 'tr', 'tc', 'p', 'run', 'subList'];
    let allBalanced = true;

    for (const tag of tagsToCheck) {
      const balance = checkTagBalance(savedXml, tag);
      const status = balance.balanced ? '✅' : '❌';
      console.log(`  <hp:${tag}>: ${balance.open}/${balance.close} ${status}`);
      if (!balance.balanced) allBalanced = false;
    }

    if (!allBalanced) {
      console.log('  ❌ FAIL: XML 구조 불균형');
      failed++;
    } else {
      console.log('  ✅ PASS');
      passed++;
    }

    fs.unlinkSync(testPath);
  } catch (e) {
    console.log('  ❌ FAIL:', e.message);
    failed++;
  }

  // 결과 요약
  console.log('\n========================================');
  console.log(`  결과: ${passed} passed, ${failed} failed`);
  console.log('========================================\n');

  return failed === 0;
}

runTests()
  .then(success => process.exit(success ? 0 : 1))
  .catch(err => {
    console.error('Error:', err);
    process.exit(1);
  });
