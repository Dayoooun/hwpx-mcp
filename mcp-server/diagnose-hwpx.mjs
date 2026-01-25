#!/usr/bin/env node
/**
 * HWPX 문서 진단 도구
 *
 * 사용법: node diagnose-hwpx.mjs <hwpx파일경로> [section_index]
 *
 * 이 도구는:
 * 1. HWPX 파일의 XML 구조를 분석
 * 2. 문단별 텍스트와 run 구조 출력
 * 3. 잠재적 문제점 식별
 */

import { HwpxDocument } from './dist/HwpxDocument.js';
import JSZip from 'jszip';
import * as fs from 'fs';

async function diagnose(filePath, sectionIdx = 0) {
  console.log('='.repeat(80));
  console.log('HWPX 문서 진단 도구');
  console.log('='.repeat(80));
  console.log(`파일: ${filePath}`);
  console.log(`섹션: ${sectionIdx}`);
  console.log('');

  // 파일 존재 확인
  if (!fs.existsSync(filePath)) {
    console.error('❌ 파일을 찾을 수 없습니다:', filePath);
    process.exit(1);
  }

  // HWPX 파일 로드
  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);

  // section XML 추출
  const sectionPath = `Contents/section${sectionIdx}.xml`;
  const sectionFile = zip.file(sectionPath);
  if (!sectionFile) {
    console.error(`❌ ${sectionPath}를 찾을 수 없습니다`);
    process.exit(1);
  }

  const sectionXml = await sectionFile.async('string');

  console.log('\n=== 원본 XML (처음 5000자) ===\n');
  console.log(sectionXml.substring(0, 5000));
  if (sectionXml.length > 5000) {
    console.log(`\n... (${sectionXml.length - 5000}자 더 있음)`);
  }

  // HwpxDocument로 파싱
  console.log('\n\n=== 파싱된 문단 목록 ===\n');

  const doc = await HwpxDocument.createFromBuffer('diagnose', filePath, buffer);

  // 모든 문단 출력
  let elementIndex = 0;
  const elements = doc.listSectionElements(sectionIdx);

  for (const element of elements) {
    const prefix = `[${elementIndex.toString().padStart(3)}]`;

    if (element.type === 'paragraph') {
      const para = doc.getParagraph(sectionIdx, elementIndex);
      const text = para?.text || '';
      const runs = para?.runs || [];

      // 문제 감지
      const warnings = [];

      // 1. 텍스트 중복 감지
      const words = text.split(/\s+/).filter(w => w.length > 3);
      const wordCounts = {};
      for (const word of words) {
        wordCounts[word] = (wordCounts[word] || 0) + 1;
      }
      for (const [word, count] of Object.entries(wordCounts)) {
        if (count > 2 && word.length > 5) {
          warnings.push(`⚠️ "${word}"가 ${count}번 반복됨`);
        }
      }

      // 2. 비정상적으로 긴 텍스트
      if (text.length > 500) {
        warnings.push(`⚠️ 매우 긴 문단 (${text.length}자)`);
      }

      // 3. Run 개수 확인
      if (runs.length > 5) {
        warnings.push(`⚠️ 많은 run 수 (${runs.length}개)`);
      }

      // 출력
      const truncatedText = text.length > 100 ? text.substring(0, 100) + '...' : text;
      console.log(`${prefix} [P] ${truncatedText}`);
      console.log(`       Runs: ${runs.length}개`);

      if (runs.length <= 5) {
        for (let i = 0; i < runs.length; i++) {
          const runText = runs[i].text || '';
          const truncatedRun = runText.length > 60 ? runText.substring(0, 60) + '...' : runText;
          console.log(`       Run[${i}]: "${truncatedRun}"`);
        }
      }

      for (const warning of warnings) {
        console.log(`       ${warning}`);
      }

      console.log('');
    } else if (element.type === 'table') {
      const table = element.data;
      console.log(`${prefix} [T] 테이블 (${table.rows}행 x ${table.cols}열)`);
      console.log('');
    } else {
      console.log(`${prefix} [${element.type}]`);
      console.log('');
    }

    elementIndex++;
  }

  // XML 구조 분석
  console.log('\n=== XML 구조 분석 ===\n');

  // hp:p 태그 개수
  const pTagCount = (sectionXml.match(/<hp:p\b/g) || []).length;
  const pCloseCount = (sectionXml.match(/<\/hp:p>/g) || []).length;
  console.log(`<hp:p> 태그: ${pTagCount}개 열림, ${pCloseCount}개 닫힘`);

  if (pTagCount !== pCloseCount) {
    console.log('❌ 태그 불균형 감지!');
  }

  // hp:run 태그 개수
  const runTagCount = (sectionXml.match(/<hp:run\b/g) || []).length;
  console.log(`<hp:run> 태그: ${runTagCount}개`);

  // hp:t 태그 개수
  const tTagCount = (sectionXml.match(/<hp:t\b/g) || []).length;
  console.log(`<hp:t> 태그: ${tTagCount}개`);

  // hp:tbl 태그 개수
  const tblTagCount = (sectionXml.match(/<hp:tbl\b/g) || []).length;
  console.log(`<hp:tbl> 태그: ${tblTagCount}개 (테이블)`);

  // 특정 인덱스 상세 분석
  console.log('\n\n=== 인덱스 22 상세 분석 ===\n');

  if (elementIndex > 22) {
    const para22 = doc.getParagraph(sectionIdx, 22);
    if (para22) {
      console.log('텍스트 전체:');
      console.log(para22.text);
      console.log('\nRuns:');
      for (let i = 0; i < para22.runs.length; i++) {
        console.log(`  [${i}] charStyle: ${para22.runs[i].charStyle || 'none'}`);
        console.log(`      text: "${para22.runs[i].text}"`);
      }
    }
  } else {
    console.log('인덱스 22가 존재하지 않습니다 (총 요소 수:', elementIndex, ')');
  }

  console.log('\n=== 진단 완료 ===');
}

// 실행
const args = process.argv.slice(2);
if (args.length < 1) {
  console.log('사용법: node diagnose-hwpx.mjs <hwpx파일경로> [section_index]');
  console.log('');
  console.log('예시:');
  console.log('  node diagnose-hwpx.mjs "C:\\Users\\문서.hwpx"');
  console.log('  node diagnose-hwpx.mjs "C:\\Users\\문서.hwpx" 0');
  process.exit(1);
}

const filePath = args[0];
const sectionIdx = parseInt(args[1] || '0', 10);

diagnose(filePath, sectionIdx).catch(err => {
  console.error('진단 중 오류 발생:', err);
  process.exit(1);
});
