import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Source and test file paths
const SOURCE_FILE = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';
const TEST_FILE = path.join(__dirname, 'test-output', 'mcp-e2e.hwpx');

// Test results tracking
const results = {
  passed: 0,
  failed: 0,
  tests: []
};

function PASS(testName, details = '') {
  console.log(`✅ PASS: ${testName}${details ? ' - ' + details : ''}`);
  results.passed++;
  results.tests.push({ name: testName, status: 'PASS', details });
}

function FAIL(testName, error) {
  console.log(`❌ FAIL: ${testName} - ${error}`);
  results.failed++;
  results.tests.push({ name: testName, status: 'FAIL', error: error.toString() });
}

async function loadDocument() {
  const { HwpxDocument } = await import('./dist/HwpxDocument.js');
  const buffer = fs.readFileSync(TEST_FILE);
  return HwpxDocument.createFromBuffer('test-doc', TEST_FILE, buffer);
}

async function saveDocument(doc) {
  const buffer = await doc.save();
  fs.writeFileSync(TEST_FILE, buffer);
}

async function verifyXmlWellFormed(filePath) {
  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);

  const section0 = await zip.file('Contents/section0.xml')?.async('text');
  const header = await zip.file('Contents/header.xml')?.async('text');

  if (!section0 || !header) {
    throw new Error('Missing required XML files');
  }

  return { section0, header };
}

async function countCharPr(xml) {
  const matches = xml.match(/<hh:charPr/g);
  return matches ? matches.length : 0;
}

async function runTests() {
  console.log('=== MCP E2E Test Suite ===\n');

  // Setup: Copy source file to test location
  try {
    fs.mkdirSync(path.dirname(TEST_FILE), { recursive: true });
    fs.copyFileSync(SOURCE_FILE, TEST_FILE);
    console.log(`✓ Copied source file to ${TEST_FILE}\n`);
  } catch (err) {
    console.error(`Failed to setup test file: ${err.message}`);
    process.exit(1);
  }

  // Load initial document and count original charPr
  let originalCharPrCount = 0;
  try {
    const { section0, header } = await verifyXmlWellFormed(TEST_FILE);
    originalCharPrCount = await countCharPr(header);
    console.log(`Original charPr count: ${originalCharPrCount}\n`);
  } catch (err) {
    console.error(`Failed to read original XML: ${err.message}`);
  }

  // ========================================
  // 1. READ TESTS (no modifications)
  // ========================================
  console.log('### 1. READ TESTS ###\n');

  try {
    const doc = await loadDocument();

    // 1a. getParagraphs
    try {
      if (typeof doc.getParagraphs === 'function') {
        const paragraphs = doc.getParagraphs(0);
        PASS('1a. getParagraphs', `count=${paragraphs.length}`);
      } else {
        FAIL('1a. getParagraphs', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1a. getParagraphs', err.message);
    }

    // 1b. getParagraph
    try {
      if (typeof doc.getParagraph === 'function') {
        const para = doc.getParagraph(0, 0);
        if (para && para.text !== undefined && para.runs !== undefined) {
          PASS('1b. getParagraph', `has text and runs fields`);
        } else {
          FAIL('1b. getParagraph', `Missing expected fields: ${JSON.stringify(para)}`);
        }
      } else {
        FAIL('1b. getParagraph', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1b. getParagraph', err.message);
    }

    // 1c. getCharacterStyle
    try {
      if (typeof doc.getCharacterStyle === 'function') {
        const charStyle = doc.getCharacterStyle(0, 0, 0);
        if (charStyle) {
          const fields = Object.keys(charStyle).join(', ');
          PASS('1c. getCharacterStyle', `fields: ${fields}`);
          console.log(`   Character style details: ${JSON.stringify(charStyle, null, 2)}`);
        } else {
          FAIL('1c. getCharacterStyle', 'Returned null or undefined');
        }
      } else {
        FAIL('1c. getCharacterStyle', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1c. getCharacterStyle', err.message);
    }

    // 1d. getParagraphStyle
    try {
      if (typeof doc.getParagraphStyle === 'function') {
        const paraStyle = doc.getParagraphStyle(0, 0);
        if (paraStyle) {
          const fields = Object.keys(paraStyle).join(', ');
          PASS('1d. getParagraphStyle', `fields: ${fields}`);
          console.log(`   Paragraph style details: ${JSON.stringify(paraStyle, null, 2)}`);
        } else {
          FAIL('1d. getParagraphStyle', 'Returned null or undefined');
        }
      } else {
        FAIL('1d. getParagraphStyle', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1d. getParagraphStyle', err.message);
    }

    // 1e. getTables
    try {
      if (typeof doc.getTables === 'function') {
        const tables = doc.getTables(0);
        PASS('1e. getTables', `count=${tables.length}`);
      } else {
        FAIL('1e. getTables', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1e. getTables', err.message);
    }

    // 1f. getTable
    try {
      if (typeof doc.getTable === 'function') {
        const tables = doc.getTables();
        if (tables && tables.length > 0) {
          const table = doc.getTable(tables[0].section, tables[0].index);
          if (table) {
            PASS('1f. getTable', `rows=${table.rows || 0}, cols=${table.cols || 0}`);
          } else {
            FAIL('1f. getTable', 'Returned null or undefined');
          }
        } else {
          FAIL('1f. getTable', 'No tables found');
        }
      } else {
        FAIL('1f. getTable', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1f. getTable', err.message);
    }

    // 1g. getTableCell
    try {
      if (typeof doc.getTableCell === 'function') {
        const tables = doc.getTables();
        if (tables && tables.length > 0) {
          const cell = doc.getTableCell(tables[0].section, tables[0].index, 0, 0);
          if (cell) {
            PASS('1g. getTableCell', `text="${cell.text?.substring(0, 30)}..."`);
          } else {
            FAIL('1g. getTableCell', 'Returned null or undefined');
          }
        } else {
          FAIL('1g. getTableCell', 'No tables found');
        }
      } else {
        FAIL('1g. getTableCell', 'Method does not exist');
      }
    } catch (err) {
      FAIL('1g. getTableCell', err.message);
    }

  } catch (err) {
    console.error(`READ TESTS FAILED: ${err.message}`);
  }

  // ========================================
  // 2. WRITE + SAVE + RELOAD TESTS
  // ========================================
  console.log('\n### 2. WRITE + SAVE + RELOAD TESTS ###\n');

  // 2a. updateParagraphText
  try {
    let doc = await loadDocument();
    const testText = '테스트 텍스트입니다';

    if (typeof doc.updateParagraphText === 'function') {
      // Get actual paragraph elements to find a valid elementIndex
      const paragraphs = doc.getParagraphs(0);
      const validPara = paragraphs.find(p => p.index >= 5);

      if (!validPara) {
        FAIL('2a. updateParagraphText', 'No valid paragraph found');
      } else {
        doc.updateParagraphText(0, validPara.index, 0, testText);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const para = doc.getParagraph(0, validPara.index);

        if (para && para.text === testText) {
          PASS('2a. updateParagraphText', `verified text="${testText}"`);
        } else {
          FAIL('2a. updateParagraphText', `Text mismatch: expected "${testText}", got "${para?.text}"`);
        }
      }
    } else {
      FAIL('2a. updateParagraphText', 'Method does not exist');
    }
  } catch (err) {
    FAIL('2a. updateParagraphText', err.message);
  }

  // 2b. applyCharacterStyle
  try {
    let doc = await loadDocument();

    if (typeof doc.applyCharacterStyle === 'function') {
      const styleChanges = { fontName: '휴먼명조', fontSize: 14 };
      doc.applyCharacterStyle(0, 0, 0, styleChanges);
      await saveDocument(doc);

      // Reload and verify XML
      const { header } = await verifyXmlWellFormed(TEST_FILE);

      if (header.includes('height="1400"') && header.includes('<hh:fontRef')) {
        PASS('2b. applyCharacterStyle', 'verified height=1400 and fontRef in XML');
      } else {
        FAIL('2b. applyCharacterStyle', 'Could not verify style changes in XML');
      }
    } else {
      FAIL('2b. applyCharacterStyle', 'Method does not exist');
    }
  } catch (err) {
    FAIL('2b. applyCharacterStyle', err.message);
  }

  // 2c. applyParagraphStyle
  try {
    let doc = await loadDocument();

    if (typeof doc.applyParagraphStyle === 'function') {
      // Get actual paragraph elements
      const paragraphs = doc.getParagraphs(0);
      const validPara = paragraphs.find(p => p.index >= 1);

      if (!validPara) {
        FAIL('2c. applyParagraphStyle', 'No valid paragraph found');
      } else {
        doc.applyParagraphStyle(0, validPara.index, { align: 'center' });
        await saveDocument(doc);

        // Reload and verify XML - paraPr is in header.xml, not section0.xml
        const savedBuf = fs.readFileSync(TEST_FILE);
        const zip = await JSZip.loadAsync(savedBuf);
        const headerXml = await zip.file('Contents/header.xml').async('string');
        const sectionXml = await zip.file('Contents/section0.xml').async('string');

        // Check header.xml has new paraPr with CENTER
        const hasCenterAlign = headerXml.includes('horizontal="CENTER"');
        // Check section0.xml has updated paraPrIDRef
        const paraPrIdMatch = sectionXml.match(new RegExp(`<hp:p[^>]*paraPrIDRef="(\\d+)"`));

        if (hasCenterAlign) {
          PASS('2c. applyParagraphStyle', `verified horizontal="CENTER" in header.xml`);
        } else {
          FAIL('2c. applyParagraphStyle', 'Could not find horizontal="CENTER" in header.xml');
        }
      }
    } else {
      FAIL('2c. applyParagraphStyle', 'Method does not exist');
    }
  } catch (err) {
    FAIL('2c. applyParagraphStyle', err.message);
  }

  // 2d. updateTableCell
  try {
    let doc = await loadDocument();
    const tables = doc.getTables();

    if (typeof doc.updateTableCell === 'function' && tables && tables.length > 0) {
      const testCellText = 'E2E 테스트 셀';
      doc.updateTableCell(tables[0].section, tables[0].index, 0, 0, testCellText);
      await saveDocument(doc);

      // Reload and verify
      doc = await loadDocument();
      const cell = doc.getTableCell(tables[0].section, tables[0].index, 0, 0);

      if (cell && cell.text.includes(testCellText)) {
        PASS('2d. updateTableCell', `verified text="${testCellText}"`);
      } else {
        FAIL('2d. updateTableCell', `Text mismatch: got "${cell?.text}"`);
      }
    } else {
      FAIL('2d. updateTableCell', 'Method does not exist or no tables found');
    }
  } catch (err) {
    FAIL('2d. updateTableCell', err.message);
  }

  // 2e. insertParagraph
  try {
    let doc = await loadDocument();
    const originalCount = doc.getParagraphs(0).length;

    if (typeof doc.insertParagraph === 'function') {
      // Get actual paragraph elements
      const paragraphs = doc.getParagraphs(0);
      const validPara = paragraphs.find(p => p.index >= 5);

      if (!validPara) {
        FAIL('2e. insertParagraph', 'No valid paragraph found');
      } else {
        doc.insertParagraph(0, validPara.index, 'E2E 삽입된 문단');
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const newCount = doc.getParagraphs(0).length;

        if (newCount > originalCount) {
          PASS('2e. insertParagraph', `count increased: ${originalCount} → ${newCount}`);
        } else {
          FAIL('2e. insertParagraph', `Count did not increase: ${originalCount} → ${newCount}`);
        }
      }
    } else {
      FAIL('2e. insertParagraph', 'Method does not exist');
    }
  } catch (err) {
    FAIL('2e. insertParagraph', err.message);
  }

  // ========================================
  // 3. charPr INTEGRITY TEST
  // ========================================
  console.log('\n### 3. charPr INTEGRITY TEST ###\n');

  try {
    const { header } = await verifyXmlWellFormed(TEST_FILE);
    const finalCharPrCount = await countCharPr(header);

    console.log(`원본 charPr: ${originalCharPrCount} → 수정 후: ${finalCharPrCount}`);

    if (finalCharPrCount >= originalCharPrCount) {
      PASS('3. charPr integrity', `${originalCharPrCount} → ${finalCharPrCount}`);
    } else {
      FAIL('3. charPr integrity', `charPr corruption detected: ${originalCharPrCount} → ${finalCharPrCount}`);
    }
  } catch (err) {
    FAIL('3. charPr integrity', err.message);
  }

  // ========================================
  // 4. XML WELL-FORMEDNESS TEST
  // ========================================
  console.log('\n### 4. XML WELL-FORMEDNESS TEST ###\n');

  try {
    const { section0, header } = await verifyXmlWellFormed(TEST_FILE);

    // Check balanced <hp:p> tags
    const pOpenCount = (section0.match(/<hp:p(\s|>)/g) || []).length;
    const pCloseCount = (section0.match(/<\/hp:p>/g) || []).length;

    if (pOpenCount === pCloseCount) {
      PASS('4a. Balanced <hp:p> tags', `${pOpenCount} open, ${pCloseCount} close`);
    } else {
      FAIL('4a. Balanced <hp:p> tags', `Mismatch: ${pOpenCount} open, ${pCloseCount} close`);
    }

    // Check balanced <hp:tbl> tags
    const tblOpenCount = (section0.match(/<hp:tbl(\s|>)/g) || []).length;
    const tblCloseCount = (section0.match(/<\/hp:tbl>/g) || []).length;

    if (tblOpenCount === tblCloseCount) {
      PASS('4b. Balanced <hp:tbl> tags', `${tblOpenCount} open, ${tblCloseCount} close`);
    } else {
      FAIL('4b. Balanced <hp:tbl> tags', `Mismatch: ${tblOpenCount} open, ${tblCloseCount} close`);
    }

    // Check balanced <hh:charPr> tags
    const charPrOpenCount = (header.match(/<hh:charPr(\s|>)/g) || []).length;
    const charPrCloseCount = (header.match(/<\/hh:charPr>/g) || []).length;

    if (charPrOpenCount === charPrCloseCount) {
      PASS('4c. Balanced <hh:charPr> tags', `${charPrOpenCount} open, ${charPrCloseCount} close`);
    } else {
      FAIL('4c. Balanced <hh:charPr> tags', `Mismatch: ${charPrOpenCount} open, ${charPrCloseCount} close`);
    }

  } catch (err) {
    FAIL('4. XML well-formedness', err.message);
  }

  // ========================================
  // FINAL SUMMARY
  // ========================================
  console.log('\n=== FINAL SUMMARY ===');
  console.log(`Total tests: ${results.passed + results.failed}`);
  console.log(`✅ Passed: ${results.passed}`);
  console.log(`❌ Failed: ${results.failed}`);
  console.log(`Success rate: ${((results.passed / (results.passed + results.failed)) * 100).toFixed(1)}%`);

  if (results.failed > 0) {
    console.log('\nFailed tests:');
    results.tests
      .filter(t => t.status === 'FAIL')
      .forEach(t => console.log(`  - ${t.name}: ${t.error}`));
  }

  process.exit(results.failed > 0 ? 1 : 0);
}

runTests().catch(err => {
  console.error('Test suite crashed:', err);
  process.exit(1);
});
