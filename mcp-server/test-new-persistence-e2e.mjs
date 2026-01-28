import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Source and test file paths
const SOURCE_FILE = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';
const TEST_FILE = path.join(__dirname, 'test-output', 'new-persistence-e2e.hwpx');

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
  // Try dist first, fall back to ../out/mcp
  let HwpxDocument;
  try {
    ({ HwpxDocument } = await import('./dist/HwpxDocument.js'));
  } catch {
    ({ HwpxDocument } = await import('../out/mcp/HwpxDocument.js'));
  }
  const buffer = fs.readFileSync(TEST_FILE);
  return HwpxDocument.createFromBuffer('test-doc', TEST_FILE, buffer);
}

async function saveDocument(doc) {
  const buffer = await doc.save();
  fs.writeFileSync(TEST_FILE, buffer);
}

async function runTests() {
  console.log('=== New Persistence E2E Test Suite ===\n');

  // Setup: Copy source file to test location
  try {
    fs.mkdirSync(path.dirname(TEST_FILE), { recursive: true });
    fs.copyFileSync(SOURCE_FILE, TEST_FILE);
    console.log(`✓ Copied source file to ${TEST_FILE}\n`);
  } catch (err) {
    console.error(`Failed to setup test file: ${err.message}`);
    process.exit(1);
  }

  // ========================================
  // 1. insertTableRow persistence
  // ========================================
  console.log('### 1. insertTableRow persistence ###\n');

  try {
    let doc = await loadDocument();
    const tables = doc.getTables();

    if (typeof doc.insertTableRow === 'function' && tables && tables.length > 0) {
      const table = doc.getTable(tables[0].section, tables[0].index);
      const originalRowCount = table.rows;

      // Find a table with at least 2 rows for better testing
      let targetTable = null;
      for (const tbl of tables) {
        const t = doc.getTable(tbl.section, tbl.index);
        if (t.rows >= 2) {
          targetTable = { section: tbl.section, index: tbl.index, rowCount: t.rows };
          break;
        }
      }

      if (!targetTable) {
        FAIL('1. insertTableRow persistence', `No table with 2+ rows found. First table has ${originalRowCount} rows.`);
      } else {
        // Insert a row at index 0 (after first row)
        doc.insertTableRow(targetTable.section, targetTable.index, 0);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const tableAfter = doc.getTable(targetTable.section, targetTable.index);
        const newRowCount = tableAfter.rows;

        if (newRowCount === targetTable.rowCount + 1) {
          PASS('1. insertTableRow persistence', `rows: ${targetTable.rowCount} → ${newRowCount}`);
        } else {
          FAIL('1. insertTableRow persistence', `Expected ${targetTable.rowCount + 1} rows, got ${newRowCount}`);
        }
      }
    } else {
      FAIL('1. insertTableRow persistence', 'Method does not exist or no tables found');
    }
  } catch (err) {
    FAIL('1. insertTableRow persistence', err.message);
  }

  // ========================================
  // 2. deleteTableRow persistence
  // ========================================
  console.log('\n### 2. deleteTableRow persistence ###\n');

  try {
    let doc = await loadDocument();
    const tables = doc.getTables();

    if (typeof doc.deleteTableRow === 'function' && tables && tables.length > 0) {
      // Find a table with at least 3 rows for safe deletion
      let targetTable = null;
      for (const tbl of tables) {
        const t = doc.getTable(tbl.section, tbl.index);
        if (t.rows >= 3) {
          targetTable = { section: tbl.section, index: tbl.index, rowCount: t.rows };
          break;
        }
      }

      if (!targetTable) {
        FAIL('2. deleteTableRow persistence', 'Need at least one table with 3+ rows');
      } else {
        // Delete row at index 1 (middle row)
        doc.deleteTableRow(targetTable.section, targetTable.index, 1);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const tableAfter = doc.getTable(targetTable.section, targetTable.index);
        const newRowCount = tableAfter.rows;

        if (newRowCount === targetTable.rowCount - 1) {
          PASS('2. deleteTableRow persistence', `rows: ${targetTable.rowCount} → ${newRowCount}`);
        } else {
          FAIL('2. deleteTableRow persistence', `Expected ${targetTable.rowCount - 1} rows, got ${newRowCount}`);
        }
      }
    } else {
      FAIL('2. deleteTableRow persistence', 'Method does not exist or no tables found');
    }
  } catch (err) {
    FAIL('2. deleteTableRow persistence', err.message);
  }

  // ========================================
  // 3. insertTableColumn persistence
  // ========================================
  console.log('\n### 3. insertTableColumn persistence ###\n');

  try {
    let doc = await loadDocument();
    const tables = doc.getTables();

    if (typeof doc.insertTableColumn === 'function' && tables && tables.length > 0) {
      // Find a table with at least 2 columns
      let targetTable = null;
      for (const tbl of tables) {
        const t = doc.getTable(tbl.section, tbl.index);
        if (t.cols >= 2) {
          targetTable = { section: tbl.section, index: tbl.index, colCount: t.cols };
          break;
        }
      }

      if (!targetTable) {
        FAIL('3. insertTableColumn persistence', `No table with 2+ cols found`);
      } else {
        // Insert a column at index 0 (after first column)
        doc.insertTableColumn(targetTable.section, targetTable.index, 0);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const tableAfter = doc.getTable(targetTable.section, targetTable.index);
        const newColCount = tableAfter.cols;

        if (newColCount === targetTable.colCount + 1) {
          PASS('3. insertTableColumn persistence', `cols: ${targetTable.colCount} → ${newColCount}`);
        } else {
          FAIL('3. insertTableColumn persistence', `Expected ${targetTable.colCount + 1} cols, got ${newColCount}`);
        }
      }
    } else {
      FAIL('3. insertTableColumn persistence', 'Method does not exist or no tables found');
    }
  } catch (err) {
    FAIL('3. insertTableColumn persistence', err.message);
  }

  // ========================================
  // 4. deleteTableColumn persistence
  // ========================================
  console.log('\n### 4. deleteTableColumn persistence ###\n');

  try {
    let doc = await loadDocument();
    const tables = doc.getTables();

    if (typeof doc.deleteTableColumn === 'function' && tables && tables.length > 0) {
      // Find a table with at least 3 columns for safe deletion
      let targetTable = null;
      for (const tbl of tables) {
        const t = doc.getTable(tbl.section, tbl.index);
        if (t.cols >= 3) {
          targetTable = { section: tbl.section, index: tbl.index, colCount: t.cols };
          break;
        }
      }

      if (!targetTable) {
        FAIL('4. deleteTableColumn persistence', 'Need at least one table with 3+ cols');
      } else {
        // Delete column at index 1 (middle column)
        doc.deleteTableColumn(targetTable.section, targetTable.index, 1);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const tableAfter = doc.getTable(targetTable.section, targetTable.index);
        const newColCount = tableAfter.cols;

        if (newColCount === targetTable.colCount - 1) {
          PASS('4. deleteTableColumn persistence', `cols: ${targetTable.colCount} → ${newColCount}`);
        } else {
          FAIL('4. deleteTableColumn persistence', `Expected ${targetTable.colCount - 1} cols, got ${newColCount}`);
        }
      }
    } else {
      FAIL('4. deleteTableColumn persistence', 'Method does not exist or no tables found');
    }
  } catch (err) {
    FAIL('4. deleteTableColumn persistence', err.message);
  }

  // ========================================
  // 5. copyParagraph persistence
  // ========================================
  console.log('\n### 5. copyParagraph persistence ###\n');

  try {
    let doc = await loadDocument();
    const originalCount = doc.getParagraphs(0).length;

    if (typeof doc.copyParagraph === 'function') {
      const paragraphs = doc.getParagraphs(0);
      const validPara = paragraphs.find(p => p.index >= 5);

      if (!validPara) {
        FAIL('5. copyParagraph persistence', 'No valid paragraph found');
      } else {
        // Copy paragraph to after itself
        doc.copyParagraph(0, validPara.index, 0, validPara.index);
        await saveDocument(doc);

        // Reload and verify
        doc = await loadDocument();
        const newCount = doc.getParagraphs(0).length;

        if (newCount > originalCount) {
          PASS('5. copyParagraph persistence', `count: ${originalCount} → ${newCount}`);
        } else {
          FAIL('5. copyParagraph persistence', `Count did not increase: ${originalCount} → ${newCount}`);
        }
      }
    } else {
      FAIL('5. copyParagraph persistence', 'Method does not exist');
    }
  } catch (err) {
    FAIL('5. copyParagraph persistence', err.message);
  }

  // ========================================
  // 6. moveParagraph same-section persistence
  // ========================================
  console.log('\n### 6. moveParagraph same-section persistence ###\n');

  try {
    let doc = await loadDocument();

    if (typeof doc.moveParagraph === 'function') {
      const paragraphs = doc.getParagraphs(0);
      const sourcePara = paragraphs.find(p => p.index >= 5);
      const targetIndex = Math.min(10, paragraphs.length - 1);

      if (!sourcePara) {
        FAIL('6. moveParagraph same-section', 'No valid source paragraph found');
      } else {
        const originalText = doc.getParagraph(0, sourcePara.index).text;

        // Move paragraph to a different position
        doc.moveParagraph(0, sourcePara.index, 0, targetIndex);
        await saveDocument(doc);

        // Reload and verify text at new position
        doc = await loadDocument();
        const paragraphsAfter = doc.getParagraphs(0);
        const movedPara = paragraphsAfter[targetIndex + 1]; // After targetIndex

        if (movedPara && movedPara.text === originalText) {
          PASS('6. moveParagraph same-section', `text preserved: "${originalText.substring(0, 30)}..."`);
        } else {
          FAIL('6. moveParagraph same-section', `Text mismatch at target position`);
        }
      }
    } else {
      FAIL('6. moveParagraph same-section', 'Method does not exist');
    }
  } catch (err) {
    FAIL('6. moveParagraph same-section', err.message);
  }

  // ========================================
  // 7. setHeader persistence
  // ========================================
  console.log('\n### 7. setHeader persistence ###\n');

  try {
    let doc = await loadDocument();

    if (typeof doc.setHeader === 'function') {
      const testHeaderText = 'E2E 테스트 헤더';

      doc.setHeader(0, testHeaderText);
      await saveDocument(doc);

      // Reload and verify
      doc = await loadDocument();
      const header = doc.getHeader(0);

      if (header && header.paragraphs && header.paragraphs.length > 0) {
        const headerText = header.paragraphs[0].text;
        if (headerText.includes(testHeaderText)) {
          PASS('7. setHeader persistence', `verified: "${testHeaderText}"`);
        } else {
          FAIL('7. setHeader persistence', `Header text not found. Got: "${headerText}"`);
        }
      } else {
        FAIL('7. setHeader persistence', `No header paragraphs found`);
      }
    } else {
      FAIL('7. setHeader persistence', 'Method does not exist');
    }
  } catch (err) {
    FAIL('7. setHeader persistence', err.message);
  }

  // ========================================
  // 8. setFooter persistence
  // ========================================
  console.log('\n### 8. setFooter persistence ###\n');

  try {
    let doc = await loadDocument();

    if (typeof doc.setFooter === 'function') {
      const testFooterText = 'E2E 테스트 푸터';

      doc.setFooter(0, testFooterText);
      await saveDocument(doc);

      // Reload and verify
      doc = await loadDocument();
      const footer = doc.getFooter(0);

      if (footer && footer.paragraphs && footer.paragraphs.length > 0) {
        const footerText = footer.paragraphs[0].text;
        if (footerText.includes(testFooterText)) {
          PASS('8. setFooter persistence', `verified: "${testFooterText}"`);
        } else {
          FAIL('8. setFooter persistence', `Footer text not found. Got: "${footerText}"`);
        }
      } else {
        FAIL('8. setFooter persistence', `No footer paragraphs found`);
      }
    } else {
      FAIL('8. setFooter persistence', 'Method does not exist');
    }
  } catch (err) {
    FAIL('8. setFooter persistence', err.message);
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
