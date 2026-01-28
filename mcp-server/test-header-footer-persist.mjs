import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Test file paths
const SOURCE_FILE = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';
const TEST_FILE = path.join(__dirname, 'test-output', 'header-footer-test.hwpx');

// Colors for terminal output
const GREEN = '\x1b[32m';
const RED = '\x1b[31m';
const YELLOW = '\x1b[33m';
const RESET = '\x1b[0m';

function log(msg) {
  console.log(msg);
}

function success(msg) {
  console.log(`${GREEN}✓${RESET} ${msg}`);
}

function error(msg) {
  console.log(`${RED}✗${RESET} ${msg}`);
}

function info(msg) {
  console.log(`${YELLOW}ℹ${RESET} ${msg}`);
}

async function checkHeaderFooterInXml(filePath) {
  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const section0 = await zip.file('Contents/section0.xml')?.async('string');

  if (!section0) {
    return { hasHeader: false, hasFooter: false, headerText: null, footerText: null };
  }

  // Check for header
  const headerMatch = section0.match(/<hp:header[\s\S]*?<\/hp:header>/);
  const hasHeader = !!headerMatch;
  let headerText = null;

  if (hasHeader && headerMatch) {
    const textMatch = headerMatch[0].match(/<hp:t>([\s\S]*?)<\/hp:t>/);
    headerText = textMatch ? textMatch[1] : '';
  }

  // Check for footer
  const footerMatch = section0.match(/<hp:footer[\s\S]*?<\/hp:footer>/);
  const hasFooter = !!footerMatch;
  let footerText = null;

  if (hasFooter && footerMatch) {
    const textMatch = footerMatch[0].match(/<hp:t>([\s\S]*?)<\/hp:t>/);
    footerText = textMatch ? textMatch[1] : '';
  }

  return { hasHeader, hasFooter, headerText, footerText };
}

async function runTest() {
  log('\n=== Header/Footer Persistence Test ===\n');

  // Setup: Copy source file to test location
  try {
    fs.mkdirSync(path.dirname(TEST_FILE), { recursive: true });
    fs.copyFileSync(SOURCE_FILE, TEST_FILE);
    success(`Copied source file to ${TEST_FILE}`);
  } catch (err) {
    error(`Failed to setup test file: ${err.message}`);
    process.exit(1);
  }

  // Import HwpxDocument
  const { HwpxDocument } = await import('./dist/HwpxDocument.js');

  // Test 1: Set header and verify in XML after save
  log('\n--- Test 1: Set Header ---');
  try {
    let doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const headerText = 'Test Header 테스트 헤더';
    const result = doc.setHeader(0, headerText);

    if (!result) {
      error('setHeader returned false');
      process.exit(1);
    }
    success('setHeader called successfully');

    // Save and reload
    const buffer = await doc.save();
    fs.writeFileSync(TEST_FILE, buffer);
    success('Saved document');

    // Check XML
    const xmlCheck = await checkHeaderFooterInXml(TEST_FILE);
    if (!xmlCheck.hasHeader) {
      error('Header not found in XML after save');
      process.exit(1);
    }
    success('Header found in XML');

    if (xmlCheck.headerText !== headerText) {
      error(`Header text mismatch: expected "${headerText}", got "${xmlCheck.headerText}"`);
      process.exit(1);
    }
    success(`Header text matches: "${headerText}"`);

    // Reload document and verify in memory
    doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const header = doc.getHeader(0);
    if (!header) {
      error('getHeader returned null after reload');
      process.exit(1);
    }
    success('getHeader returned data after reload');

    const reloadedText = header.paragraphs[0]?.text || '';
    if (reloadedText !== headerText) {
      error(`Reloaded header text mismatch: expected "${headerText}", got "${reloadedText}"`);
      process.exit(1);
    }
    success(`Reloaded header text matches: "${headerText}"`);

  } catch (err) {
    error(`Test 1 failed: ${err.message}`);
    console.error(err);
    process.exit(1);
  }

  // Test 2: Set footer and verify in XML after save
  log('\n--- Test 2: Set Footer ---');
  try {
    let doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const footerText = 'Test Footer 테스트 푸터';
    const result = doc.setFooter(0, footerText);

    if (!result) {
      error('setFooter returned false');
      process.exit(1);
    }
    success('setFooter called successfully');

    // Save and reload
    const buffer = await doc.save();
    fs.writeFileSync(TEST_FILE, buffer);
    success('Saved document');

    // Check XML
    const xmlCheck = await checkHeaderFooterInXml(TEST_FILE);
    if (!xmlCheck.hasFooter) {
      error('Footer not found in XML after save');
      process.exit(1);
    }
    success('Footer found in XML');

    if (xmlCheck.footerText !== footerText) {
      error(`Footer text mismatch: expected "${footerText}", got "${xmlCheck.footerText}"`);
      process.exit(1);
    }
    success(`Footer text matches: "${footerText}"`);

    // Reload document and verify in memory
    doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const footer = doc.getFooter(0);
    if (!footer) {
      error('getFooter returned null after reload');
      process.exit(1);
    }
    success('getFooter returned data after reload');

    const reloadedText = footer.paragraphs[0]?.text || '';
    if (reloadedText !== footerText) {
      error(`Reloaded footer text mismatch: expected "${footerText}", got "${reloadedText}"`);
      process.exit(1);
    }
    success(`Reloaded footer text matches: "${footerText}"`);

  } catch (err) {
    error(`Test 2 failed: ${err.message}`);
    console.error(err);
    process.exit(1);
  }

  // Test 3: Set both header and footer together
  log('\n--- Test 3: Set Both Header and Footer ---');
  try {
    let doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const headerText = 'Combined Test Header 결합 테스트 헤더';
    const footerText = 'Combined Test Footer 결합 테스트 푸터';

    doc.setHeader(0, headerText);
    doc.setFooter(0, footerText);
    success('Set both header and footer');

    // Save and reload
    const buffer = await doc.save();
    fs.writeFileSync(TEST_FILE, buffer);
    success('Saved document');

    // Check XML
    const xmlCheck = await checkHeaderFooterInXml(TEST_FILE);
    if (!xmlCheck.hasHeader || !xmlCheck.hasFooter) {
      error('Header or footer not found in XML after save');
      process.exit(1);
    }
    success('Both header and footer found in XML');

    if (xmlCheck.headerText !== headerText || xmlCheck.footerText !== footerText) {
      error(`Text mismatch in XML`);
      process.exit(1);
    }
    success('Both header and footer text match in XML');

    // Reload document and verify in memory
    doc = await HwpxDocument.createFromBuffer(
      'test-doc',
      TEST_FILE,
      fs.readFileSync(TEST_FILE)
    );

    const header = doc.getHeader(0);
    const footer = doc.getFooter(0);

    if (!header || !footer) {
      error('getHeader or getFooter returned null after reload');
      process.exit(1);
    }
    success('Both getHeader and getFooter returned data after reload');

    const reloadedHeaderText = header.paragraphs[0]?.text || '';
    const reloadedFooterText = footer.paragraphs[0]?.text || '';

    if (reloadedHeaderText !== headerText || reloadedFooterText !== footerText) {
      error('Reloaded text mismatch');
      process.exit(1);
    }
    success(`Reloaded texts match: header="${headerText}", footer="${footerText}"`);

  } catch (err) {
    error(`Test 3 failed: ${err.message}`);
    console.error(err);
    process.exit(1);
  }

  log('\n=== All Tests Passed! ===\n');
  process.exit(0);
}

runTest().catch(err => {
  error(`Test suite crashed: ${err.message}`);
  console.error(err);
  process.exit(1);
});
