#!/usr/bin/env node
/**
 * Test script to verify paragraph indexing fix
 * Tests that replaceTextInElementByIndex correctly handles documents with tables
 */

import { createRequire } from 'module';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const require = createRequire(import.meta.url);
const { HwpxDocument } = require('./dist/HwpxDocument.js');

async function testParagraphIndexing() {
  console.log('Testing paragraph indexing fix...\n');

  // Open test document
  const testFilePath = path.join(__dirname, 'test-temp.hwpx');
  const data = await fs.readFile(testFilePath);

  const testDoc = await HwpxDocument.createFromBuffer('test-doc-1', testFilePath, data);

  // Get all elements to see the structure
  const elements = testDoc.getElements();
  console.log('Document structure:');
  elements.forEach((el, idx) => {
    if (el.type === 'paragraph') {
      console.log(`  [${idx}] paragraph: "${el.text.substring(0, 50)}${el.text.length > 50 ? '...' : ''}"`);
    } else if (el.type === 'table') {
      console.log(`  [${idx}] table (${el.rows} rows x ${el.cols} cols)`);
    }
  });

  console.log('\n--- Testing text replacement ---');

  // Find a paragraph to test
  const paragraphIndex = elements.findIndex(el => el.type === 'paragraph' && el.text.trim().length > 0);

  if (paragraphIndex === -1) {
    console.log('No paragraph found to test');
    return false;
  }

  const targetPara = elements[paragraphIndex];
  const oldText = targetPara.text;
  const newText = 'TEST REPLACEMENT TEXT';

  console.log(`Target element index: ${paragraphIndex}`);
  console.log(`Original text: "${oldText}"`);
  console.log(`New text: "${newText}"\n`);

  // Perform replacement
  testDoc.replaceTextInElement(paragraphIndex, oldText, newText);

  // Save and reload to verify
  const testOutputDir = path.join(__dirname, 'test-output');
  await fs.mkdir(testOutputDir, { recursive: true });

  const testOutputPath = path.join(testOutputDir, 'paragraph-indexing-test.hwpx');
  const outputBuffer = await testDoc.toBuffer();
  await fs.writeFile(testOutputPath, outputBuffer);
  console.log(`Saved to: ${testOutputPath}`);

  // Reload and verify
  const verifyData = await fs.readFile(testOutputPath);
  const verifyDoc = await HwpxDocument.createFromBuffer('test-doc-2', testOutputPath, verifyData);
  const verifyElements = verifyDoc.getElements();
  const verifyPara = verifyElements[paragraphIndex];

  console.log(`\nVerification:`);
  console.log(`  Element ${paragraphIndex} text: "${verifyPara.text}"`);

  if (verifyPara.text.includes(newText)) {
    console.log('  ✓ Text replacement SUCCESSFUL');
    return true;
  } else {
    console.log('  ✗ Text replacement FAILED');
    console.log(`  Expected to find: "${newText}"`);
    console.log(`  Got: "${verifyPara.text}"`);
    return false;
  }
}

testParagraphIndexing().then(success => {
  process.exit(success ? 0 : 1);
}).catch(err => {
  console.error('Test error:', err);
  process.exit(1);
});
