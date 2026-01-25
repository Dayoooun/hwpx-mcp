import { describe, it, expect, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';

describe('Debug paragraph save', () => {
  const testFile = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';

  it('should debug replaceTextInElementByIndex', async () => {
    // Read the original XML
    const originalZip = await JSZip.loadAsync(fs.readFileSync(testFile));
    const originalXml = await originalZip.file('Contents/section0.xml')!.async('string');

    // Clean the XML (remove MEMO, footnote, endnote)
    const cleanedXml = originalXml
      .replace(/<hp:fieldBegin[^>]*type="MEMO"[^>]*>[\s\S]*?<\/hp:fieldBegin>/gi, '')
      .replace(/<hp:footNote\b[^>]*>[\s\S]*?<\/hp:footNote>/gi, '')
      .replace(/<hp:endNote\b[^>]*>[\s\S]*?<\/hp:endNote>/gi, '');

    // Extract all paragraphs
    const extractAllParagraphs = (xmlStr: string): { xml: string; start: number; end: number }[] => {
      const results: { xml: string; start: number; end: number }[] = [];
      const closeTag = '</hp:p>';
      const pOpenRegex = /<hp:p\b[^>]*>/g;
      const pOpenSearchRegex = /<hp:p[\s>]/g;
      let match;

      while ((match = pOpenRegex.exec(xmlStr)) !== null) {
        const startPos = match.index;
        let depth = 1;
        let searchPos = startPos + match[0].length;

        while (depth > 0 && searchPos < xmlStr.length) {
          pOpenSearchRegex.lastIndex = searchPos;
          const nextOpenMatch = pOpenSearchRegex.exec(xmlStr);
          const nextOpen = nextOpenMatch ? nextOpenMatch.index : -1;
          const nextClose = xmlStr.indexOf(closeTag, searchPos);

          if (nextClose === -1) break;

          if (nextOpen !== -1 && nextOpen < nextClose) {
            depth++;
            searchPos = nextOpen + 6;
          } else {
            depth--;
            if (depth === 0) {
              const endPos = nextClose + closeTag.length;
              results.push({
                xml: xmlStr.substring(startPos, endPos),
                start: startPos,
                end: endPos
              });
            }
            searchPos = nextClose + closeTag.length;
          }
        }
      }
      return results;
    };

    // Extract tables
    const extractBalancedTags = (xmlStr: string, tagName: string): string[] => {
      const results: string[] = [];
      const openTag = `<${tagName}`;
      const closeTag = `</${tagName}>`;
      let searchStart = 0;

      while (true) {
        const openIndex = xmlStr.indexOf(openTag, searchStart);
        if (openIndex === -1) break;

        let depth = 1;
        let pos = openIndex + openTag.length;

        while (depth > 0 && pos < xmlStr.length) {
          const nextOpen = xmlStr.indexOf(openTag, pos);
          const nextClose = xmlStr.indexOf(closeTag, pos);

          if (nextClose === -1) break;

          if (nextOpen !== -1 && nextOpen < nextClose) {
            depth++;
            pos = nextOpen + openTag.length;
          } else {
            depth--;
            if (depth === 0) {
              results.push(xmlStr.substring(openIndex, nextClose + closeTag.length));
            }
            pos = nextClose + closeTag.length;
          }
        }
        searchStart = openIndex + 1;
      }
      return results;
    };

    // Build cleaned elements array
    const paragraphs = extractAllParagraphs(cleanedXml);
    const tables = extractBalancedTags(cleanedXml, 'hp:tbl');

    const tableRanges: { start: number; end: number }[] = [];
    for (const tableXml of tables) {
      const tableIndex = cleanedXml.indexOf(tableXml);
      if (tableIndex !== -1) {
        tableRanges.push({ start: tableIndex, end: tableIndex + tableXml.length });
      }
    }

    interface CleanedElement {
      type: string;
      start: number;
      end: number;
      xml: string;
    }
    const elements: CleanedElement[] = [];

    // Add tables
    for (const range of tableRanges) {
      elements.push({ type: 'tbl', start: range.start, end: range.end, xml: cleanedXml.substring(range.start, range.end) });
    }

    // Add paragraphs
    for (const para of paragraphs) {
      const isInsideTable = tableRanges.some(
        range => para.start > range.start && para.start < range.end
      );
      if (!isInsideTable) {
        const hasTextContent = /<hp:t\b[^>]*>/.test(para.xml);
        if (hasTextContent || !tableRanges.some(range => range.start >= para.start && range.end <= para.end)) {
          elements.push({ type: 'p', start: para.start, end: para.end, xml: para.xml });
        }
      }
    }

    // Sort by position
    elements.sort((a, b) => a.start - b.start);

    console.log('Total elements in cleaned XML:', elements.length);
    console.log('Element 22 type:', elements[22]?.type);

    // Count paragraphs before index 22
    let pCountBefore22 = 0;
    for (let i = 0; i < 22; i++) {
      if (elements[i].type === 'p') pCountBefore22++;
    }
    console.log('Paragraphs before element 22:', pCountBefore22);

    // Build originalTopLevelParas
    const originalParagraphs = extractAllParagraphs(originalXml);
    const originalTables = extractBalancedTags(originalXml, 'hp:tbl');
    const originalTableRanges: { start: number; end: number }[] = [];
    for (const tableXml of originalTables) {
      const tableIndex = originalXml.indexOf(tableXml);
      if (tableIndex !== -1) {
        originalTableRanges.push({ start: tableIndex, end: tableIndex + tableXml.length });
      }
    }

    const originalTopLevelParas: { start: number; end: number; xml: string }[] = [];
    for (const para of originalParagraphs) {
      const isInsideTable = originalTableRanges.some(
        range => para.start > range.start && para.start < range.end
      );
      if (!isInsideTable) {
        originalTopLevelParas.push({ start: para.start, end: para.end, xml: para.xml });
      }
    }

    console.log('Total originalTopLevelParas:', originalTopLevelParas.length);

    // Check paragraph at index 13 (which is what we expect for element 22)
    const targetPara = originalTopLevelParas[pCountBefore22];
    console.log('Target paragraph (index ' + pCountBefore22 + '):');

    // Extract text from hp:t tags
    const textMatches = targetPara?.xml.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/g);
    if (textMatches) {
      console.log('  Text content:', textMatches.map(m => {
        const match = m.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/);
        return match ? match[1] : '';
      }).join(''));
    }

    // Also check what's at element 22
    console.log('\nElement 22 text:');
    const el22Text = elements[22]?.xml.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/g);
    if (el22Text) {
      console.log('  ', el22Text.map(m => {
        const match = m.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/);
        return match ? match[1] : '';
      }).join(''));
    }

    // Check if " ◦ " exists in originalTopLevelParas[13]
    const containsTarget = targetPara?.xml.includes(' ◦ ');
    console.log('\nTarget para contains " ◦ ":', containsTarget);

    // Find which paragraph has " ◦ "
    console.log('\nParagraphs containing " ◦ ":');
    originalTopLevelParas.forEach((para, idx) => {
      if (para.xml.includes(' ◦ ')) {
        const textMatch = para.xml.match(/<hp:t[^>]*>([^<]*◦[^<]*)<\/hp:t>/);
        console.log(`  Index ${idx}: ${textMatch?.[1]}`);
      }
    });

    expect(true).toBe(true);
  });
});
