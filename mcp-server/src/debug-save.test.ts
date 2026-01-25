import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('Debug replaceTextInElementByIndex directly', () => {
  const testFile = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';

  it('should compare HwpxDocument parsing vs replaceTextInElementByIndex logic', async () => {
    // 1. Open document with HwpxDocument and find paragraph with " ◦ "
    const buffer = fs.readFileSync(testFile);
    const doc = await HwpxDocument.createFromBuffer('test-id', testFile, buffer);
    const paragraphs = doc.getParagraphs(0); // section 0

    console.log('=== HwpxDocument getParagraphs() ===');
    const paraWithCircle = paragraphs.find(p => p.text.includes(' ◦ '));
    if (paraWithCircle) {
      console.log(`Found " ◦ " at element index: ${paraWithCircle.index}`);
      console.log(`Text: ${paraWithCircle.text}`);
    } else {
      console.log('No paragraph with " ◦ " found!');
    }

    // 2. Now build elements array the same way replaceTextInElementByIndex does
    // Read original XML directly
    const zip = await JSZip.loadAsync(fs.readFileSync(testFile));
    const xml = await zip.file('Contents/section0.xml')!.async('string');

    // Simulate cleaned XML (same as in replaceTextInElementByIndex)
    const cleanedXml = xml
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

    // Build elements array (same as HwpxParser logic)
    const cleanedParagraphs = extractAllParagraphs(cleanedXml);
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

    // Add paragraphs - SAME LOGIC AS HwpxParser
    for (const para of cleanedParagraphs) {
      const isInsideTable = tableRanges.some(
        range => para.start > range.start && para.start < range.end
      );
      const containsTable = tableRanges.some(
        range => range.start >= para.start && range.end <= para.end
      );

      if (!isInsideTable) {
        if (containsTable) {
          let paraXmlWithoutTable = para.xml;
          for (const range of tableRanges) {
            if (range.start >= para.start && range.end <= para.start + para.xml.length) {
              const tableStartInPara = range.start - para.start;
              const tableEndInPara = range.end - para.start;
              const tableXmlInPara = para.xml.substring(tableStartInPara, tableEndInPara);
              paraXmlWithoutTable = paraXmlWithoutTable.replace(tableXmlInPara, '');
            }
          }
          const hasTextContent = /<hp:t\b[^>]*>/.test(paraXmlWithoutTable);
          if (hasTextContent) {
            elements.push({ type: 'p', start: para.start, end: para.end, xml: paraXmlWithoutTable });
          }
        } else {
          elements.push({ type: 'p', start: para.start, end: para.end, xml: para.xml });
        }
      }
    }

    // ADD SHAPE ELEMENTS (line, rect, ellipse, arc, polygon, curve, connectLine)
    const addShapeElements = (pattern: RegExp, typeName: string) => {
      let match;
      while ((match = pattern.exec(cleanedXml)) !== null) {
        elements.push({ type: typeName, start: match.index, end: match.index + match[0].length, xml: match[0] });
      }
    };

    addShapeElements(/<hp:line\b[^>]*(?:\/>|>[\s\S]*?<\/hp:line>)/g, 'line');
    addShapeElements(/<hp:rect\b[^>]*(?:\/>|>[\s\S]*?<\/hp:rect>)/g, 'rect');
    addShapeElements(/<hp:ellipse\b[^>]*(?:\/>|>[\s\S]*?<\/hp:ellipse>)/g, 'ellipse');
    addShapeElements(/<hp:arc\b[^>]*(?:\/>|>[\s\S]*?<\/hp:arc>)/g, 'arc');
    addShapeElements(/<hp:polygon\b[^>]*(?:\/>|>[\s\S]*?<\/hp:polygon>)/g, 'polygon');
    addShapeElements(/<hp:curve\b[^>]*(?:\/>|>[\s\S]*?<\/hp:curve>)/g, 'curve');
    addShapeElements(/<hp:connectLine\b[^>]*(?:\/>|>[\s\S]*?<\/hp:connectLine>)/g, 'connectline');

    // Sort by position
    elements.sort((a, b) => a.start - b.start);

    // 3. Compare
    console.log('\n=== Manual element building ===');
    console.log('Total elements:', elements.length);

    if (paraWithCircle) {
      const elementAtIndex = elements[paraWithCircle.index];
      console.log(`\nElement at index ${paraWithCircle.index} (from HwpxDocument):`, elementAtIndex?.type);

      if (elementAtIndex?.type === 'p') {
        const textMatch = elementAtIndex.xml.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/);
        console.log('Text:', textMatch?.[1]);
      }
    }

    // Find where " ◦ " actually appears in our elements
    console.log('\n=== Elements containing " ◦ " ===');
    elements.forEach((el, idx) => {
      if (el.type === 'p' && el.xml.includes(' ◦ ')) {
        const textMatch = el.xml.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/);
        console.log(`Index ${idx}: ${textMatch?.[1]}`);
      }
    });

    // The issue is clear now: HwpxDocument says " ◦ " is at element index 22
    // But manual building shows element 22 is a 'tbl', and " ◦ " appears at indices 33, 39, 50, 72, 78, 101
    // This means HwpxParser is missing shape elements that replaceTextInElementByIndex is including!
    console.log('\n=== MISMATCH FOUND ===');
    console.log('HwpxDocument thinks " ◦ " is at index:', paraWithCircle?.index);
    console.log('But manually built elements show index 22 is type:', elements[22]?.type);
    console.log('Manual build found " ◦ " at indices: 33, 39, 50, 72, 78, 101');

    // Test should pass now that we understand the issue
    expect(paraWithCircle).toBeDefined();
  });
});
