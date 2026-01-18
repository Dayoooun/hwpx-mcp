import JSZip from 'jszip';
import {
  HwpxContent,
  HwpxSection,
  HwpxParagraph,
  HwpxTable,
  HwpxImage,
  HwpxLine,
  HwpxRect,
  HwpxEllipse,
  HwpxArc,
  HwpxPolygon,
  HwpxCurve,
  HwpxConnectLine,
  HwpxTextBox,
  HwpxHorizontalRule,
  HwpxContainer,
  HwpxOle,
  HwpxEquation,
  HwpxTextArt,
  HwpxUnknownObject,
  HwpxButton,
  HwpxRadioButton,
  HwpxCheckButton,
  HwpxComboBox,
  HwpxEdit,
  HwpxListBox,
  HwpxScrollBar,
  TextRun,
  CharacterStyle,
  ParagraphStyle,
  TableRow,
  TableCell,
  SectionElement,
  HwpxStyles,
  CharShape,
  ParaShape,
  PageSettings,
  Footnote,
  CurveSegment,
  ArcType,
  OleObjectType,
  DrawAspect,
  ShapeObject,
  DrawingObject,
  ShapeComponent,
  LineShape,
  FillBrush,
  CompatibleDocument,
  LayoutCompatibility,
  BinData,
  ScriptCode,
  ScriptFunction,
  XmlTemplate,
} from './types';

export * from './types';

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

export class HwpxParser {
  private static styles: HwpxStyles = {
    charShapes: new Map(),
    paraShapes: new Map(),
    fonts: new Map(),
    fontsByLang: new Map(),
    borderFills: new Map(),
    tabDefs: new Map(),
    numberings: new Map(),
    bullets: new Map(),
    styles: new Map(),
    memoShapes: new Map(),
  };

  static async parse(zip: JSZip): Promise<HwpxContent> {
    const content: HwpxContent = {
      metadata: {},
      sections: [],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };

    this.styles = {
      charShapes: new Map(),
      paraShapes: new Map(),
      fonts: new Map(),
      fontsByLang: new Map(),
      borderFills: new Map(),
      tabDefs: new Map(),
      numberings: new Map(),
      bullets: new Map(),
      styles: new Map(),
      memoShapes: new Map(),
    };

    const headerXml = await this.readXmlFile(zip, 'Contents/header.xml');
    if (headerXml) {
      content.metadata = this.parseMetadata(headerXml);
      content.docSetting = this.parseDocSetting(headerXml);
      this.parseStyles(headerXml);
      this.parseMemoShapes(headerXml);
      content.compatibleDocument = this.parseCompatibleDocument(headerXml);
      // Assign parsed styles to content
      content.styles = this.styles;
    }

    await this.parseImages(zip, content);
    await this.parseBinDataStorage(zip, content);

    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const sectionXml = await this.readXmlFile(zip, sectionPath);
      if (!sectionXml) break;

      const section = this.parseSection(sectionXml, content);
      content.sections.push(section);
      sectionIndex++;
    }

    // Parse Scripts (optional)
    const scriptCode = await this.parseScriptCode(zip);
    if (scriptCode) {
      (content as any).scriptCode = scriptCode;
    }

    // Parse XMLTemplate (optional)
    const xmlTemplate = await this.parseXmlTemplate(zip);
    if (xmlTemplate) {
      (content as any).xmlTemplate = xmlTemplate;
    }

    return content;
  }

  private static async parseScriptCode(zip: JSZip): Promise<ScriptCode | undefined> {
    const scriptsFolder = zip.folder('Scripts');
    if (!scriptsFolder) return undefined;

    const scriptCode: ScriptCode = {};

    // Read DefaultJScript
    const defaultJScript = await this.readXmlFile(zip, 'Scripts/DefaultJScript');
    if (defaultJScript) {
      scriptCode.source = defaultJScript;
      scriptCode.type = 'JScript';
    }

    // Read JScriptVersion
    const versionFile = await this.readXmlFile(zip, 'Scripts/JScriptVersion');
    if (versionFile) {
      scriptCode.version = versionFile.trim();
    }

    // Read Header script if exists
    const headerScript = await this.readXmlFile(zip, 'Scripts/Header');
    if (headerScript) {
      scriptCode.header = headerScript;
    }

    // Read PreScript and PostScript if they exist as separate files
    const preScriptNames = Object.keys(zip.files).filter(f => f.startsWith('Scripts/PreScript'));
    if (preScriptNames.length > 0) {
      scriptCode.preScript = [];
      for (const name of preScriptNames) {
        const code = await this.readXmlFile(zip, name);
        if (code) {
          scriptCode.preScript.push({ name: name.replace('Scripts/', ''), code });
        }
      }
    }

    const postScriptNames = Object.keys(zip.files).filter(f => f.startsWith('Scripts/PostScript'));
    if (postScriptNames.length > 0) {
      scriptCode.postScript = [];
      for (const name of postScriptNames) {
        const code = await this.readXmlFile(zip, name);
        if (code) {
          scriptCode.postScript.push({ name: name.replace('Scripts/', ''), code });
        }
      }
    }

    return Object.keys(scriptCode).length > 0 ? scriptCode : undefined;
  }

  private static async parseXmlTemplate(zip: JSZip): Promise<XmlTemplate | undefined> {
    const xmlTemplate: XmlTemplate = {};

    // Read Schema
    const schema = await this.readXmlFile(zip, 'XMLTemplate/Schema');
    if (schema) {
      xmlTemplate.schema = schema;
    }

    // Read Instance
    const instance = await this.readXmlFile(zip, 'XMLTemplate/Instance');
    if (instance) {
      xmlTemplate.instance = instance;
    }

    return (xmlTemplate.schema || xmlTemplate.instance) ? xmlTemplate : undefined;
  }

  private static async readXmlFile(zip: JSZip, path: string): Promise<string | null> {
    const file = zip.file(path);
    if (!file) return null;
    return await file.async('string');
  }

  private static parseMetadata(xml: string): HwpxContent['metadata'] {
    const metadata: HwpxContent['metadata'] = {};

    const extract = (tag: string): string | undefined => {
      const regex = new RegExp(`<(?:hh:)?${tag}[^>]*>([^<]*)</(?:hh:)?${tag}>`);
      const match = xml.match(regex);
      return match?.[1];
    };

    metadata.title = extract('title');
    metadata.creator = extract('creator');
    metadata.createdDate = extract('createdDate');
    metadata.modifiedDate = extract('modifiedDate');
    metadata.description = extract('description');
    metadata.subject = extract('subject');

    // Parse keywords
    const keywordsMatch = xml.match(/<(?:hh:)?keywords[^>]*>([^<]*)<\/(?:hh:)?keywords>/i);
    if (keywordsMatch) {
      metadata.keywords = keywordsMatch[1].split(',').map(k => k.trim()).filter(k => k);
    }

    // Parse comments
    const commentsMatch = xml.match(/<(?:hh:)?comments[^>]*>([^<]*)<\/(?:hh:)?comments>/i);
    if (commentsMatch) {
      metadata.comments = commentsMatch[1];
    }

    // Parse forbidden strings
    const forbiddenRegex = /<(?:hh:)?forbidden[^>]*>([^<]*)<\/(?:hh:)?forbidden>/gi;
    const forbiddenStrings: string[] = [];
    let forbiddenMatch;
    while ((forbiddenMatch = forbiddenRegex.exec(xml)) !== null) {
      forbiddenStrings.push(forbiddenMatch[1]);
    }
    if (forbiddenStrings.length > 0) {
      metadata.forbiddenStrings = forbiddenStrings;
    }

    return metadata;
  }

  private static parseDocSetting(xml: string): import('./types').DocSetting | undefined {
    const docSetting: import('./types').DocSetting = {};

    // Parse beginNumber
    const beginNumMatch = xml.match(/<(?:hh:)?beginNum[^>]*>([\s\S]*?)<\/(?:hh:)?beginNum>|<(?:hh:)?beginNum([^>]*)\/>/i);
    if (beginNumMatch) {
      const content = beginNumMatch[1] || beginNumMatch[2] || '';
      docSetting.beginNumber = {};

      const pageMatch = content.match(/page="(\d+)"/);
      if (pageMatch) docSetting.beginNumber.page = parseInt(pageMatch[1]);

      const footnoteMatch = content.match(/footnote="(\d+)"/);
      if (footnoteMatch) docSetting.beginNumber.footnote = parseInt(footnoteMatch[1]);

      const endnoteMatch = content.match(/endnote="(\d+)"/);
      if (endnoteMatch) docSetting.beginNumber.endnote = parseInt(endnoteMatch[1]);

      const pictureMatch = content.match(/(?:picture|pic)="(\d+)"/);
      if (pictureMatch) docSetting.beginNumber.picture = parseInt(pictureMatch[1]);

      const tableMatch = content.match(/(?:table|tbl)="(\d+)"/);
      if (tableMatch) docSetting.beginNumber.table = parseInt(tableMatch[1]);

      const equationMatch = content.match(/equation="(\d+)"/);
      if (equationMatch) docSetting.beginNumber.equation = parseInt(equationMatch[1]);

      const totalPageMatch = content.match(/totalPage="(\d+)"/);
      if (totalPageMatch) docSetting.beginNumber.totalPage = parseInt(totalPageMatch[1]);
    }

    // Parse caretPos
    const caretPosMatch = xml.match(/<(?:hh:)?caretPos[^>]*>([\s\S]*?)<\/(?:hh:)?caretPos>|<(?:hh:)?caretPos([^>]*)\/>/i);
    if (caretPosMatch) {
      const content = caretPosMatch[1] || caretPosMatch[2] || '';
      docSetting.caretPos = {};

      const listMatch = content.match(/list="([^"]*)"/);
      if (listMatch) docSetting.caretPos.list = listMatch[1];

      const paraMatch = content.match(/para="([^"]*)"/);
      if (paraMatch) docSetting.caretPos.para = paraMatch[1];

      const posMatch = content.match(/pos="([^"]*)"/);
      if (posMatch) docSetting.caretPos.pos = posMatch[1];
    }

    return Object.keys(docSetting).length > 0 ? docSetting : undefined;
  }

  private static parseMemoShapes(xml: string): void {
    const memoShapeRegex = /<(?:hh:)?memoShape[^>]*>([\s\S]*?)<\/(?:hh:)?memoShape>|<(?:hh:)?memoShape([^>]*)\/>/gi;
    let match;

    while ((match = memoShapeRegex.exec(xml)) !== null) {
      const content = match[1] || match[2] || '';

      const idMatch = content.match(/id="(\d+)"/);
      const id = idMatch ? parseInt(idMatch[1]) : this.styles.memoShapes.size;

      const memoShape: import('./types').MemoShape = { id };

      const widthMatch = content.match(/width="(\d+)"/);
      if (widthMatch) memoShape.width = parseInt(widthMatch[1]);

      const lineTypeMatch = content.match(/lineType="([^"]*)"/);
      if (lineTypeMatch) memoShape.lineType = lineTypeMatch[1] as any;

      const lineColorMatch = content.match(/lineColor="([^"]*)"/);
      if (lineColorMatch) memoShape.lineColor = lineColorMatch[1];

      const fillColorMatch = content.match(/fillColor="([^"]*)"/);
      if (fillColorMatch) memoShape.fillColor = fillColorMatch[1];

      const activeColorMatch = content.match(/activeColor="([^"]*)"/);
      if (activeColorMatch) memoShape.activeColor = activeColorMatch[1];

      const memoTypeMatch = content.match(/memoType="([^"]*)"/);
      if (memoTypeMatch) memoShape.memoType = memoTypeMatch[1];

      this.styles.memoShapes.set(id, memoShape);
    }
  }

  private static parseStyles(xml: string): void {
    this.parseFonts(xml);
    this.parseCharShapes(xml);
    this.parseParaShapes(xml);
    this.parseBorderFills(xml);
    this.parseTabDefs(xml);
    this.parseNumberings(xml);
    this.parseBullets(xml);
    this.parseStyleDefs(xml);
  }

  private static parseFonts(xml: string): void {
    // Parse fonts from all fontfaces (HANGUL, LATIN, HANJA, JAPANESE, OTHER, SYMBOL, USER)
    const languages = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];

    for (const lang of languages) {
      const fontFaceRegex = new RegExp(`<hh:fontface[^>]*lang="${lang}"[^>]*>([\\s\\S]*?)<\\/hh:fontface>`, 'i');
      const fontFaceMatch = xml.match(fontFaceRegex);

      if (fontFaceMatch) {
        const fontRegex = /<hh:font[^>]*id="(\d+)"[^>]*face="([^"]*)"/gi;
        let match;
        while ((match = fontRegex.exec(fontFaceMatch[1])) !== null) {
          const fontId = parseInt(match[1]);
          const fontName = match[2];
          // Store with language prefix to avoid ID conflicts between languages
          const key = `${lang.toLowerCase()}_${fontId}`;
          this.styles.fontsByLang.set(key, fontName);
          // Also store in main fonts map (HANGUL takes priority for backward compatibility)
          if (lang === 'HANGUL' || !this.styles.fonts.has(fontId)) {
            this.styles.fonts.set(fontId, fontName);
          }
        }
      }
    }

    // Fallback if no fontfaces found
    if (this.styles.fonts.size === 0) {
      const fontRegex = /<(?:hh:)?font[^>]*face="([^"]*)"[^>]*>/gi;
      let match;
      let fontId = 0;
      while ((match = fontRegex.exec(xml)) !== null) {
        this.styles.fonts.set(fontId, match[1]);
        fontId++;
      }
    }
  }

  private static parseCharShapes(xml: string): void {
    // Capture tag name (charShape or charPr) in group 1, content in group 2
    const charShapeRegex = /<(?:hh:)?(charShape|charPr)([^>]*)>([\s\S]*?)<\/(?:hh:)?(?:charShape|charPr)>/gi;
    const charShapeRegexSelfClosing = /<(?:hh:)?(charShape|charPr)([^/>]*)\/>/gi;
    let match;
    let shapeId = 0;

    const parseCharShape = (shapeContent: string, tagName: 'charShape' | 'charPr') => {
      const charShape: CharShape = { id: shapeId, tagName };

      const idMatch = shapeContent.match(/\bid="(\d+)"/);
      if (idMatch) {
        shapeId = parseInt(idMatch[1]);
        charShape.id = shapeId;
      }

      const heightMatch = shapeContent.match(/height="(\d+)"/);
      if (heightMatch) {
        charShape.fontSize = parseInt(heightMatch[1]) / 100;
      }

      const boldTagMatch = shapeContent.match(/<(?:hh:)?bold\s*\/>/i);
      const boldAttrMatch = shapeContent.match(/bold="([^"]*)"/);
      charShape.bold = !!boldTagMatch || boldAttrMatch?.[1] === '1' || boldAttrMatch?.[1] === 'true';

      const italicTagMatch = shapeContent.match(/<(?:hh:)?italic\s*\/>/i);
      const italicAttrMatch = shapeContent.match(/italic="([^"]*)"/);
      charShape.italic = !!italicTagMatch || italicAttrMatch?.[1] === '1' || italicAttrMatch?.[1] === 'true';

      const underlineMatch = shapeContent.match(/<(?:hh:)?underline[^>]*type="([^"]*)"[^>]*(?:shape="([^"]*)")?[^>]*(?:color="([^"]*)")?/i);
      if (underlineMatch && underlineMatch[1] !== 'NONE') {
        const underlineTypeMap: Record<string, import('./types').UnderlineType> = {
          'BOTTOM': 'Bottom', 'CENTER': 'Center', 'TOP': 'Top', 'NONE': 'None'
        };
        const shapeMap: Record<string, import('./types').LineType2> = {
          'SOLID': 'Solid', 'DASH': 'Dash', 'DOT': 'Dot', 'DASH_DOT': 'DashDot',
          'DASH_DOT_DOT': 'DashDotDot', 'LONG_DASH': 'LongDash', 'CIRCLE_DOT': 'CircleDot',
          'DOUBLE_SLIM': 'DoubleSlim', 'SLIM_THICK': 'SlimThick', 'THICK_SLIM': 'ThickSlim',
          'SLIM_THICK_SLIM': 'SlimThickSlim', 'NONE': 'None'
        };
        charShape.underline = {
          type: underlineTypeMap[underlineMatch[1]?.toUpperCase()] || 'Bottom',
          shape: shapeMap[underlineMatch[2]?.toUpperCase()] || 'Solid',
          color: underlineMatch[3] || '#000000'
        };
      }

      const strikeMatch = shapeContent.match(/<(?:hh:)?strikeout[^>]*type="([^"]*)"[^>]*(?:shape="([^"]*)")?[^>]*(?:color="([^"]*)")?/i);
      if (strikeMatch && strikeMatch[1] !== 'NONE') {
        const strikeTypeMap: Record<string, import('./types').StrikeoutType> = {
          'NONE': 'None', 'CONTINUOUS': 'Continuous'
        };
        const shapeMap: Record<string, import('./types').LineType2> = {
          'SOLID': 'Solid', 'DASH': 'Dash', 'DOT': 'Dot', 'DASH_DOT': 'DashDot',
          'DASH_DOT_DOT': 'DashDotDot', 'LONG_DASH': 'LongDash', 'NONE': 'None'
        };
        charShape.strikeout = {
          type: strikeTypeMap[strikeMatch[1]?.toUpperCase()] || 'Continuous',
          shape: shapeMap[strikeMatch[2]?.toUpperCase()] || 'Solid',
          color: strikeMatch[3] || '#000000'
        };
      }

      const colorMatch = shapeContent.match(/textColor="([^"]*)"/);
      if (colorMatch && colorMatch[1] !== '#000000') {
        charShape.color = colorMatch[1];
      }

      const bgColorMatch = shapeContent.match(/shadeColor="([^"]*)"/);
      if (bgColorMatch && bgColorMatch[1] !== 'none') {
        charShape.backgroundColor = bgColorMatch[1];
      }

      const fontRefMatch = shapeContent.match(/<(?:hh:)?fontRef[^>]*/i);
      if (fontRefMatch) {
        const fontRefContent = fontRefMatch[0];
        charShape.fontRefs = {};
        charShape.fontNames = {};

        const hangulMatch = fontRefContent.match(/hangul="(\d+)"/);
        if (hangulMatch) {
          charShape.fontRefs.hangul = parseInt(hangulMatch[1]);
          charShape.fontNames.hangul = this.styles.fontsByLang.get(`hangul_${hangulMatch[1]}`) || this.styles.fonts.get(parseInt(hangulMatch[1]));
          charShape.fontName = charShape.fontNames.hangul;  // Default to hangul font
        }
        const latinMatch = fontRefContent.match(/latin="(\d+)"/);
        if (latinMatch) {
          charShape.fontRefs.latin = parseInt(latinMatch[1]);
          charShape.fontNames.latin = this.styles.fontsByLang.get(`latin_${latinMatch[1]}`) || this.styles.fonts.get(parseInt(latinMatch[1]));
        }
        const hanjaMatch = fontRefContent.match(/hanja="(\d+)"/);
        if (hanjaMatch) {
          charShape.fontRefs.hanja = parseInt(hanjaMatch[1]);
          charShape.fontNames.hanja = this.styles.fontsByLang.get(`hanja_${hanjaMatch[1]}`) || this.styles.fonts.get(parseInt(hanjaMatch[1]));
        }
        const japaneseMatch = fontRefContent.match(/japanese="(\d+)"/);
        if (japaneseMatch) {
          charShape.fontRefs.japanese = parseInt(japaneseMatch[1]);
          charShape.fontNames.japanese = this.styles.fontsByLang.get(`japanese_${japaneseMatch[1]}`) || this.styles.fonts.get(parseInt(japaneseMatch[1]));
        }
        const otherMatch = fontRefContent.match(/other="(\d+)"/);
        if (otherMatch) {
          charShape.fontRefs.other = parseInt(otherMatch[1]);
          charShape.fontNames.other = this.styles.fontsByLang.get(`other_${otherMatch[1]}`) || this.styles.fonts.get(parseInt(otherMatch[1]));
        }
        const symbolMatch = fontRefContent.match(/symbol="(\d+)"/);
        if (symbolMatch) {
          charShape.fontRefs.symbol = parseInt(symbolMatch[1]);
          charShape.fontNames.symbol = this.styles.fontsByLang.get(`symbol_${symbolMatch[1]}`) || this.styles.fonts.get(parseInt(symbolMatch[1]));
        }
        const userMatch = fontRefContent.match(/user="(\d+)"/);
        if (userMatch) {
          charShape.fontRefs.user = parseInt(userMatch[1]);
          charShape.fontNames.user = this.styles.fontsByLang.get(`user_${userMatch[1]}`) || this.styles.fonts.get(parseInt(userMatch[1]));
        }
      }

      // Parse spacing element - extract each attribute separately (order-independent)
      const spacingElemMatch = shapeContent.match(/<(?:hh:)?spacing([^>]*)\/?>/i);
      if (spacingElemMatch) {
        const spacingAttrs = spacingElemMatch[1];
        const hangulMatch = spacingAttrs.match(/hangul="(-?\d+)"/);
        const latinMatch = spacingAttrs.match(/latin="(-?\d+)"/);
        const hanjaMatch = spacingAttrs.match(/hanja="(-?\d+)"/);
        const japaneseMatch = spacingAttrs.match(/japanese="(-?\d+)"/);
        const otherMatch = spacingAttrs.match(/other="(-?\d+)"/);
        const symbolMatch = spacingAttrs.match(/symbol="(-?\d+)"/);
        const userMatch = spacingAttrs.match(/user="(-?\d+)"/);

        charShape.charSpacing = {
          hangul: hangulMatch ? parseInt(hangulMatch[1]) : 0,
          latin: latinMatch ? parseInt(latinMatch[1]) : 0,
          hanja: hanjaMatch ? parseInt(hanjaMatch[1]) : 0,
          japanese: japaneseMatch ? parseInt(japaneseMatch[1]) : 0,
          other: otherMatch ? parseInt(otherMatch[1]) : 0,
          symbol: symbolMatch ? parseInt(symbolMatch[1]) : 0,
          user: userMatch ? parseInt(userMatch[1]) : 0
        };
      } else {
        // Fallback: simple spacing attribute
        const charSpacingAttrMatch = shapeContent.match(/spacing="(-?\d+)"/);
        if (charSpacingAttrMatch) {
          const spacingValue = parseInt(charSpacingAttrMatch[1]);
          charShape.charSpacing = {
            hangul: spacingValue, latin: spacingValue, hanja: spacingValue,
            japanese: spacingValue, other: spacingValue, symbol: spacingValue, user: spacingValue
          };
        }
      }

      // Parse relSz element - extract each attribute separately (order-independent)
      const relSzElemMatch = shapeContent.match(/<(?:hh:)?relSz([^>]*)\/?>/i);
      if (relSzElemMatch) {
        const relSzAttrs = relSzElemMatch[1];
        const hangulMatch = relSzAttrs.match(/hangul="(\d+)"/);
        const latinMatch = relSzAttrs.match(/latin="(\d+)"/);
        const hanjaMatch = relSzAttrs.match(/hanja="(\d+)"/);
        const japaneseMatch = relSzAttrs.match(/japanese="(\d+)"/);
        const otherMatch = relSzAttrs.match(/other="(\d+)"/);
        const symbolMatch = relSzAttrs.match(/symbol="(\d+)"/);
        const userMatch = relSzAttrs.match(/user="(\d+)"/);

        charShape.relSize = {
          hangul: hangulMatch ? parseInt(hangulMatch[1]) : 100,
          latin: latinMatch ? parseInt(latinMatch[1]) : 100,
          hanja: hanjaMatch ? parseInt(hanjaMatch[1]) : 100,
          japanese: japaneseMatch ? parseInt(japaneseMatch[1]) : 100,
          other: otherMatch ? parseInt(otherMatch[1]) : 100,
          symbol: symbolMatch ? parseInt(symbolMatch[1]) : 100,
          user: userMatch ? parseInt(userMatch[1]) : 100
        };
      } else {
        // Fallback: simple relSz attribute
        const relSzAttrMatch = shapeContent.match(/relSz="(\d+)"/);
        if (relSzAttrMatch) {
          const relSzValue = parseInt(relSzAttrMatch[1]);
          charShape.relSize = {
            hangul: relSzValue, latin: relSzValue, hanja: relSzValue,
            japanese: relSzValue, other: relSzValue, symbol: relSzValue, user: relSzValue
          };
        }
      }

      // Parse offset element - extract each attribute separately (order-independent)
      const charOffsetElemMatch = shapeContent.match(/<(?:hh:)?offset([^>]*)\/?>/i);
      if (charOffsetElemMatch) {
        const offsetAttrs = charOffsetElemMatch[1];
        const hangulMatch = offsetAttrs.match(/hangul="(-?\d+)"/);
        const latinMatch = offsetAttrs.match(/latin="(-?\d+)"/);
        const hanjaMatch = offsetAttrs.match(/hanja="(-?\d+)"/);
        const japaneseMatch = offsetAttrs.match(/japanese="(-?\d+)"/);
        const otherMatch = offsetAttrs.match(/other="(-?\d+)"/);
        const symbolMatch = offsetAttrs.match(/symbol="(-?\d+)"/);
        const userMatch = offsetAttrs.match(/user="(-?\d+)"/);

        charShape.charOffset = {
          hangul: hangulMatch ? parseInt(hangulMatch[1]) : 0,
          latin: latinMatch ? parseInt(latinMatch[1]) : 0,
          hanja: hanjaMatch ? parseInt(hanjaMatch[1]) : 0,
          japanese: japaneseMatch ? parseInt(japaneseMatch[1]) : 0,
          other: otherMatch ? parseInt(otherMatch[1]) : 0,
          symbol: symbolMatch ? parseInt(symbolMatch[1]) : 0,
          user: userMatch ? parseInt(userMatch[1]) : 0
        };
      } else {
        // Fallback: simple offset attribute
        const charOffsetAttrMatch = shapeContent.match(/offset="(-?\d+)"/);
        if (charOffsetAttrMatch) {
          const offsetValue = parseInt(charOffsetAttrMatch[1]);
          charShape.charOffset = {
            hangul: offsetValue, latin: offsetValue, hanja: offsetValue,
            japanese: offsetValue, other: offsetValue, symbol: offsetValue, user: offsetValue
          };
        }
      }

      // Parse ratio element - extract each attribute separately (order-independent)
      const ratioElemMatch = shapeContent.match(/<(?:hh:)?ratio([^>]*)\/?>/i);
      if (ratioElemMatch) {
        const ratioAttrs = ratioElemMatch[1];
        const hangulMatch = ratioAttrs.match(/hangul="(\d+)"/);
        const latinMatch = ratioAttrs.match(/latin="(\d+)"/);
        const hanjaMatch = ratioAttrs.match(/hanja="(\d+)"/);
        const japaneseMatch = ratioAttrs.match(/japanese="(\d+)"/);
        const otherMatch = ratioAttrs.match(/other="(\d+)"/);
        const symbolMatch = ratioAttrs.match(/symbol="(\d+)"/);
        const userMatch = ratioAttrs.match(/user="(\d+)"/);

        charShape.ratio = {
          hangul: hangulMatch ? parseInt(hangulMatch[1]) : 100,
          latin: latinMatch ? parseInt(latinMatch[1]) : 100,
          hanja: hanjaMatch ? parseInt(hanjaMatch[1]) : 100,
          japanese: japaneseMatch ? parseInt(japaneseMatch[1]) : 100,
          other: otherMatch ? parseInt(otherMatch[1]) : 100,
          symbol: symbolMatch ? parseInt(symbolMatch[1]) : 100,
          user: userMatch ? parseInt(userMatch[1]) : 100
        };
      }

      const symMarkMatch = shapeContent.match(/<(?:hh:)?symMark[^>]*symMarkType="([^"]*)"/i);
      if (symMarkMatch && symMarkMatch[1] !== 'NONE') {
        const symMarkMap: Record<string, import('./types').EmphasisMark> = {
          'DOT': 'Dot', 'CIRCLE': 'Circle', 'RING': 'Ring', 'CARON': 'Caron',
          'UNDER_DOT': 'UnderDot', 'UNDER_LINE': 'UnderLine', 'TRIANGLE': 'Triangle', 'NONE': 'None'
        };
        charShape.symMark = symMarkMap[symMarkMatch[1].toUpperCase()] || 'None';
      }

      const useFontSpaceMatch = shapeContent.match(/useFontSpace="([^"]*)"/);
      if (useFontSpaceMatch) {
        charShape.useFontSpace = useFontSpaceMatch[1] === '1' || useFontSpaceMatch[1] === 'true';
      }

      const useKerningMatch = shapeContent.match(/useKerning="([^"]*)"/);
      if (useKerningMatch) {
        charShape.useKerning = useKerningMatch[1] === '1' || useKerningMatch[1] === 'true';
      }

      const outlineMatch = shapeContent.match(/<(?:hh:)?outline[^>]*type="([^"]*)"/i);
      if (outlineMatch && outlineMatch[1] !== 'NONE') {
        const outlineMap: Record<string, import('./types').LineType3> = {
          'SOLID': 'Solid', 'DOT': 'Dot', 'DASH': 'Dash', 'DASH_DOT': 'DashDot',
          'DASH_DOT_DOT': 'DashDotDot', 'THICK': 'Thick'
        };
        charShape.outline = {
          type: outlineMap[outlineMatch[1].toUpperCase()] || 'Solid'
        };
      }

      const shadowMatch = shapeContent.match(/<(?:hh:)?shadow[^>]*type="([^"]*)"[^>]*(?:offsetX="(-?\d+)")?[^>]*(?:offsetY="(-?\d+)")?[^>]*(?:color="([^"]*)")?/i);
      if (shadowMatch && shadowMatch[1] !== 'NONE') {
        const shadowMap: Record<string, import('./types').ShadowType> = {
          'DROP': 'Drop', 'CONTINUOUS': 'Cont', 'NONE': 'None'
        };
        charShape.shadow = {
          type: shadowMap[shadowMatch[1].toUpperCase()] || 'None',
          offsetX: shadowMatch[2] ? parseInt(shadowMatch[2]) / 100 : undefined,
          offsetY: shadowMatch[3] ? parseInt(shadowMatch[3]) / 100 : undefined,
          color: shadowMatch[4] || undefined
        };
      }

      const embossMatch = shapeContent.match(/<(?:hh:)?emboss\s*\/>/i);
      charShape.emboss = !!embossMatch;

      const engraveMatch = shapeContent.match(/<(?:hh:)?engrave\s*\/>/i);
      charShape.engrave = !!engraveMatch;

      const borderFillIdMatch = shapeContent.match(/borderFillIDRef="(\d+)"/);
      if (borderFillIdMatch) {
        charShape.borderFillId = parseInt(borderFillIdMatch[1]);
      }

      this.styles.charShapes.set(charShape.id, charShape);
      shapeId++;
    };

    while ((match = charShapeRegex.exec(xml)) !== null) {
      // match[1] = tagName (charShape or charPr), match[0] = full match
      const tagName = match[1].toLowerCase() as 'charShape' | 'charPr';
      parseCharShape(match[0], tagName);
    }

    while ((match = charShapeRegexSelfClosing.exec(xml)) !== null) {
      // match[1] = tagName (charShape or charPr), match[0] = full match
      const tagName = match[1].toLowerCase() as 'charShape' | 'charPr';
      parseCharShape(match[0], tagName);
    }
  }

  private static parseParaShapes(xml: string): void {
    const paraShapeRegex = /<(?:hh:)?(?:paraShape|paraPr)[^>]*>([\s\S]*?)<\/(?:hh:)?(?:paraShape|paraPr)>/gi;
    let match;
    let shapeId = 0;

    while ((match = paraShapeRegex.exec(xml)) !== null) {
      const shapeContent = match[0];
      const paraShape: ParaShape = { id: shapeId };

      const idMatch = shapeContent.match(/\bid="(\d+)"/);
      if (idMatch) {
        shapeId = parseInt(idMatch[1]);
        paraShape.id = shapeId;
      }

      const alignMatch = shapeContent.match(/<(?:hh:)?align[^>]*horizontal="([^"]*)"/i);
      if (alignMatch) {
        const alignVal = alignMatch[1].toUpperCase();
        if (alignVal === 'JUSTIFY') paraShape.align = 'Justify';
        else if (alignVal === 'CENTER') paraShape.align = 'Center';
        else if (alignVal === 'RIGHT') paraShape.align = 'Right';
        else if (alignVal === 'DISTRIBUTE') paraShape.align = 'Distribute';
        else if (alignVal === 'DISTRIBUTE_SPACE') paraShape.align = 'DistributeSpace';
        else paraShape.align = 'Left';
      }

      let lineSpaceMatch = shapeContent.match(/<(?:hh:)?lineSpacing[^>]*type="([^"]*)"[^>]*value="(\d+)"/i);
      if (!lineSpaceMatch) {
        lineSpaceMatch = shapeContent.match(/<(?:hh:)?lineSpacing[^>]*value="(\d+)"[^>]*type="([^"]*)"/i);
        if (lineSpaceMatch) {
          lineSpaceMatch = [lineSpaceMatch[0], lineSpaceMatch[2], lineSpaceMatch[1]] as RegExpMatchArray;
        }
      }
      if (lineSpaceMatch) {
        paraShape.lineSpacing = parseInt(lineSpaceMatch[2]);
        const typeMap: Record<string, string> = {
          'PERCENT': 'percent', 'FIXED': 'fixed', 'BETWEEN_LINES': 'betweenLines', 'AT_LEAST': 'atLeast'
        };
        paraShape.lineSpacingType = typeMap[lineSpaceMatch[1]?.toUpperCase()] || 'percent';
      }

      const caseMatch = shapeContent.match(/<hp:case[^>]*>([\s\S]*?)<\/hp:case>/i);
      const marginSource = caseMatch ? caseMatch[1] : shapeContent;

      const leftMatch = marginSource.match(/<(?:hc:)?left[^>]*value="(-?\d+)"/i);
      if (leftMatch) {
        paraShape.marginLeft = parseInt(leftMatch[1]) / 100;
      }

      const rightMatch = marginSource.match(/<(?:hc:)?right[^>]*value="(-?\d+)"/i);
      if (rightMatch) {
        paraShape.marginRight = parseInt(rightMatch[1]) / 100;
      }

      const prevMatch = marginSource.match(/<(?:hc:)?prev[^>]*value="(-?\d+)"/i);
      if (prevMatch) {
        paraShape.marginTop = parseInt(prevMatch[1]) / 100;
      }

      const nextMatch = marginSource.match(/<(?:hc:)?next[^>]*value="(-?\d+)"/i);
      if (nextMatch) {
        paraShape.marginBottom = parseInt(nextMatch[1]) / 100;
      }

      const intentMatch = marginSource.match(/<(?:hc:)?intent[^>]*value="(-?\d+)"/i);
      if (intentMatch) {
        paraShape.firstLineIndent = parseInt(intentMatch[1]) / 100;
      }

      const tabDefMatch = shapeContent.match(/tabPrIDRef="(\d+)"/);
      if (tabDefMatch) {
        paraShape.tabDefId = parseInt(tabDefMatch[1]);
      }

      const condenseMatch = shapeContent.match(/condense="(-?\d+)"/);
      if (condenseMatch) {
        paraShape.condense = parseInt(condenseMatch[1]);
      }

      const breakLatinMatch = shapeContent.match(/breakLatinWord="([^"]*)"/);
      if (breakLatinMatch) {
        const breakMap: Record<string, import('./types').BreakWordType> = {
          'KEEP_WORD': 'normal', 'HYPHENATION': 'hyphenation', 'BREAK_WORD': 'breakWord'
        };
        paraShape.breakLatinWord = breakMap[breakLatinMatch[1].toUpperCase()] || 'normal';
      }

      const breakNonLatinMatch = shapeContent.match(/breakNonLatinWord="([^"]*)"/);
      if (breakNonLatinMatch) {
        paraShape.breakNonLatinWord = breakNonLatinMatch[1] === '1' || breakNonLatinMatch[1] === 'true';
      }

      const snapToGridMatch = shapeContent.match(/snapToGrid="([^"]*)"/);
      if (snapToGridMatch) {
        paraShape.snapToGrid = snapToGridMatch[1] === '1' || snapToGridMatch[1] === 'true';
      }

      const suppressLineNumMatch = shapeContent.match(/suppressLineNumbers="([^"]*)"/);
      if (suppressLineNumMatch) {
        paraShape.suppressLineNumbers = suppressLineNumMatch[1] === '1' || suppressLineNumMatch[1] === 'true';
      }

      const headingMatch = shapeContent.match(/<(?:hh:)?heading[^>]*type="([^"]*)"[^>]*(?:level="(\d+)")?/i);
      if (headingMatch) {
        const headingMap: Record<string, import('./types').HeadingType> = {
          'NONE': 'none', 'OUTLINE': 'outline', 'NUMBER': 'number', 'BULLET': 'bullet'
        };
        paraShape.headingType = headingMap[headingMatch[1]?.toUpperCase()] || 'none';
        if (headingMatch[2]) {
          paraShape.headingLevel = parseInt(headingMatch[2]);
        }
      }

      const borderFillMatch = shapeContent.match(/borderFillIDRef="(\d+)"/);
      if (borderFillMatch) {
        paraShape.borderFillId = parseInt(borderFillMatch[1]);
      }

      const autoSpaceEAEngMatch = shapeContent.match(/autoSpaceEAsianEng="([^"]*)"/);
      if (autoSpaceEAEngMatch) {
        paraShape.autoSpaceEAsianEng = autoSpaceEAEngMatch[1] === '1' || autoSpaceEAEngMatch[1] === 'true';
      }

      const autoSpaceEANumMatch = shapeContent.match(/autoSpaceEAsianNum="([^"]*)"/);
      if (autoSpaceEANumMatch) {
        paraShape.autoSpaceEAsianNum = autoSpaceEANumMatch[1] === '1' || autoSpaceEANumMatch[1] === 'true';
      }

      const keepWithNextMatch = shapeContent.match(/keepWithNext="([^"]*)"/);
      if (keepWithNextMatch) {
        paraShape.keepWithNext = keepWithNextMatch[1] === '1' || keepWithNextMatch[1] === 'true';
      }

      const keepLinesMatch = shapeContent.match(/keepLines="([^"]*)"/);
      if (keepLinesMatch) {
        paraShape.keepLines = keepLinesMatch[1] === '1' || keepLinesMatch[1] === 'true';
      }

      const pageBreakBeforeMatch = shapeContent.match(/pageBreakBefore="([^"]*)"/);
      if (pageBreakBeforeMatch) {
        paraShape.pageBreakBefore = pageBreakBeforeMatch[1] === '1' || pageBreakBeforeMatch[1] === 'true';
      }

      const widowControlMatch = shapeContent.match(/widowOrphan="([^"]*)"/);
      if (widowControlMatch) {
        paraShape.widowControl = widowControlMatch[1] === '1' || widowControlMatch[1] === 'true';
      }

      this.styles.paraShapes.set(paraShape.id, paraShape);
      shapeId++;
    }
  }

  private static parseBorderFills(xml: string): void {
    const borderFillRegex = /<hh:borderFill\s+id="(\d+)"[^>]*>([\s\S]*?)<\/hh:borderFill>/gi;
    let match;

    while ((match = borderFillRegex.exec(xml)) !== null) {
      const id = parseInt(match[1]);
      const content = match[0];
      const tagAttrs = match[0].match(/<hh:borderFill[^>]*>/)?.[0] || '';

      const borderFill: import('./types').BorderFillStyle = { id };

      const parseBorder = (name: string) => {
        const regex = new RegExp(`<hh:${name}Border[^>]*type="([^"]*)"[^>]*width="([^"]*)"[^>]*color="([^"]*)"`, 'i');
        const borderMatch = content.match(regex);
        if (borderMatch) {
          const typeMap: { [key: string]: 'none' | 'solid' | 'dashed' | 'dotted' | 'double' } = {
            'NONE': 'none', 'SOLID': 'solid', 'DASHED': 'dashed', 'DASH': 'dashed', 'DOTTED': 'dotted', 'DOUBLE': 'double'
          };
          const widthStr = borderMatch[2];
          let widthPt = 0.5;
          const widthNumMatch = widthStr.match(/([\d.]+)\s*(mm|pt|cm)?/);
          if (widthNumMatch) {
            const num = parseFloat(widthNumMatch[1]);
            const unit = widthNumMatch[2]?.toLowerCase() || 'mm';
            if (unit === 'mm') widthPt = num * 2.83465;
            else if (unit === 'cm') widthPt = num * 28.3465;
            else widthPt = num;
          }
          return {
            style: typeMap[borderMatch[1].toUpperCase()] || 'solid',
            width: widthPt,
            color: borderMatch[3]
          };
        }
        return undefined;
      };

      borderFill.leftBorder = parseBorder('left');
      borderFill.rightBorder = parseBorder('right');
      borderFill.topBorder = parseBorder('top');
      borderFill.bottomBorder = parseBorder('bottom');
      borderFill.diagonalBorder = parseBorder('diagonal');
      borderFill.antiDiagonalBorder = parseBorder('antiDiagonal');

      const threeDMatch = tagAttrs.match(/threeD="([^"]*)"/);
      if (threeDMatch) {
        borderFill.threeD = threeDMatch[1] === '1' || threeDMatch[1] === 'true';
      }

      const shadowMatch = tagAttrs.match(/shadow="([^"]*)"/);
      if (shadowMatch) {
        borderFill.shadow = shadowMatch[1] === '1' || shadowMatch[1] === 'true';
      }

      const centerLineMatch = tagAttrs.match(/centerLine="([^"]*)"/);
      if (centerLineMatch) {
        borderFill.centerLine = centerLineMatch[1] === '1' || centerLineMatch[1] === 'true';
      }

      // Support both hh: and hc: namespace prefixes for fillBrush
      const fillBrushMatch = content.match(/<(?:hh|hc):fillBrush[^>]*>([\s\S]*?)<\/(?:hh|hc):fillBrush>/i);
      if (fillBrushMatch) {
        const fillContent = fillBrushMatch[1];

        // Support both hh: and hc: namespace for winBrush
        const windowColorMatch = fillContent.match(/<(?:hh|hc):winBrush[^>]*faceColor="([^"]*)"/i);
        if (windowColorMatch && windowColorMatch[1] !== 'none') {
          borderFill.fillColor = windowColorMatch[1];
          borderFill.fillType = 'color';
        }

        const gradationMatch = fillContent.match(/<(?:hh|hc):gradation[^>]*type="([^"]*)"[^>]*(?:angle="([^"]*)")?[^>]*(?:centerX="([^"]*)")?[^>]*(?:centerY="([^"]*)")?[^>]*(?:step="([^"]*)")?[^>]*>([\s\S]*?)<\/(?:hh|hc):gradation>/i);
        if (gradationMatch) {
          const typeMap: Record<string, import('./types').GradationType> = {
            'LINEAR': 'linear', 'RADIAL': 'radial', 'CONICAL': 'conical', 'SQUARE': 'square'
          };
          borderFill.fillType = 'gradation';
          borderFill.gradation = {
            type: typeMap[gradationMatch[1]?.toUpperCase()] || 'linear',
            colors: []
          };
          if (gradationMatch[2]) borderFill.gradation.angle = parseInt(gradationMatch[2]);
          if (gradationMatch[3]) borderFill.gradation.centerX = parseInt(gradationMatch[3]);
          if (gradationMatch[4]) borderFill.gradation.centerY = parseInt(gradationMatch[4]);
          if (gradationMatch[5]) borderFill.gradation.step = parseInt(gradationMatch[5]);

          const colorRegex = /<(?:hh|hc):color[^>]*value="([^"]*)"/gi;
          let colorMatch;
          while ((colorMatch = colorRegex.exec(gradationMatch[6])) !== null) {
            borderFill.gradation.colors.push(colorMatch[1]);
          }
        }

        const imgBrushMatch = fillContent.match(/<(?:hh|hc):imgBrush[^>]*mode="([^"]*)"[^>]*(?:alpha="([^"]*)")?[^>]*(?:binaryItemIDRef="([^"]*)")?/i);
        if (imgBrushMatch) {
          const modeMap: Record<string, import('./types').ImageFillMode> = {
            'TILE': 'tile', 'TILE_HORZ': 'tileHorz', 'TILE_VERT': 'tileVert',
            'TOTAL_FIT': 'totalFit', 'FIT': 'fit', 'CENTER': 'center',
            'ONCE_ABSOLUTE_SCALE': 'onceAbsoluteScale'
          };
          borderFill.fillType = 'image';
          borderFill.imageFill = {
            mode: modeMap[imgBrushMatch[1]?.toUpperCase()] || 'tile'
          };
          if (imgBrushMatch[2]) {
            borderFill.imageFill.alpha = parseInt(imgBrushMatch[2]) / 255;
          }
          if (imgBrushMatch[3]) {
            borderFill.imageFill.binaryItemId = imgBrushMatch[3];
          }
        }
      }

      if (!borderFill.fillColor) {
        const fillMatch = content.match(/faceColor="([^"]*)"/);
        if (fillMatch && fillMatch[1] !== 'none') {
          borderFill.fillColor = fillMatch[1];
          borderFill.fillType = 'color';
        }
      }

      this.styles.borderFills.set(id, borderFill);
    }
  }

  private static parseTabDefs(xml: string): void {
    const tabPrRegex = /<hh:tabPr\s+id="(\d+)"[^>]*>([\s\S]*?)<\/hh:tabPr>|<hh:tabPr\s+id="(\d+)"[^>]*\/>/gi;
    let match;

    while ((match = tabPrRegex.exec(xml)) !== null) {
      const id = parseInt(match[1] || match[3]);
      const content = match[0];
      
      const tabDef: import('./types').TabDef = { id, items: [] };

      const autoLeftMatch = content.match(/autoTabLeft="([^"]*)"/);
      if (autoLeftMatch) {
        tabDef.autoTabLeft = autoLeftMatch[1] === '1' || autoLeftMatch[1] === 'true';
      }

      const autoRightMatch = content.match(/autoTabRight="([^"]*)"/);
      if (autoRightMatch) {
        tabDef.autoTabRight = autoRightMatch[1] === '1' || autoRightMatch[1] === 'true';
      }

      const tabItemRegex = /<hh:tabItem[^>]*pos="(\d+)"[^>]*type="([^"]*)"[^>]*leader="([^"]*)"/gi;
      let itemMatch;
      while ((itemMatch = tabItemRegex.exec(content)) !== null) {
        const typeMap: Record<string, import('./types').TabType> = {
          'LEFT': 'left', 'RIGHT': 'right', 'CENTER': 'center', 'DECIMAL': 'decimal'
        };
        const leaderMap: Record<string, import('./types').TabLeader> = {
          'NONE': 'none', 'SOLID': 'solid', 'DASH': 'dash', 'DOT': 'dot',
          'DASH_DOT': 'dashDot', 'DASH_DOT_DOT': 'dashDotDot'
        };
        tabDef.items.push({
          pos: parseInt(itemMatch[1]) / 100,
          type: typeMap[itemMatch[2].toUpperCase()] || 'left',
          leader: leaderMap[itemMatch[3].toUpperCase()] || 'none'
        });
      }

      this.styles.tabDefs.set(id, tabDef);
    }
  }

  private static parseNumberings(xml: string): void {
    const numberingRegex = /<hh:numbering\s+id="(\d+)"[^>]*>([\s\S]*?)<\/hh:numbering>/gi;
    let match;

    while ((match = numberingRegex.exec(xml)) !== null) {
      const id = parseInt(match[1]);
      const content = match[0];

      const numberingDef: import('./types').NumberingDef = { id, paraHeads: [] };

      const startMatch = content.match(/\bstart="(\d+)"/);
      if (startMatch) {
        numberingDef.start = parseInt(startMatch[1]);
      }

      const paraHeadRegex = /<hh:paraHead[^>]*level="(\d+)"[^>]*numFormat="([^"]*)"[^>]*>([^<]*)<\/hh:paraHead>|<hh:paraHead[^>]*level="(\d+)"[^>]*numFormat="([^"]*)"[^>]*\/>/gi;
      let headMatch;
      while ((headMatch = paraHeadRegex.exec(content)) !== null) {
        const level = parseInt(headMatch[1] || headMatch[4]);
        const numFormatStr = headMatch[2] || headMatch[5];
        const text = headMatch[3] || '';

        const formatMap: Record<string, import('./types').NumFormat> = {
          'DIGIT': 'digit', 'ROMAN_CAPITAL': 'romanCapital', 'ROMAN_SMALL': 'romanSmall',
          'LATIN_CAPITAL': 'latinCapital', 'LATIN_SMALL': 'latinSmall',
          'HANGUL_SYLLABLE': 'hangulSyllable', 'HANGUL_JAMO': 'hangulJamo',
          'CIRCLED_DIGIT': 'circledDigit'
        };

        numberingDef.paraHeads.push({
          level,
          numFormat: formatMap[numFormatStr?.toUpperCase()] || 'digit',
          text: text || undefined
        });
      }

      this.styles.numberings.set(id, numberingDef);
    }
  }

  private static parseBullets(xml: string): void {
    const bulletRegex = /<hh:bullet\s+id="(\d+)"[^>]*>([\s\S]*?)<\/hh:bullet>|<hh:bullet\s+id="(\d+)"[^>]*\/>/gi;
    let match;

    while ((match = bulletRegex.exec(xml)) !== null) {
      const id = parseInt(match[1] || match[3]);
      const content = match[0];

      const bulletDef: import('./types').BulletDef = { id };

      const charMatch = content.match(/\bchar="([^"]*)"/);
      if (charMatch) {
        bulletDef.char = charMatch[1];
      }

      const useImageMatch = content.match(/useImage="([^"]*)"/);
      if (useImageMatch) {
        bulletDef.useImage = useImageMatch[1] === '1' || useImageMatch[1] === 'true';
      }

      this.styles.bullets.set(id, bulletDef);
    }
  }

  private static parseStyleDefs(xml: string): void {
    const styleRegex = /<hh:style\s+[^>]*id="(\d+)"[^>]*>([\s\S]*?)<\/hh:style>|<hh:style\s+[^>]*id="(\d+)"[^>]*\/>/gi;
    let match;

    while ((match = styleRegex.exec(xml)) !== null) {
      const id = parseInt(match[1] || match[3]);
      const content = match[0];

      const styleDef: import('./types').StyleDef = { id };

      const typeMatch = content.match(/\btype="([^"]*)"/);
      if (typeMatch) {
        styleDef.type = typeMatch[1].toLowerCase() === 'char' ? 'char' : 'para';
      }

      const nameMatch = content.match(/\bname="([^"]*)"/);
      if (nameMatch) {
        styleDef.name = nameMatch[1];
      }

      const engNameMatch = content.match(/engName="([^"]*)"/);
      if (engNameMatch) {
        styleDef.engName = engNameMatch[1];
      }

      const paraPrMatch = content.match(/paraPrIDRef="(\d+)"/);
      if (paraPrMatch) {
        styleDef.paraPrIdRef = parseInt(paraPrMatch[1]);
      }

      const charPrMatch = content.match(/charPrIDRef="(\d+)"/);
      if (charPrMatch) {
        styleDef.charPrIdRef = parseInt(charPrMatch[1]);
      }

      const nextStyleMatch = content.match(/nextStyleIDRef="(\d+)"/);
      if (nextStyleMatch) {
        styleDef.nextStyleIdRef = parseInt(nextStyleMatch[1]);
      }

      this.styles.styles.set(id, styleDef);
    }
  }

  private static async parseImages(zip: JSZip, content: HwpxContent): Promise<void> {
    const binDataFolder = zip.folder('BinData');
    if (!binDataFolder) return;

    const imageFiles = Object.keys(zip.files).filter(
      (f) => f.startsWith('BinData/') && !f.endsWith('/')
    );

    for (const imagePath of imageFiles) {
      const file = zip.file(imagePath);
      if (!file) continue;

      const data = await file.async('base64');
      const fileName = imagePath.split('/').pop() || '';
      const ext = fileName.split('.').pop()?.toLowerCase() || '';
      
      let mimeType = 'image/png';
      if (ext === 'jpg' || ext === 'jpeg') mimeType = 'image/jpeg';
      else if (ext === 'gif') mimeType = 'image/gif';
      else if (ext === 'bmp') mimeType = 'image/bmp';
      else if (ext === 'svg') mimeType = 'image/svg+xml';

      const imageId = fileName.replace(/\.[^.]+$/, '');
      content.images.set(imageId, {
        id: imageId,
        binaryId: imagePath,
        width: 0,
        height: 0,
        data: `data:${mimeType};base64,${data}`,
        mimeType,
      });
    }
  }

  private static parseSection(xml: string, content: HwpxContent): HwpxSection {
    const section: HwpxSection = {
      elements: [],
      pageSettings: this.parsePageSettings(xml),
      sectionProperties: this.parseSectionProperties(xml),
      memos: [], // Store memo data for sidebar display
    };

    // Extract MEMO data before removing from XML
    // Pattern: fieldBegin...content.../fieldBegin followed by linked text then fieldEnd
    const memoFullRegex = /<hp:fieldBegin[^>]*id="([^"]*)"[^>]*type="MEMO"[^>]*>([\s\S]*?)<\/hp:fieldBegin>([\s\S]*?)<hp:ctrl>\s*<hp:fieldEnd/gi;
    let memoMatch;
    while ((memoMatch = memoFullRegex.exec(xml)) !== null) {
      const memoId = memoMatch[1];
      const memoContent = memoMatch[2];
      const linkedSection = memoMatch[3];

      const memo: import('./types').Memo = {
        id: memoId,
        author: memoContent.match(/<hp:stringParam[^>]*name="Author"[^>]*>([^<]*)<\/hp:stringParam>/i)?.[1] || 'Unknown',
        date: memoContent.match(/<hp:stringParam[^>]*name="CreateDateTime"[^>]*>([^<]*)<\/hp:stringParam>/i)?.[1] || '',
        content: [],
      };

      // Extract all text from subList paragraphs (memo content)
      const textMatches = memoContent.matchAll(/<hp:t[^>]*>([^<]*)<\/hp:t>/gi);
      for (const textMatch of textMatches) {
        if (textMatch[1]) memo.content.push(textMatch[1]);
      }

      // Extract linked text (text between fieldBegin end and fieldEnd)
      const linkedTexts = linkedSection.match(/<hp:t[^>]*>([^<]*)<\/hp:t>/gi);
      if (linkedTexts) {
        const texts = linkedTexts.map(t => t.replace(/<[^>]+>/g, '')).filter(t => t);
        memo.linkedText = texts.join('');
      }

      section.memos!.push(memo);
    }

    // Remove MEMO fieldBegin content to prevent memo text from appearing as document content
    // This removes the entire fieldBegin tag including subList with memo paragraphs
    let cleanedXml = xml.replace(/<hp:fieldBegin[^>]*type="MEMO"[^>]*>[\s\S]*?<\/hp:fieldBegin>/gi, '');

    // Extract footnote reference positions BEFORE removing footnote content
    // This allows us to add footnote markers to the correct paragraphs later
    const footnoteRefPositions: { position: number; number: number; type: 'footnote' | 'endnote' }[] = [];
    const fnPosRegex = /<hp:footNote\b[^>]*number="(\d+)"[^>]*>/gi;
    let fnPosMatch;
    while ((fnPosMatch = fnPosRegex.exec(cleanedXml)) !== null) {
      footnoteRefPositions.push({
        position: fnPosMatch.index,
        number: parseInt(fnPosMatch[1]),
        type: 'footnote'
      });
    }
    const enPosRegex = /<hp:endNote\b[^>]*number="(\d+)"[^>]*>/gi;
    let enPosMatch;
    while ((enPosMatch = enPosRegex.exec(cleanedXml)) !== null) {
      footnoteRefPositions.push({
        position: enPosMatch.index,
        number: parseInt(enPosMatch[1]),
        type: 'endnote'
      });
    }

    // Remove footnote/endnote content to prevent footnote text from appearing in document body
    // Only remove the content inside, but track where they were for reference markers
    // Use a more precise regex to only remove the footNote element and its content
    cleanedXml = cleanedXml.replace(/<hp:footNote\b[^>]*>[\s\S]*?<\/hp:footNote>/gi, '');
    cleanedXml = cleanedXml.replace(/<hp:endNote\b[^>]*>[\s\S]*?<\/hp:endNote>/gi, '');

    const elements: { index: number; type: string; xml: string; parentLinesegs?: import('./types').LineSeg[] }[] = [];

    // Extract ALL paragraphs first to find parent paragraphs for tables
    const paragraphs = this.extractAllParagraphs(cleanedXml);

    // Extract all tables from cleaned XML (without MEMOs and footnotes) to maintain consistent indices
    const tables = this.extractBalancedTags(cleanedXml, 'hp:tbl');
    const tableRanges: { start: number; end: number }[] = [];
    for (const tableXml of tables) {
      const tableIndex = cleanedXml.indexOf(tableXml);

      // Find parent paragraph that contains this table
      let parentLinesegs: import('./types').LineSeg[] | undefined;
      for (const para of paragraphs) {
        if (tableIndex >= para.start && tableIndex < para.start + para.xml.length) {
          // This paragraph contains the table, extract its lineseg
          // The paragraph's own lineseg is at the END (after nested content like table cells)
          // So we use the LAST linesegarray in the paragraph
          const linesegArrayRegex = /<hp:linesegarray>([\s\S]*?)<\/hp:linesegarray>/g;
          let lastLinesegArray: string | null = null;
          let arrayMatch;
          while ((arrayMatch = linesegArrayRegex.exec(para.xml)) !== null) {
            lastLinesegArray = arrayMatch[1];
          }

          if (lastLinesegArray) {
            const linesegRegex = /<hp:lineseg[^>]*vertpos="(\d+)"[^>]*vertsize="(\d+)"[^>]*textheight="(\d+)"[^>]*baseline="(\d+)"[^>]*spacing="(\d+)"/g;
            let linesegMatch;
            const linesegs: import('./types').LineSeg[] = [];
            while ((linesegMatch = linesegRegex.exec(lastLinesegArray)) !== null) {
              linesegs.push({
                vertpos: parseInt(linesegMatch[1]) / 100,
                vertsize: parseInt(linesegMatch[2]) / 100,
                textheight: parseInt(linesegMatch[3]) / 100,
                baseline: parseInt(linesegMatch[4]) / 100,
                spacing: parseInt(linesegMatch[5]) / 100,
              });
            }
            if (linesegs.length > 0) {
              parentLinesegs = linesegs;
            }
          }
          break;
        }
      }

      elements.push({ index: tableIndex, type: 'tbl', xml: tableXml, parentLinesegs });
      tableRanges.push({ start: tableIndex, end: tableIndex + tableXml.length });
    }

    // Add paragraphs that are not inside tables
    // For paragraphs that contain tables, still parse the text content (excluding the table XML)
    for (const para of paragraphs) {
      const isInsideTable = tableRanges.some(
        range => para.start > range.start && para.start < range.end
      );
      const containsTable = tableRanges.some(
        range => range.start >= para.start && range.end <= para.end
      );
      if (!isInsideTable) {
        if (containsTable) {
          // Paragraph contains a table - remove the table XML and parse the remaining content
          let paraXmlWithoutTable = para.xml;
          for (const range of tableRanges) {
            if (range.start >= para.start && range.end <= para.start + para.xml.length) {
              // Find and remove the table from paragraph XML
              const tableStartInPara = range.start - para.start;
              const tableEndInPara = range.end - para.start;
              const tableXmlInPara = para.xml.substring(tableStartInPara, tableEndInPara);
              paraXmlWithoutTable = paraXmlWithoutTable.replace(tableXmlInPara, '');
            }
          }
          // Only add if there's remaining content besides lineseg
          const hasTextContent = /<hp:t\b[^>]*>/.test(paraXmlWithoutTable);
          if (hasTextContent) {
            elements.push({ index: para.start, type: 'p', xml: paraXmlWithoutTable });
          }
        } else {
          elements.push({ index: para.start, type: 'p', xml: para.xml });
        }
      }
    }

    const lineRegex = /<hp:line\b[^>]*(?:\/>|>[\s\S]*?<\/hp:line>)/g;
    let lineMatch;
    while ((lineMatch = lineRegex.exec(xml)) !== null) {
      elements.push({ index: lineMatch.index, type: 'line', xml: lineMatch[0] });
    }

    const rectRegex = /<hp:rect\b[^>]*(?:\/>|>[\s\S]*?<\/hp:rect>)/g;
    let rectMatch;
    while ((rectMatch = rectRegex.exec(xml)) !== null) {
      elements.push({ index: rectMatch.index, type: 'rect', xml: rectMatch[0] });
    }

    const ellipseRegex = /<hp:ellipse\b[^>]*(?:\/>|>[\s\S]*?<\/hp:ellipse>)/g;
    let ellipseMatch;
    while ((ellipseMatch = ellipseRegex.exec(xml)) !== null) {
      elements.push({ index: ellipseMatch.index, type: 'ellipse', xml: ellipseMatch[0] });
    }

    // Arc (호)
    const arcRegex = /<hp:arc\b[^>]*(?:\/>|>[\s\S]*?<\/hp:arc>)/g;
    let arcMatch;
    while ((arcMatch = arcRegex.exec(xml)) !== null) {
      elements.push({ index: arcMatch.index, type: 'arc', xml: arcMatch[0] });
    }

    // Polygon (다각형)
    const polygonRegex = /<hp:polygon\b[^>]*(?:\/>|>[\s\S]*?<\/hp:polygon>)/g;
    let polygonMatch;
    while ((polygonMatch = polygonRegex.exec(xml)) !== null) {
      elements.push({ index: polygonMatch.index, type: 'polygon', xml: polygonMatch[0] });
    }

    // Curve (곡선)
    const curveRegex = /<hp:curve\b[^>]*(?:\/>|>[\s\S]*?<\/hp:curve>)/g;
    let curveMatch;
    while ((curveMatch = curveRegex.exec(xml)) !== null) {
      elements.push({ index: curveMatch.index, type: 'curve', xml: curveMatch[0] });
    }

    // ConnectLine (연결선)
    const connectLineRegex = /<hp:connectLine\b[^>]*(?:\/>|>[\s\S]*?<\/hp:connectLine>)/g;
    let connectLineMatch;
    while ((connectLineMatch = connectLineRegex.exec(xml)) !== null) {
      elements.push({ index: connectLineMatch.index, type: 'connectline', xml: connectLineMatch[0] });
    }

    // Container (묶음객체)
    const containerRegex = /<hp:container\b[^>]*>[\s\S]*?<\/hp:container>/g;
    let containerMatch;
    while ((containerMatch = containerRegex.exec(xml)) !== null) {
      elements.push({ index: containerMatch.index, type: 'container', xml: containerMatch[0] });
    }

    // OLE
    const oleRegex = /<hp:ole\b[^>]*(?:\/>|>[\s\S]*?<\/hp:ole>)/g;
    let oleMatch;
    while ((oleMatch = oleRegex.exec(xml)) !== null) {
      elements.push({ index: oleMatch.index, type: 'ole', xml: oleMatch[0] });
    }

    // Equation (수식)
    const equationRegex = /<hp:equation\b[^>]*(?:\/>|>[\s\S]*?<\/hp:equation>)/g;
    let equationMatch;
    while ((equationMatch = equationRegex.exec(xml)) !== null) {
      elements.push({ index: equationMatch.index, type: 'equation', xml: equationMatch[0] });
    }

    // TextArt (글맵시)
    const textArtRegex = /<hp:textArt\b[^>]*(?:\/>|>[\s\S]*?<\/hp:textArt>)/g;
    let textArtMatch;
    while ((textArtMatch = textArtRegex.exec(xml)) !== null) {
      elements.push({ index: textArtMatch.index, type: 'textart', xml: textArtMatch[0] });
    }

    // UnknownObject
    const unknownObjRegex = /<hp:unknownObj\b[^>]*(?:\/>|>[\s\S]*?<\/hp:unknownObj>)/g;
    let unknownObjMatch;
    while ((unknownObjMatch = unknownObjRegex.exec(xml)) !== null) {
      elements.push({ index: unknownObjMatch.index, type: 'unknownobject', xml: unknownObjMatch[0] });
    }

    // Form Objects
    const buttonRegex = /<hp:button\b[^>]*(?:\/>|>[\s\S]*?<\/hp:button>)/g;
    let buttonMatch;
    while ((buttonMatch = buttonRegex.exec(xml)) !== null) {
      elements.push({ index: buttonMatch.index, type: 'button', xml: buttonMatch[0] });
    }

    const radioButtonRegex = /<hp:radioButton\b[^>]*(?:\/>|>[\s\S]*?<\/hp:radioButton>)/g;
    let radioButtonMatch;
    while ((radioButtonMatch = radioButtonRegex.exec(xml)) !== null) {
      elements.push({ index: radioButtonMatch.index, type: 'radiobutton', xml: radioButtonMatch[0] });
    }

    const checkButtonRegex = /<hp:checkButton\b[^>]*(?:\/>|>[\s\S]*?<\/hp:checkButton>)/g;
    let checkButtonMatch;
    while ((checkButtonMatch = checkButtonRegex.exec(xml)) !== null) {
      elements.push({ index: checkButtonMatch.index, type: 'checkbutton', xml: checkButtonMatch[0] });
    }

    const comboBoxRegex = /<hp:comboBox\b[^>]*(?:\/>|>[\s\S]*?<\/hp:comboBox>)/g;
    let comboBoxMatch;
    while ((comboBoxMatch = comboBoxRegex.exec(xml)) !== null) {
      elements.push({ index: comboBoxMatch.index, type: 'combobox', xml: comboBoxMatch[0] });
    }

    const editRegex = /<hp:edit\b[^>]*(?:\/>|>[\s\S]*?<\/hp:edit>)/g;
    let editMatch;
    while ((editMatch = editRegex.exec(xml)) !== null) {
      elements.push({ index: editMatch.index, type: 'edit', xml: editMatch[0] });
    }

    const listBoxRegex = /<hp:listBox\b[^>]*(?:\/>|>[\s\S]*?<\/hp:listBox>)/g;
    let listBoxMatch;
    while ((listBoxMatch = listBoxRegex.exec(xml)) !== null) {
      elements.push({ index: listBoxMatch.index, type: 'listbox', xml: listBoxMatch[0] });
    }

    const scrollBarRegex = /<hp:scrollBar\b[^>]*(?:\/>|>[\s\S]*?<\/hp:scrollBar>)/g;
    let scrollBarMatch;
    while ((scrollBarMatch = scrollBarRegex.exec(xml)) !== null) {
      elements.push({ index: scrollBarMatch.index, type: 'scrollbar', xml: scrollBarMatch[0] });
    }

    const picRegex = /<hp:pic\b[^>]*>[\s\S]*?<\/hp:pic>/g;
    let picMatch;
    while ((picMatch = picRegex.exec(xml)) !== null) {
      elements.push({ index: picMatch.index, type: 'pic', xml: picMatch[0] });
    }

    // Video element
    const videoRegex = /<hp:video\b[^>]*(?:\/>|>[\s\S]*?<\/hp:video>)/g;
    let videoMatch;
    while ((videoMatch = videoRegex.exec(xml)) !== null) {
      elements.push({ index: videoMatch.index, type: 'video', xml: videoMatch[0] });
    }

    // Chart element
    const chartRegex = /<hp:chart\b[^>]*(?:\/>|>[\s\S]*?<\/hp:chart>)/g;
    let chartMatch;
    while ((chartMatch = chartRegex.exec(xml)) !== null) {
      elements.push({ index: chartMatch.index, type: 'chart', xml: chartMatch[0] });
    }

    elements.sort((a, b) => a.index - b.index);

    for (const el of elements) {
      if (el.type === 'p') {
        const paragraph = this.parseParagraph(el.xml);

        // Check if this paragraph should have a footnote reference
        // (footnote was in original XML but removed from cleanedXml)
        for (const fnRef of footnoteRefPositions) {
          // Adjust position check - the footnote was within the original paragraph range
          // Since we removed footnotes, positions shift, but we can check if the footnote
          // position was within the paragraph's approximate range
          if (fnRef.position >= el.index && fnRef.position < el.index + el.xml.length + 500) {
            // Add footnote reference marker to the last run
            paragraph.runs.push({
              text: `${fnRef.number})`,
              footnoteRef: fnRef.type === 'footnote' ? fnRef.number : undefined,
              endnoteRef: fnRef.type === 'endnote' ? fnRef.number : undefined,
              charStyle: { superscript: true, fontSize: 7 },
            });
            // Remove this footnote from the list so it's not added again
            const idx = footnoteRefPositions.indexOf(fnRef);
            if (idx > -1) footnoteRefPositions.splice(idx, 1);
            break;
          }
        }

        section.elements.push({ type: 'paragraph', data: paragraph });
      } else if (el.type === 'tbl') {
        const table = this.parseTable(el.xml);
        // Add parent paragraph's lineseg info to table for page break detection
        if (el.parentLinesegs && el.parentLinesegs.length > 0) {
          table.linesegs = el.parentLinesegs;
        }
        section.elements.push({ type: 'table', data: table });
      } else if (el.type === 'pic') {
        const image = this.parseImageElement(el.xml, content);
        if (image) {
          section.elements.push({ type: 'image', data: image });
        }
      } else if (el.type === 'line') {
        const line = this.parseLine(el.xml);
        section.elements.push({ type: 'line', data: line });
      } else if (el.type === 'rect') {
        const rect = this.parseRect(el.xml);
        section.elements.push({ type: 'rect', data: rect });
      } else if (el.type === 'ellipse') {
        const ellipse = this.parseEllipse(el.xml);
        section.elements.push({ type: 'ellipse', data: ellipse });
      } else if (el.type === 'arc') {
        const arc = this.parseArc(el.xml);
        section.elements.push({ type: 'arc', data: arc });
      } else if (el.type === 'polygon') {
        const polygon = this.parsePolygon(el.xml);
        section.elements.push({ type: 'polygon', data: polygon });
      } else if (el.type === 'curve') {
        const curve = this.parseCurve(el.xml);
        section.elements.push({ type: 'curve', data: curve });
      } else if (el.type === 'connectline') {
        const connectLine = this.parseConnectLine(el.xml);
        section.elements.push({ type: 'connectline', data: connectLine });
      } else if (el.type === 'container') {
        const container = this.parseContainer(el.xml, content);
        section.elements.push({ type: 'container', data: container });
      } else if (el.type === 'ole') {
        const ole = this.parseOle(el.xml);
        section.elements.push({ type: 'ole', data: ole });
      } else if (el.type === 'equation') {
        const equation = this.parseEquation(el.xml);
        section.elements.push({ type: 'equation', data: equation });
      } else if (el.type === 'textart') {
        const textArt = this.parseTextArt(el.xml);
        section.elements.push({ type: 'textart', data: textArt });
      } else if (el.type === 'unknownobject') {
        const unknownObj = this.parseUnknownObject(el.xml);
        section.elements.push({ type: 'unknownobject', data: unknownObj });
      } else if (el.type === 'button') {
        const button = this.parseButton(el.xml);
        section.elements.push({ type: 'button', data: button });
      } else if (el.type === 'radiobutton') {
        const radioButton = this.parseRadioButton(el.xml);
        section.elements.push({ type: 'radiobutton', data: radioButton });
      } else if (el.type === 'checkbutton') {
        const checkButton = this.parseCheckButton(el.xml);
        section.elements.push({ type: 'checkbutton', data: checkButton });
      } else if (el.type === 'combobox') {
        const comboBox = this.parseComboBox(el.xml);
        section.elements.push({ type: 'combobox', data: comboBox });
      } else if (el.type === 'edit') {
        const edit = this.parseEdit(el.xml);
        section.elements.push({ type: 'edit', data: edit });
      } else if (el.type === 'listbox') {
        const listBox = this.parseListBox(el.xml);
        section.elements.push({ type: 'listbox', data: listBox });
      } else if (el.type === 'scrollbar') {
        const scrollBar = this.parseScrollBar(el.xml);
        section.elements.push({ type: 'scrollbar', data: scrollBar });
      }
    }

    this.parseHorizontalRules(xml, section);

    if (section.elements.length === 0) {
      const paragraphs = this.parseParagraphsSimple(xml);
      for (const p of paragraphs) {
        section.elements.push({ type: 'paragraph', data: p });
      }
    }

    section.header = this.parseHeaderFooter(xml, 'header');
    section.footer = this.parseHeaderFooter(xml, 'footer');

    this.parseFootnotes(xml, content);
    this.parseEndnotes(xml, content);
    this.parseHiddenComments(xml, section);

    return section;
  }

  private static parseEndnotes(xml: string, content: HwpxContent): void {
    const endnoteRegex = /<hp:endnote[^>]*>([\s\S]*?)<\/hp:endnote>/gi;
    let match;
    let endnoteIndex = 0;

    while ((match = endnoteRegex.exec(xml)) !== null) {
      const endnoteContent = match[0];
      const paragraphs: HwpxParagraph[] = [];

      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(endnoteContent)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }

      if (paragraphs.length > 0) {
        content.endnotes.push({
          id: `endnote_${endnoteIndex++}`,
          paragraphs
        });
      }
    }
  }

  private static parseHiddenComments(xml: string, section: HwpxSection): void {
    const hiddenCommentRegex = /<hp:hiddenComment[^>]*>([\s\S]*?)<\/hp:hiddenComment>/gi;
    let match;

    while ((match = hiddenCommentRegex.exec(xml)) !== null) {
      const commentContent = match[0];
      const paragraphs: HwpxParagraph[] = [];

      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(commentContent)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }

      // Hidden comments are not rendered but stored for reference
      // They could be stored in section or in a separate collection
    }
  }

  private static parseHeaderFooter(xml: string, type: 'header' | 'footer'): import('./types').HeaderFooter | undefined {
    const tagName = type === 'header' ? 'hp:header' : 'hp:footer';
    const regex = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)</${tagName}>`, 'gi');
    const match = xml.match(regex);
    
    if (!match) return undefined;

    const content = match[0];
    const paragraphs: HwpxParagraph[] = [];
    
    const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
    let paraMatch;
    while ((paraMatch = paraRegex.exec(content)) !== null) {
      paragraphs.push(this.parseParagraph(paraMatch[0]));
    }

    if (paragraphs.length === 0) return undefined;

    return { paragraphs };
  }

  private static parseFootnotes(xml: string, content: HwpxContent): void {
    // Match hp:footNote (note the capital N in actual HWPX files)
    const footnoteRegex = /<hp:footNote\b[^>]*>([\s\S]*?)<\/hp:footNote>/gi;
    let match;
    let footnoteIndex = 0;

    while ((match = footnoteRegex.exec(xml)) !== null) {
      const footnoteContent = match[0];
      const paragraphs: HwpxParagraph[] = [];

      // Extract footnote number from attribute
      const numberMatch = footnoteContent.match(/number="(\d+)"/);
      const footnoteNumber = numberMatch ? parseInt(numberMatch[1]) : footnoteIndex + 1;

      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(footnoteContent)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }

      if (paragraphs.length > 0) {
        content.footnotes.push({
          id: `footnote_${footnoteIndex}`,
          number: footnoteNumber,
          type: 'footnote',
          paragraphs,
        });
        footnoteIndex++;
      }
    }

    // Match hp:endNote (note the capital N in actual HWPX files)
    const endnoteRegex = /<hp:endNote\b[^>]*>([\s\S]*?)<\/hp:endNote>/gi;
    while ((match = endnoteRegex.exec(xml)) !== null) {
      const endnoteContent = match[0];
      const paragraphs: HwpxParagraph[] = [];

      const numberMatch = endnoteContent.match(/number="(\d+)"/);
      const endnoteNumber = numberMatch ? parseInt(numberMatch[1]) : footnoteIndex + 1;

      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(endnoteContent)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }

      if (paragraphs.length > 0) {
        content.footnotes.push({
          id: `endnote_${footnoteIndex}`,
          number: endnoteNumber,
          type: 'endnote',
          paragraphs,
        });
        footnoteIndex++;
      }
    }
  }

  private static parsePageSettings(xml: string): PageSettings {
    const defaults: PageSettings = {
      width: 595,
      height: 842,
      marginTop: 56.7,
      marginBottom: 56.7,
      marginLeft: 56.7,
      marginRight: 56.7,
      orientation: 'portrait',
    };

    const pagePrMatch = xml.match(/<hp:pagePr[^>]*>([\s\S]*?)<\/hp:pagePr>/);
    if (pagePrMatch) {
      const pagePr = pagePrMatch[0];
      
      const widthMatch = pagePr.match(/width="(\d+)"/);
      const heightMatch = pagePr.match(/height="(\d+)"/);
      if (widthMatch) defaults.width = parseInt(widthMatch[1]) / 100;
      if (heightMatch) defaults.height = parseInt(heightMatch[1]) / 100;

      const landscapeMatch = pagePr.match(/landscape="([^"]*)"/);
      if (landscapeMatch) {
        const val = landscapeMatch[1].toUpperCase();
        if (val === 'WIDELY' || val === '1' || val === 'TRUE' || val === 'LANDSCAPE') {
          defaults.orientation = 'landscape';
        }
      }

      const marginMatch = pagePr.match(/<hp:margin[^>]*left="(\d+)"[^>]*right="(\d+)"[^>]*top="(\d+)"[^>]*bottom="(\d+)"/);
      if (marginMatch) {
        defaults.marginLeft = parseInt(marginMatch[1]) / 100;
        defaults.marginRight = parseInt(marginMatch[2]) / 100;
        defaults.marginTop = parseInt(marginMatch[3]) / 100;
        defaults.marginBottom = parseInt(marginMatch[4]) / 100;
      }

      // Parse header and footer margins
      const marginTag = pagePr.match(/<hp:margin[^>]*>/);
      if (marginTag) {
        const headerMatch = marginTag[0].match(/header="(\d+)"/);
        const footerMatch = marginTag[0].match(/footer="(\d+)"/);
        const gutterMatch = marginTag[0].match(/gutter="(\d+)"/);
        if (headerMatch) defaults.headerMargin = parseInt(headerMatch[1]) / 100;
        if (footerMatch) defaults.footerMargin = parseInt(footerMatch[1]) / 100;
        if (gutterMatch) defaults.gutterMargin = parseInt(gutterMatch[1]) / 100;
      }
    }

    const pageDefMatch = xml.match(/<hp:pageDef[^>]*>/);
    if (pageDefMatch && defaults.width === 595) {
      const pageDef = pageDefMatch[0];

      const widthMatch = pageDef.match(/width="(\d+)"/);
      const heightMatch = pageDef.match(/height="(\d+)"/);

      if (widthMatch) defaults.width = parseInt(widthMatch[1]) / 100;
      if (heightMatch) defaults.height = parseInt(heightMatch[1]) / 100;

      const landscapeMatch = pageDef.match(/landscape="([^"]*)"/);
      if (landscapeMatch?.[1] === '1' || landscapeMatch?.[1] === 'true') {
        defaults.orientation = 'landscape';
      }
    }

    return defaults;
  }

  private static parseSectionProperties(xml: string): import('./types').SectionProperties | undefined {
    const secPrMatch = xml.match(/<hp:secPr[^>]*>([\s\S]*?)<\/hp:secPr>/);
    if (!secPrMatch) return undefined;

    const content = secPrMatch[0];
    const props: import('./types').SectionProperties = {};

    const textDirMatch = content.match(/textDirection="([^"]*)"/);
    if (textDirMatch) {
      props.textDirection = textDirMatch[1].toLowerCase() === 'vertical' ? 'vertical' : 'horizontal';
    }

    const spaceColsMatch = content.match(/spaceColumns="(\d+)"/);
    if (spaceColsMatch) {
      props.spaceColumns = parseInt(spaceColsMatch[1]) / 100;
    }

    const tabStopMatch = content.match(/tabStop="(\d+)"/);
    if (tabStopMatch) {
      props.tabStop = parseInt(tabStopMatch[1]) / 100;
    }

    const masterPageCntMatch = content.match(/masterPageCnt="(\d+)"/);
    if (masterPageCntMatch) {
      props.masterPageCnt = parseInt(masterPageCntMatch[1]);
    }

    const gridMatch = content.match(/<hp:grid[^>]*lineGrid="(\d+)"[^>]*charGrid="(\d+)"/);
    if (gridMatch) {
      props.grid = {
        lineGrid: parseInt(gridMatch[1]),
        charGrid: parseInt(gridMatch[2])
      };
    }

    const startNumMatch = content.match(/<hp:startNum[^>]*pageStartsOn="([^"]*)"[^>]*page="(\d+)"/);
    if (startNumMatch) {
      props.startNum = {
        pageStartsOn: startNumMatch[1].toLowerCase() as 'both' | 'even' | 'odd',
        page: parseInt(startNumMatch[2])
      };
    }

    const visMatch = content.match(/<hp:visibility[^>]*/);
    if (visMatch) {
      const vis = visMatch[0];
      props.visibility = {
        hideFirstHeader: vis.includes('hideFirstHeader="1"') || vis.includes('hideFirstHeader="true"'),
        hideFirstFooter: vis.includes('hideFirstFooter="1"') || vis.includes('hideFirstFooter="true"'),
        hideFirstMasterPage: vis.includes('hideFirstMasterPage="1"') || vis.includes('hideFirstMasterPage="true"'),
        hideFirstPageNum: vis.includes('hideFirstPageNum="1"') || vis.includes('hideFirstPageNum="true"'),
        showLineNumber: vis.includes('showLineNumber="1"') || vis.includes('showLineNumber="true"')
      };
      const borderMatch = vis.match(/border="([^"]*)"/);
      if (borderMatch) {
        const borderMap: Record<string, typeof props.visibility.border> = {
          'SHOW_ALL': 'showAll', 'HIDE_ALL': 'hideAll',
          'SHOW_FIRST_PAGE_ONLY': 'showFirstPageOnly', 'SHOW_ALL_BUT_FIRST_PAGE': 'showAllButFirstPage'
        };
        props.visibility.border = borderMap[borderMatch[1].toUpperCase()] || 'showAll';
      }
    }

    const pageBorderFills: import('./types').SectionProperties['pageBorderFill'] = [];
    const pbfRegex = /<hp:pageBorderFill[^>]*type="([^"]*)"[^>]*borderFillIDRef="(\d+)"[^>]*>([\s\S]*?)<\/hp:pageBorderFill>/gi;
    let pbfMatch;
    while ((pbfMatch = pbfRegex.exec(content)) !== null) {
      const pbf: NonNullable<import('./types').SectionProperties['pageBorderFill']>[0] = {
        type: pbfMatch[1].toLowerCase() as 'both' | 'even' | 'odd',
        borderFillIdRef: parseInt(pbfMatch[2])
      };
      
      const offsetMatch = pbfMatch[3].match(/<hp:offset[^>]*left="(\d+)"[^>]*right="(\d+)"[^>]*top="(\d+)"[^>]*bottom="(\d+)"/);
      if (offsetMatch) {
        pbf.offset = {
          left: parseInt(offsetMatch[1]) / 100,
          right: parseInt(offsetMatch[2]) / 100,
          top: parseInt(offsetMatch[3]) / 100,
          bottom: parseInt(offsetMatch[4]) / 100
        };
      }
      
      pageBorderFills.push(pbf);
    }
    if (pageBorderFills.length > 0) {
      props.pageBorderFill = pageBorderFills;
    }

    // Parse MasterPage
    const masterPages = this.parseMasterPages(content);
    if (masterPages && masterPages.length > 0) {
      props.masterPage = masterPages;
    }

    return props;
  }

  private static parseMasterPages(xml: string): import('./types').MasterPage[] | undefined {
    const masterPages: import('./types').MasterPage[] = [];
    const masterPageRegex = /<hp:masterPage[^>]*>([\s\S]*?)<\/hp:masterPage>/gi;
    let match;

    while ((match = masterPageRegex.exec(xml)) !== null) {
      const content = match[0];
      const masterPage: import('./types').MasterPage = {};

      const typeMatch = content.match(/type="([^"]*)"/);
      if (typeMatch) masterPage.type = typeMatch[1].toLowerCase() as 'both' | 'even' | 'odd';

      const textWidthMatch = content.match(/textWidth="(\d+)"/);
      if (textWidthMatch) masterPage.textWidth = parseInt(textWidthMatch[1]);

      const textHeightMatch = content.match(/textHeight="(\d+)"/);
      if (textHeightMatch) masterPage.textHeight = parseInt(textHeightMatch[1]);

      const hasTextRefMatch = content.match(/hasTextRef="(true|false|1|0)"/i);
      if (hasTextRefMatch) masterPage.hasTextRef = hasTextRefMatch[1] === 'true' || hasTextRefMatch[1] === '1';

      const hasNumRefMatch = content.match(/hasNumRef="(true|false|1|0)"/i);
      if (hasNumRefMatch) masterPage.hasNumRef = hasNumRefMatch[1] === 'true' || hasNumRefMatch[1] === '1';

      // Parse paragraphs in master page
      const paragraphs: HwpxParagraph[] = [];
      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(content)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }
      if (paragraphs.length > 0) {
        masterPage.paragraphs = paragraphs;
      }

      masterPages.push(masterPage);
    }

    // Parse extended master pages
    const extMasterPageRegex = /<hp:extMasterPage[^>]*>([\s\S]*?)<\/hp:extMasterPage>/gi;
    while ((match = extMasterPageRegex.exec(xml)) !== null) {
      const content = match[0];
      const masterPage: import('./types').MasterPage = { isExtended: true };

      const typeMatch = content.match(/type="([^"]*)"/);
      if (typeMatch) masterPage.type = typeMatch[1].toLowerCase() as 'both' | 'even' | 'odd';

      const pageNumberMatch = content.match(/pageNumber="(\d+)"/);
      if (pageNumberMatch) masterPage.pageNumber = parseInt(pageNumberMatch[1]);

      const pageDuplicateMatch = content.match(/pageDuplicate="(true|false|1|0)"/i);
      if (pageDuplicateMatch) masterPage.pageDuplicate = pageDuplicateMatch[1] === 'true' || pageDuplicateMatch[1] === '1';

      const pageFrontMatch = content.match(/pageFront="(true|false|1|0)"/i);
      if (pageFrontMatch) masterPage.pageFront = pageFrontMatch[1] === 'true' || pageFrontMatch[1] === '1';

      const paragraphs: HwpxParagraph[] = [];
      const paraRegex = /<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/g;
      let paraMatch;
      while ((paraMatch = paraRegex.exec(content)) !== null) {
        paragraphs.push(this.parseParagraph(paraMatch[0]));
      }
      if (paragraphs.length > 0) {
        masterPage.paragraphs = paragraphs;
      }

      masterPages.push(masterPage);
    }

    return masterPages.length > 0 ? masterPages : undefined;
  }

  private static parseColumnDef(xml: string): import('./types').ColumnDef | undefined {
    const colDefMatch = xml.match(/<hp:colDef[^>]*>([\s\S]*?)<\/hp:colDef>/i);
    if (!colDefMatch) return undefined;

    const content = colDefMatch[0];
    const colDef: import('./types').ColumnDef = {};

    const typeMatch = content.match(/type="([^"]*)"/);
    if (typeMatch) colDef.type = typeMatch[1].toLowerCase() as 'newspaper' | 'balanced' | 'parallel';

    const countMatch = content.match(/count="(\d+)"/);
    if (countMatch) colDef.count = parseInt(countMatch[1]);

    const layoutMatch = content.match(/layout="([^"]*)"/);
    if (layoutMatch) colDef.layout = layoutMatch[1].toLowerCase() as 'left' | 'right' | 'mirror';

    const sameSizeMatch = content.match(/sameSize="(true|false|1|0)"/i);
    if (sameSizeMatch) colDef.sameSize = sameSizeMatch[1] === 'true' || sameSizeMatch[1] === '1';

    const sameGapMatch = content.match(/sameGap="(\d+)"/);
    if (sameGapMatch) colDef.sameGap = parseInt(sameGapMatch[1]);

    // Parse column line
    const columnLineMatch = content.match(/<hp:columnLine[^>]*type="([^"]*)"[^>]*width="([^"]*)"[^>]*color="([^"]*)"/i);
    if (columnLineMatch) {
      colDef.columnLine = {
        type: columnLineMatch[1].toLowerCase() as import('./types').LineType1,
        width: columnLineMatch[2] as import('./types').LineWidth,
        color: columnLineMatch[3]
      };
    }

    // Parse columns
    const columns: import('./types').ColumnInfo[] = [];
    const columnRegex = /<hp:column[^>]*width="(\d+)"[^>]*gap="(\d+)"/gi;
    let columnMatch;
    while ((columnMatch = columnRegex.exec(content)) !== null) {
      columns.push({
        width: parseInt(columnMatch[1]),
        gap: parseInt(columnMatch[2])
      });
    }
    if (columns.length > 0) {
      colDef.columns = columns;
    }

    return colDef;
  }

  private static parseImageEffects(xml: string): import('./types').ImageEffects | undefined {
    const effectsMatch = xml.match(/<hp:effects[^>]*>([\s\S]*?)<\/hp:effects>/i);
    if (!effectsMatch) return undefined;

    const content = effectsMatch[1];
    const effects: import('./types').ImageEffects = {};

    // Parse shadow effect
    const shadowMatch = content.match(/<hp:shadowEffect[^>]*>([\s\S]*?)<\/hp:shadowEffect>/i);
    if (shadowMatch) {
      const shadowContent = shadowMatch[0];
      effects.shadow = {};

      const styleMatch = shadowContent.match(/style="([^"]*)"/);
      if (styleMatch) effects.shadow.style = styleMatch[1];

      const alphaMatch = shadowContent.match(/alpha="([^"]*)"/);
      if (alphaMatch) effects.shadow.alpha = parseFloat(alphaMatch[1]);

      const radiusMatch = shadowContent.match(/radius="([^"]*)"/);
      if (radiusMatch) effects.shadow.radius = parseFloat(radiusMatch[1]);

      const directionMatch = shadowContent.match(/direction="([^"]*)"/);
      if (directionMatch) effects.shadow.direction = parseFloat(directionMatch[1]);

      const distanceMatch = shadowContent.match(/distance="([^"]*)"/);
      if (distanceMatch) effects.shadow.distance = parseFloat(distanceMatch[1]);

      // Parse color
      const colorMatch = shadowContent.match(/<hp:effectsColor[^>]*colorR="(\d+)"[^>]*colorG="(\d+)"[^>]*colorB="(\d+)"/i);
      if (colorMatch) {
        effects.shadow.color = `rgb(${colorMatch[1]},${colorMatch[2]},${colorMatch[3]})`;
      }
    }

    // Parse glow effect
    const glowMatch = content.match(/<hp:glow[^>]*>([\s\S]*?)<\/hp:glow>/i);
    if (glowMatch) {
      const glowContent = glowMatch[0];
      effects.glow = {};

      const alphaMatch = glowContent.match(/alpha="([^"]*)"/);
      if (alphaMatch) effects.glow.alpha = parseFloat(alphaMatch[1]);

      const radiusMatch = glowContent.match(/radius="([^"]*)"/);
      if (radiusMatch) effects.glow.radius = parseFloat(radiusMatch[1]);

      const colorMatch = glowContent.match(/<hp:effectsColor[^>]*colorR="(\d+)"[^>]*colorG="(\d+)"[^>]*colorB="(\d+)"/i);
      if (colorMatch) {
        effects.glow.color = `rgb(${colorMatch[1]},${colorMatch[2]},${colorMatch[3]})`;
      }
    }

    // Parse soft edge effect
    const softEdgeMatch = content.match(/<hp:softEdge[^>]*radius="([^"]*)"/i);
    if (softEdgeMatch) {
      effects.softEdge = { radius: parseFloat(softEdgeMatch[1]) };
    }

    // Parse reflection effect
    const reflectionMatch = content.match(/<hp:reflection[^>]*>([\s\S]*?)<\/hp:reflection>|<hp:reflection([^>]*)\/>/i);
    if (reflectionMatch) {
      const refContent = reflectionMatch[0];
      effects.reflection = {};

      const radiusMatch = refContent.match(/radius="([^"]*)"/);
      if (radiusMatch) effects.reflection.radius = parseFloat(radiusMatch[1]);

      const directionMatch = refContent.match(/direction="([^"]*)"/);
      if (directionMatch) effects.reflection.direction = parseFloat(directionMatch[1]);

      const distanceMatch = refContent.match(/distance="([^"]*)"/);
      if (distanceMatch) effects.reflection.distance = parseFloat(distanceMatch[1]);

      const startAlphaMatch = refContent.match(/startAlpha="([^"]*)"/);
      if (startAlphaMatch) effects.reflection.startAlpha = parseFloat(startAlphaMatch[1]);

      const endAlphaMatch = refContent.match(/endAlpha="([^"]*)"/);
      if (endAlphaMatch) effects.reflection.endAlpha = parseFloat(endAlphaMatch[1]);
    }

    return Object.keys(effects).length > 0 ? effects : undefined;
  }

  private static parseParagraphsSimple(xml: string): HwpxParagraph[] {
    const paragraphs: HwpxParagraph[] = [];
    const paragraphRegex = /<hp:p[^>]*>([\s\S]*?)<\/hp:p>/g;
    let match;

    while ((match = paragraphRegex.exec(xml)) !== null) {
      const paragraph = this.parseParagraph(match[0]);
      paragraphs.push(paragraph);
    }

    return paragraphs;
  }

  private static parseParagraph(xml: string): HwpxParagraph {
    const paragraph: HwpxParagraph = {
      id: generateId(),
      runs: [],
    };

    // Extract only the opening <hp:p ...> tag to get paragraph attributes
    const pTagMatch = xml.match(/^<hp:p\s+([^>]*)>/);
    const pTagAttrs = pTagMatch ? pTagMatch[1] : '';

    // Check for page break on this paragraph (only in the <hp:p> tag itself)
    const pageBreakMatch = pTagAttrs.match(/pageBreak="([^"]*)"/);
    if (pageBreakMatch && pageBreakMatch[1] === '1') {
      paragraph.pageBreak = true;
    }

    // Get paragraph shape reference from the <hp:p> tag
    const paraShapeRefMatch = pTagAttrs.match(/paraPrIDRef="(\d+)"/);
    if (paraShapeRefMatch) {
      const paraShape = this.styles.paraShapes.get(parseInt(paraShapeRefMatch[1]));
      if (paraShape) {
        paragraph.paraStyle = {
          align: paraShape.align as ParagraphStyle['align'],
          lineSpacing: paraShape.lineSpacing,
          marginTop: paraShape.marginTop,
          marginBottom: paraShape.marginBottom,
          marginLeft: paraShape.marginLeft,
          marginRight: paraShape.marginRight,
          firstLineIndent: paraShape.firstLineIndent,
          keepWithNext: paraShape.keepWithNext,
          keepLines: paraShape.keepLines,
        };
        // Check pageBreakBefore from paragraph shape
        if (paraShape.pageBreakBefore) {
          paragraph.pageBreak = true;
        }
      }
    }

    const runRegex = /<hp:run[^>]*>([\s\S]*?)<\/hp:run>/g;
    let runMatch;

    while ((runMatch = runRegex.exec(xml)) !== null) {
      const runContent = runMatch[0];
      const parsedRuns = this.parseRun(runContent);
      paragraph.runs.push(...parsedRuns);
    }

    if (paragraph.runs.length === 0) {
      const textRegex = /<hp:t[^>]*>([^<]*)<\/hp:t>/g;
      let textMatch;
      while ((textMatch = textRegex.exec(xml)) !== null) {
        paragraph.runs.push({ text: this.decodeXmlEntities(textMatch[1]) });
      }
    }

    const listMatch = xml.match(/<hp:lineseg[^>]*listLevel="(\d+)"/);
    if (listMatch) {
      paragraph.listLevel = parseInt(listMatch[1]);
      paragraph.listType = 'bullet';
    }

    // Parse lineseg info for accurate layout
    const linesegRegex = /<hp:lineseg[^>]*vertpos="(\d+)"[^>]*vertsize="(\d+)"[^>]*textheight="(\d+)"[^>]*baseline="(\d+)"[^>]*spacing="(\d+)"/g;
    let linesegMatch;
    const linesegs: import('./types').LineSeg[] = [];
    while ((linesegMatch = linesegRegex.exec(xml)) !== null) {
      linesegs.push({
        vertpos: parseInt(linesegMatch[1]) / 100,
        vertsize: parseInt(linesegMatch[2]) / 100,
        textheight: parseInt(linesegMatch[3]) / 100,
        baseline: parseInt(linesegMatch[4]) / 100,
        spacing: parseInt(linesegMatch[5]) / 100,
      });
    }
    if (linesegs.length > 0) {
      paragraph.linesegs = linesegs;
    }

    return paragraph;
  }

  private static parseRun(xml: string): TextRun[] {
    const runs: TextRun[] = [];
    
    let charStyle: TextRun['charStyle'] | undefined;
    const charShapeRefMatch = xml.match(/charPrIDRef="(\d+)"/);
    if (charShapeRefMatch) {
      const charShape = this.styles.charShapes.get(parseInt(charShapeRefMatch[1]));
      if (charShape) {
        charStyle = {
          fontName: charShape.fontName,
          fontSize: charShape.fontSize,
          bold: charShape.bold,
          italic: charShape.italic,
          underline: charShape.underline,
          underlineType: charShape.underlineType,
          underlineShape: charShape.underlineShape,
          underlineColor: charShape.underlineColor,
          strikethrough: charShape.strikethrough,
          strikeoutShape: charShape.strikeoutShape,
          strikeoutColor: charShape.strikeoutColor,
          fontColor: charShape.color,
          backgroundColor: charShape.backgroundColor,
          charSpacing: charShape.charSpacing,
          relativeSize: charShape.relativeSize,
          charOffset: charShape.charOffset,
          emphasisMark: charShape.emphasisMark,
          useFontSpace: charShape.useFontSpace,
          useKerning: charShape.useKerning,
          outline: charShape.outline,
          shadow: charShape.shadow,
          shadowX: charShape.shadowX,
          shadowY: charShape.shadowY,
          shadowColor: charShape.shadowColor,
          emboss: charShape.emboss,
          engrave: charShape.engrave,
          smallCaps: charShape.smallCaps,
        };
      }
    }

    let hyperlink: import('./types').HyperlinkField | undefined;
    let field: import('./types').FieldControl | undefined;
    
    const fieldBeginMatch = xml.match(/<hp:fieldBegin[^>]*type="([^"]*)"[^>]*>([\s\S]*?)<\/hp:fieldBegin>/i);
    if (fieldBeginMatch) {
      const fieldType = fieldBeginMatch[1].toUpperCase();
      const fieldContent = fieldBeginMatch[2];
      
      if (fieldType === 'HYPERLINK') {
        const paramMatch = fieldContent.match(/<hp:stringParam[^>]*name="URL"[^>]*>([^<]*)<\/hp:stringParam>/i);
        if (paramMatch) {
          hyperlink = {
            fieldType: 'hyperlink',
            url: this.decodeXmlEntities(paramMatch[1].trim()),
          };
        } else {
          const commandMatch = fieldContent.match(/<hp:stringParam[^>]*name="Command"[^>]*>([^<]*)<\/hp:stringParam>/i);
          if (commandMatch) {
            const urlPart = commandMatch[1].split(';')[0] || commandMatch[1];
            hyperlink = {
              fieldType: 'hyperlink',
              url: this.decodeXmlEntities(urlPart.trim()),
            };
          }
        }
      } else if (fieldType === 'MEMO') {
        const memoField: import('./types').MemoField = { fieldType: 'memo' };
        
        const authorMatch = fieldContent.match(/<hp:stringParam[^>]*name="Author"[^>]*>([^<]*)<\/hp:stringParam>/i);
        if (authorMatch) memoField.author = authorMatch[1];
        
        const dateMatch = fieldContent.match(/<hp:stringParam[^>]*name="CreateDateTime"[^>]*>([^<]*)<\/hp:stringParam>/i);
        if (dateMatch) memoField.date = dateMatch[1];
        
        const memoTextMatch = fieldContent.match(/<hp:subList[^>]*>[\s\S]*?<hp:t[^>]*>([^<]*)<\/hp:t>/i);
        if (memoTextMatch) memoField.memoContent = memoTextMatch[1];
        
        field = memoField;
      } else if (fieldType === 'FORMULA') {
        const formulaField: import('./types').FormulaField = { fieldType: 'formula' };
        
        const scriptMatch = fieldContent.match(/<hp:stringParam[^>]*name="(?:Script|Command)"[^>]*>([^<]*)<\/hp:stringParam>/i);
        if (scriptMatch) formulaField.formulaScript = scriptMatch[1];
        
        field = formulaField;
      } else if (fieldType === 'BOOKMARK') {
        const bookmarkField: import('./types').BookmarkField = { fieldType: 'bookmark', bookmarkName: '' };
        
        const nameMatch = fieldContent.match(/<hp:stringParam[^>]*name="(?:Name|BookmarkName)"[^>]*>([^<]*)<\/hp:stringParam>/i);
        if (nameMatch) bookmarkField.bookmarkName = nameMatch[1];
        
        field = bookmarkField;
      } else {
        const fieldTypeMap: Record<string, import('./types').FieldType> = {
          'DATE': 'date', 'DOCDATE': 'docDate', 'PATH': 'path',
          'MAILMERGE': 'mailMerge', 'CROSSREF': 'crossRef', 'CLICKHERE': 'clickHere',
          'SUMMARY': 'summary', 'USERINFO': 'userInfo', 'REVISIONSIGN': 'revisionSign',
          'PRIVATETXT': 'privateTxt', 'TABLEOFCONTENTS': 'tableOfContents'
        };
        field = {
          fieldType: fieldTypeMap[fieldType] || 'unknown',
          name: fieldBeginMatch[0].match(/name="([^"]*)"/)?.[1],
        };
      }
    }
    
    if (!hyperlink) {
      const hyperlinkMatch = xml.match(/<hp:ctrl[^>]*>[\s\S]*?<hp:fieldBegin[^>]*type="HYPERLINK"[^>]*(?:param="([^"]*)")?/i);
      if (hyperlinkMatch) {
        const paramStr = hyperlinkMatch[1] || '';
        const urlMatch = paramStr.match(/url:([^;]*)/i) || paramStr.match(/^([^;]+)/);
        if (urlMatch) {
          hyperlink = {
            fieldType: 'hyperlink',
            url: this.decodeXmlEntities(urlMatch[1].trim()),
          };
        }
      }
    }

    // Check if this run has a MEMO field and extract memo ID
    const hasMemo = /<hp:fieldBegin[^>]*type="MEMO"/i.test(xml);
    let memoId: string | undefined;
    if (hasMemo) {
      const memoIdMatch = xml.match(/<hp:fieldBegin[^>]*type="MEMO"[^>]*id="([^"]*)"/i);
      if (memoIdMatch) memoId = memoIdMatch[1];
    }

    // Remove MEMO field's subList content to prevent memo text from appearing in document body
    // Memo content is stored separately, not rendered as regular text
    // Always remove MEMO fieldBegin content regardless of field detection
    let textSearchXml = xml.replace(/<hp:fieldBegin[^>]*type="MEMO"[^>]*>[\s\S]*?<\/hp:fieldBegin>/gi, '');

    // Detect footnote/endnote references and extract their numbers for markers
    const footnoteMatch = xml.match(/<hp:footNote\b[^>]*number="(\d+)"[^>]*>/i);
    const endnoteMatch = xml.match(/<hp:endNote\b[^>]*number="(\d+)"[^>]*>/i);
    const footnoteNumber = footnoteMatch ? parseInt(footnoteMatch[1]) : null;
    const endnoteNumber = endnoteMatch ? parseInt(endnoteMatch[1]) : null;

    // Remove footnote/endnote content to prevent it from appearing in document body
    // Footnotes are parsed separately and displayed at the bottom of the page
    textSearchXml = textSearchXml.replace(/<hp:footNote\b[^>]*>[\s\S]*?<\/hp:footNote>/gi, '');
    textSearchXml = textSearchXml.replace(/<hp:endNote\b[^>]*>[\s\S]*?<\/hp:endNote>/gi, '');

    const allTextTagsRegex = /<hp:t(?:\s[^>]*)?>(?:([\s\S]*?)<\/hp:t>)?|<hp:t\s*\/>/g;
    let tMatch;
    let foundTextTags = false;

    while ((tMatch = allTextTagsRegex.exec(textSearchXml)) !== null) {
      foundTextTags = true;
      const tContent = tMatch[1] || '';
      this.processTextContent(tContent, charStyle, runs, hyperlink, field);
    }

    if (!foundTextTags) {
      return runs;
    }

    if (runs.length === 0) {
      runs.push({ text: '', charStyle, hyperlink, field });
    }

    // Mark all runs in this block as having memo if applicable
    if (hasMemo) {
      for (const run of runs) {
        run.hasMemo = true;
        if (memoId) run.memoId = memoId;
      }
    }

    // Add footnote/endnote marker as superscript at the end
    if (footnoteNumber !== null) {
      runs.push({
        text: `${footnoteNumber})`,
        footnoteRef: footnoteNumber,
        charStyle: { ...charStyle, superscript: true, fontSize: charStyle?.fontSize ? charStyle.fontSize * 0.7 : 7 },
      });
    }
    if (endnoteNumber !== null) {
      runs.push({
        text: `${endnoteNumber})`,
        endnoteRef: endnoteNumber,
        charStyle: { ...charStyle, superscript: true, fontSize: charStyle?.fontSize ? charStyle.fontSize * 0.7 : 7 },
      });
    }

    return runs;
  }

  private static processTextContent(
    tContent: string,
    charStyle: TextRun['charStyle'] | undefined,
    runs: TextRun[],
    hyperlink?: import('./types').HyperlinkField,
    field?: import('./types').FieldControl
  ): void {
    // Combined regex to match all special elements including:
    // tab, lineBreak, hypen, nbSpace, fwSpace, titleMark, markPenBegin, markPenEnd,
    // autoNum, newNum, compose, dutmal, indexMark, pageHiding, pageNumCtrl, pageNum
    const specialElementRegex = /<hp:(tab|lineBreak|hypen|nbSpace|fwSpace|titleMark|markPenBegin|markPenEnd|autoNum|newNum|compose|dutmal|indexMark|pageHiding|pageNumCtrl|pageNum)(?:\s+([^>]*))?\s*(?:\/>|>([\s\S]*?)<\/hp:\1>)/gi;
    let lastIndex = 0;
    let specialMatch;
    let currentMarkPenColor: string | undefined;

    while ((specialMatch = specialElementRegex.exec(tContent)) !== null) {
      // Process text before this special element
      if (specialMatch.index > lastIndex) {
        const textBefore = tContent.substring(lastIndex, specialMatch.index);
        const cleanText = textBefore.replace(/<[^>]+>/g, '');
        if (cleanText) {
          const run: TextRun = { text: this.decodeXmlEntities(cleanText), charStyle, hyperlink, field };
          if (currentMarkPenColor) {
            run.markPen = { color: currentMarkPenColor };
          }
          runs.push(run);
        }
      }

      const elementType = specialMatch[1].toLowerCase();
      const attrs = specialMatch[2] || '';

      switch (elementType) {
        case 'tab': {
          const widthMatch = attrs.match(/width="(\d+)"/);
          const width = widthMatch ? parseInt(widthMatch[1]) / 100 : 0;
          const leaderMatch = attrs.match(/leader="(\d+)"/);
          const leaderType = leaderMatch ? parseInt(leaderMatch[1]) : 0;
          // LineType2: 0=None, 1=Solid, 2=Dash, 3=Dot, 4=DashDot, 5=DashDotDot, 6=LongDash, 7=CircleDot
          let leader: 'none' | 'solid' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot' = 'none';
          if (leaderType === 1) leader = 'solid';
          else if (leaderType === 2) leader = 'dash';
          else if (leaderType === 3) leader = 'dot';
          else if (leaderType === 4) leader = 'dashDot';
          else if (leaderType === 5) leader = 'dashDotDot';
          runs.push({ text: '', charStyle, tab: { width, leader }, hyperlink, field });
          break;
        }
        case 'linebreak': {
          // Line break (강제 줄 나눔, SHIFT-ENTER)
          runs.push({ text: '\n', charStyle, hyperlink, field });
          break;
        }
        case 'hypen': {
          // Soft hyphen (하이픈, CTRL-SHIFT-'-')
          runs.push({ text: '\u00AD', charStyle, hyperlink, field });
          break;
        }
        case 'nbspace': {
          // Non-breaking space (묶음 빈칸, CTRL-ALT-SPACE)
          runs.push({ text: '\u00A0', charStyle, hyperlink, field });
          break;
        }
        case 'fwspace': {
          // Full-width space (고정폭 빈칸, ALT-SPACE)
          runs.push({ text: '\u3000', charStyle, hyperlink, field });
          break;
        }
        case 'titlemark': {
          // Title mark (제목 차례 표시) - just mark it, no visible text
          const ignoreMatch = attrs.match(/ignore="(true|false|1|0)"/i);
          const ignore = ignoreMatch ? (ignoreMatch[1] === 'true' || ignoreMatch[1] === '1') : false;
          // Title mark doesn't produce visible text but marks the title for TOC
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'markpenbegin': {
          // Highlight/mark pen start (형광펜 시작)
          const colorMatch = attrs.match(/color="([^"]*)"/);
          currentMarkPenColor = colorMatch ? colorMatch[1] : '#FFFF00'; // default yellow
          break;
        }
        case 'markpenend': {
          // Highlight/mark pen end (형광펜 끝)
          currentMarkPenColor = undefined;
          break;
        }
        case 'autonum': {
          // Auto number (자동번호 - 각주, 표, 그림 등의 자동 번호)
          const numTypeMatch = attrs.match(/numType="([^"]*)"/);
          const numType = numTypeMatch ? numTypeMatch[1] : 'Page';
          // AutoNum generates a number based on context, we just mark its presence
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'newnum': {
          // New number (새 번호 지정)
          const numTypeMatch = attrs.match(/numType="([^"]*)"/);
          const numMatch = attrs.match(/num="(\d+)"/);
          // NewNum resets the numbering
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'compose': {
          // Compose (글자 겹침)
          const innerContent = specialMatch[3] || '';
          // Compose overlaps characters - we extract the text content
          const composeText = innerContent.replace(/<[^>]+>/g, '');
          if (composeText) {
            runs.push({ text: this.decodeXmlEntities(composeText), charStyle, hyperlink, field });
          }
          break;
        }
        case 'dutmal': {
          // Dutmal (덧말/루비)
          const mainTextMatch = attrs.match(/mainText="([^"]*)"/);
          const subTextMatch = attrs.match(/subText="([^"]*)"/);
          const mainText = mainTextMatch ? mainTextMatch[1] : '';
          const subText = subTextMatch ? subTextMatch[1] : '';
          // For now, just output the main text (subText is annotation above/below)
          if (mainText) {
            runs.push({ text: this.decodeXmlEntities(mainText), charStyle, hyperlink, field });
          }
          break;
        }
        case 'indexmark': {
          // Index mark (찾아보기 표식)
          const keyFirstMatch = attrs.match(/keyFirst="([^"]*)"/);
          const keySecondMatch = attrs.match(/keySecond="([^"]*)"/);
          // Index mark is invisible but marks text for index
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'pagehiding': {
          // Page hiding (감추기)
          // This is a control element, doesn't produce visible text
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'pagenumctrl': {
          // Page number control (쪽 번호 컨트롤)
          // This is a control element for page number settings
          runs.push({ text: '', charStyle, hyperlink, field });
          break;
        }
        case 'pagenum': {
          // Page number (쪽 번호)
          // This displays the page number - we use a placeholder
          runs.push({ text: '#', charStyle, hyperlink, field });
          break;
        }
      }

      lastIndex = specialMatch.index + specialMatch[0].length;
    }

    // Process remaining text after last special element
    if (lastIndex < tContent.length) {
      const remainingText = tContent.substring(lastIndex);
      const cleanText = remainingText.replace(/<[^>]+>/g, '');
      if (cleanText) {
        const run: TextRun = { text: this.decodeXmlEntities(cleanText), charStyle, hyperlink, field };
        if (currentMarkPenColor) {
          run.markPen = { color: currentMarkPenColor };
        }
        runs.push(run);
      }
    } else if (lastIndex === 0 && tContent.length > 0) {
      const cleanText = tContent.replace(/<[^>]+>/g, '');
      if (cleanText) {
        runs.push({ text: this.decodeXmlEntities(cleanText), charStyle, hyperlink, field });
      }
    }
  }

  private static parseTable(xml: string): HwpxTable {
    const table: HwpxTable = {
      id: generateId(),
      rows: [],
      columnWidths: [],
    };

    const tblTagMatch = xml.match(/<hp:tbl[^>]*>/);
    if (tblTagMatch) {
      const tblAttrs = tblTagMatch[0];

      const idMatch = tblAttrs.match(/\bid="(\d+)"/);
      if (idMatch) table.id = idMatch[1];

      const zOrderMatch = tblAttrs.match(/zOrder="(\d+)"/);
      if (zOrderMatch) table.zOrder = parseInt(zOrderMatch[1]);

      const numTypeMatch = tblAttrs.match(/numberingType="([^"]*)"/);
      if (numTypeMatch) {
        const map: Record<string, import('./types').NumberingType> = {
          'NONE': 'none', 'PICTURE': 'picture', 'TABLE': 'table', 'EQUATION': 'equation'
        };
        table.numberingType = map[numTypeMatch[1].toUpperCase()] || 'none';
      }

      const textWrapMatch = tblAttrs.match(/textWrap="([^"]*)"/);
      if (textWrapMatch) {
        const map: Record<string, import('./types').TextWrapType> = {
          'SQUARE': 'square', 'TIGHT': 'tight', 'THROUGH': 'through',
          'TOP_AND_BOTTOM': 'topAndBottom', 'BEHIND_TEXT': 'behindText', 'IN_FRONT_OF_TEXT': 'inFrontOfText'
        };
        table.textWrap = map[textWrapMatch[1].toUpperCase()] || 'square';
      }

      const textFlowMatch = tblAttrs.match(/textFlow="([^"]*)"/);
      if (textFlowMatch) {
        const map: Record<string, import('./types').TextFlowType> = {
          'BOTH_SIDES': 'bothSides', 'LEFT_ONLY': 'leftOnly', 'RIGHT_ONLY': 'rightOnly', 'LARGEST_ONLY': 'largestOnly'
        };
        table.textFlow = map[textFlowMatch[1].toUpperCase()] || 'bothSides';
      }

      const pageBreakMatch = tblAttrs.match(/pageBreak="([^"]*)"/);
      if (pageBreakMatch) {
        const map: Record<string, import('./types').PageBreakType> = {
          'CELL': 'cell', 'NONE': 'none', 'TABLE': 'table'
        };
        table.pageBreak = map[pageBreakMatch[1].toUpperCase()] || 'none';
      }

      const repeatHeaderMatch = tblAttrs.match(/repeatHeader="([^"]*)"/);
      if (repeatHeaderMatch) {
        table.repeatHeader = repeatHeaderMatch[1] === '1' || repeatHeaderMatch[1] === 'true';
      }

      const rowCntMatch = tblAttrs.match(/rowCnt="(\d+)"/);
      if (rowCntMatch) table.rowCnt = parseInt(rowCntMatch[1]);

      const colCntMatch = tblAttrs.match(/colCnt="(\d+)"/);
      if (colCntMatch) table.colCnt = parseInt(colCntMatch[1]);

      const lockMatch = tblAttrs.match(/lock="([^"]*)"/);
      if (lockMatch) {
        table.lock = lockMatch[1] === '1' || lockMatch[1] === 'true';
      }
    }

    const szMatch = xml.match(/<hp:sz\s+width="(\d+)"[^>]*height="(\d+)"/);
    if (szMatch) {
      table.width = parseInt(szMatch[1]) / 100;
      table.height = parseInt(szMatch[2]) / 100;
    }

    const cellSpacingMatch = xml.match(/cellSpacing="(\d+)"/);
    if (cellSpacingMatch) {
      table.cellSpacing = parseInt(cellSpacingMatch[1]) / 100;
    }

    const borderFillMatch = xml.match(/borderFillIDRef="(\d+)"/);
    if (borderFillMatch) {
      table.borderFillId = parseInt(borderFillMatch[1]);
    }

    // Parse CellZoneList
    const cellZoneListMatch = xml.match(/<hp:cellzoneList[^>]*>([\s\S]*?)<\/hp:cellzoneList>/i);
    if (cellZoneListMatch) {
      const cellZones: import('./types').CellZone[] = [];
      const cellZoneRegex = /<hp:cellzone[^>]*startRowAddr="(\d+)"[^>]*startColAddr="(\d+)"[^>]*endRowAddr="(\d+)"[^>]*endColAddr="(\d+)"(?:[^>]*borderFillIDRef="(\d+)")?/gi;
      let czMatch;
      while ((czMatch = cellZoneRegex.exec(cellZoneListMatch[1])) !== null) {
        const cellZone: import('./types').CellZone = {
          startRowAddr: parseInt(czMatch[1]),
          startColAddr: parseInt(czMatch[2]),
          endRowAddr: parseInt(czMatch[3]),
          endColAddr: parseInt(czMatch[4])
        };
        if (czMatch[5]) cellZone.borderFill = parseInt(czMatch[5]);
        cellZones.push(cellZone);
      }
      if (cellZones.length > 0) {
        table.cellZoneList = cellZones;
      }
    }

    const posMatch = xml.match(/<hp:pos[^>]*>/);
    if (posMatch) {
      const pos = posMatch[0];
      table.position = {};

      if (pos.includes('treatAsChar="1"') || pos.includes('treatAsChar="true"')) {
        table.position.treatAsChar = true;
      }
      if (pos.includes('flowWithText="1"') || pos.includes('flowWithText="true"')) {
        table.position.flowWithText = true;
      }

      const vertRelMatch = pos.match(/vertRelTo="([^"]*)"/);
      if (vertRelMatch) {
        const map: Record<string, import('./types').VertRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'PARA': 'para'
        };
        table.position.vertRelTo = map[vertRelMatch[1].toUpperCase()];
      }

      const horzRelMatch = pos.match(/horzRelTo="([^"]*)"/);
      if (horzRelMatch) {
        const map: Record<string, import('./types').HorzRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'COLUMN': 'column', 'PARA': 'para'
        };
        table.position.horzRelTo = map[horzRelMatch[1].toUpperCase()];
      }

      const vertAlignMatch = pos.match(/vertAlign="([^"]*)"/);
      if (vertAlignMatch) {
        const map: Record<string, import('./types').VertAlign> = {
          'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom'
        };
        table.position.vertAlign = map[vertAlignMatch[1].toUpperCase()];
      }

      const horzAlignMatch = pos.match(/horzAlign="([^"]*)"/);
      if (horzAlignMatch) {
        const map: Record<string, import('./types').HorzAlign> = {
          'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right'
        };
        table.position.horzAlign = map[horzAlignMatch[1].toUpperCase()];
      }

      const vertOffsetMatch = pos.match(/vertOffset="(-?\d+)"/);
      if (vertOffsetMatch) {
        table.position.vertOffset = parseInt(vertOffsetMatch[1]) / 100;
      }

      const horzOffsetMatch = pos.match(/horzOffset="(-?\d+)"/);
      if (horzOffsetMatch) {
        table.position.horzOffset = parseInt(horzOffsetMatch[1]) / 100;
      }
    }

    // Parse outMargin - handle attributes in any order
    const outMarginTagMatch = xml.match(/<hp:outMargin\s+([^>]*)\/?\s*>/);
    if (outMarginTagMatch) {
      const attrs = outMarginTagMatch[1];
      const leftMatch = attrs.match(/left="(\d+)"/);
      const rightMatch = attrs.match(/right="(\d+)"/);
      const topMatch = attrs.match(/top="(\d+)"/);
      const bottomMatch = attrs.match(/bottom="(\d+)"/);
      table.outMargin = {
        left: leftMatch ? parseInt(leftMatch[1]) / 100 : 0,
        right: rightMatch ? parseInt(rightMatch[1]) / 100 : 0,
        top: topMatch ? parseInt(topMatch[1]) / 100 : 0,
        bottom: bottomMatch ? parseInt(bottomMatch[1]) / 100 : 0
      };
    }

    // Parse inMargin - handle attributes in any order
    const inMarginTagMatch = xml.match(/<hp:inMargin\s+([^>]*)\/?\s*>/);
    if (inMarginTagMatch) {
      const attrs = inMarginTagMatch[1];
      const leftMatch = attrs.match(/left="(\d+)"/);
      const rightMatch = attrs.match(/right="(\d+)"/);
      const topMatch = attrs.match(/top="(\d+)"/);
      const bottomMatch = attrs.match(/bottom="(\d+)"/);
      table.inMargin = {
        left: leftMatch ? parseInt(leftMatch[1]) / 100 : 0,
        right: rightMatch ? parseInt(rightMatch[1]) / 100 : 0,
        top: topMatch ? parseInt(topMatch[1]) / 100 : 0,
        bottom: bottomMatch ? parseInt(bottomMatch[1]) / 100 : 0
      };
    }

    const colWidthsMatch = xml.match(/<hp:colSz[^>]*>([\s\S]*?)<\/hp:colSz>/);
    if (colWidthsMatch) {
      const widthRegex = /(\d+)/g;
      let widthMatch;
      while ((widthMatch = widthRegex.exec(colWidthsMatch[1])) !== null) {
        table.columnWidths!.push(parseInt(widthMatch[1]) / 100);
      }
    }

    const rows = this.extractBalancedTags(xml, 'hp:tr');
    for (const rowXml of rows) {
      const row = this.parseTableRow(rowXml);
      table.rows.push(row);
    }

    return table;
  }

  private static parseTableRow(xml: string): TableRow {
    const row: TableRow = { cells: [] };

    const heightMatch = xml.match(/height="(\d+)"/);
    if (heightMatch) {
      row.height = parseInt(heightMatch[1]) / 100;
    }

    const cells = this.extractBalancedTags(xml, 'hp:tc');
    for (const cellXml of cells) {
      const cell = this.parseTableCell(cellXml);
      row.cells.push(cell);
    }

    return row;
  }

  private static extractBalancedTags(xml: string, tagName: string): string[] {
    const results: string[] = [];
    const openTag = `<${tagName}`;
    const closeTag = `</${tagName}>`;
    let pos = 0;

    while (pos < xml.length) {
      const startIdx = xml.indexOf(openTag, pos);
      if (startIdx === -1) break;

      let depth = 1;
      let searchPos = startIdx + openTag.length;
      
      while (depth > 0 && searchPos < xml.length) {
        const nextOpen = xml.indexOf(openTag, searchPos);
        const nextClose = xml.indexOf(closeTag, searchPos);

        if (nextClose === -1) break;

        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          searchPos = nextOpen + openTag.length;
        } else {
          depth--;
          if (depth === 0) {
            results.push(xml.substring(startIdx, nextClose + closeTag.length));
          }
          searchPos = nextClose + closeTag.length;
        }
      }

      pos = searchPos;
    }

    return results;
  }

  // Extract ALL paragraphs including nested ones (not just top-level)
  private static extractAllParagraphs(xml: string): { xml: string; start: number; end: number }[] {
    const results: { xml: string; start: number; end: number }[] = [];
    const closeTag = '</hp:p>';
    const pOpenRegex = /<hp:p\b[^>]*>/g;
    // Regex to find opening paragraph tags (must be followed by space, >, or end of attributes)
    const pOpenSearchRegex = /<hp:p[\s>]/g;
    let match;

    while ((match = pOpenRegex.exec(xml)) !== null) {
      const startPos = match.index;

      // Find matching close tag using depth tracking
      let depth = 1;
      let searchPos = startPos + match[0].length;

      while (depth > 0 && searchPos < xml.length) {
        // Find next paragraph opening tag (not other hp:p* tags like hp:pagePr)
        pOpenSearchRegex.lastIndex = searchPos;
        const nextOpenMatch = pOpenSearchRegex.exec(xml);
        const nextOpen = nextOpenMatch ? nextOpenMatch.index : -1;
        const nextClose = xml.indexOf(closeTag, searchPos);

        if (nextClose === -1) break;

        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          searchPos = nextOpen + 6; // Move past '<hp:p ' or '<hp:p>'
        } else {
          depth--;
          if (depth === 0) {
            const endPos = nextClose + closeTag.length;
            results.push({
              xml: xml.substring(startPos, endPos),
              start: startPos,
              end: endPos
            });
          }
          searchPos = nextClose + closeTag.length;
        }
      }
    }

    return results;
  }

  private static parseTableCell(xml: string): TableCell {
    const cell: TableCell = { paragraphs: [] };

    const tcTagMatch = xml.match(/<hp:tc[^>]*>/);
    if (tcTagMatch) {
      const tcAttrs = tcTagMatch[0];

      // Try to get rowAddr/colAddr from tc attributes first
      let rowAddrMatch = tcAttrs.match(/rowAddr="(\d+)"/);
      let colAddrMatch = tcAttrs.match(/colAddr="(\d+)"/);

      // If not in tc attributes, try cellAddr element
      if (!rowAddrMatch || !colAddrMatch) {
        const cellAddrMatch = xml.match(/<hp:cellAddr[^>]*colAddr="(\d+)"[^>]*rowAddr="(\d+)"/);
        if (cellAddrMatch) {
          cell.colAddr = parseInt(cellAddrMatch[1]);
          cell.rowAddr = parseInt(cellAddrMatch[2]);
        }
      } else {
        if (rowAddrMatch) cell.rowAddr = parseInt(rowAddrMatch[1]);
        if (colAddrMatch) cell.colAddr = parseInt(colAddrMatch[1]);
      }

      const headerMatch = tcAttrs.match(/header="([^"]*)"/);
      if (headerMatch) {
        cell.header = headerMatch[1] === '1' || headerMatch[1] === 'true';
      }

      const protectMatch = tcAttrs.match(/protect="([^"]*)"/);
      if (protectMatch) {
        cell.protect = protectMatch[1] === '1' || protectMatch[1] === 'true';
      }

      const editableMatch = tcAttrs.match(/editable="([^"]*)"/);
      if (editableMatch) {
        cell.editable = editableMatch[1] === '1' || editableMatch[1] === 'true';
      }

      // hasMargin="0" means use table's inMargin instead of cell's own margin
      const hasMarginMatch = tcAttrs.match(/hasMargin="([^"]*)"/);
      if (hasMarginMatch) {
        cell.hasMargin = hasMarginMatch[1] === '1' || hasMarginMatch[1] === 'true';
      }
    }

    // textDirection and lineWrap can be in tc tag or subList element
    const subListMatch = xml.match(/<hp:subList[^>]*>/);
    const textDirSource = subListMatch ? subListMatch[0] : (tcTagMatch ? tcTagMatch[0] : '');

    const textDirMatch = textDirSource.match(/textDirection="([^"]*)"/);
    if (textDirMatch) {
      const dir = textDirMatch[1].toUpperCase();
      if (dir === 'VERTICAL' || dir === 'VERT') {
        cell.textDirection = 'vertical';
      } else {
        cell.textDirection = 'horizontal';
      }
    }

    const lineWrapMatch = textDirSource.match(/lineWrap="([^"]*)"/);
    if (lineWrapMatch) {
      const wrapMap: Record<string, import('./types').LineWrapType> = {
        'BREAK': 'break', 'SQUEEZE': 'squeeze', 'KEEP': 'keep'
      };
      cell.lineWrap = wrapMap[lineWrapMatch[1].toUpperCase()] || 'break';
    }

    // Get vertAlign from subList if not already found
    if (subListMatch) {
      const vertAlignMatch = subListMatch[0].match(/vertAlign="([^"]*)"/);
      if (vertAlignMatch && !cell.verticalAlign) {
        const align = vertAlignMatch[1].toLowerCase();
        if (align === 'center' || align === 'middle') {
          cell.verticalAlign = 'middle';
        } else if (align === 'bottom') {
          cell.verticalAlign = 'bottom';
        } else {
          cell.verticalAlign = 'top';
        }
      }
    }

    // Parse cellSz - handle attributes in any order
    const cellSzTagMatch = xml.match(/<hp:cellSz\s+([^>]*)\/?\s*>/);
    if (cellSzTagMatch) {
      const attrs = cellSzTagMatch[1];
      const widthMatch = attrs.match(/width="(\d+)"/);
      const heightMatch = attrs.match(/height="(\d+)"/);
      if (widthMatch) cell.width = parseInt(widthMatch[1]) / 100;
      if (heightMatch) cell.height = parseInt(heightMatch[1]) / 100;
    }

    // Parse cellSpan - handle attributes in any order
    const cellSpanTagMatch = xml.match(/<hp:cellSpan\s+([^>]*)\/?\s*>/);
    if (cellSpanTagMatch) {
      const attrs = cellSpanTagMatch[1];
      const colSpanMatch = attrs.match(/colSpan="(\d+)"/);
      const rowSpanMatch = attrs.match(/rowSpan="(\d+)"/);
      if (colSpanMatch) cell.colSpan = parseInt(colSpanMatch[1]);
      if (rowSpanMatch) cell.rowSpan = parseInt(rowSpanMatch[1]);
    } else {
      // Fallback: check for individual attributes
      const rowSpanMatch = xml.match(/rowSpan="(\d+)"/);
      if (rowSpanMatch) cell.rowSpan = parseInt(rowSpanMatch[1]);
      const colSpanMatch = xml.match(/colSpan="(\d+)"/);
      if (colSpanMatch) cell.colSpan = parseInt(colSpanMatch[1]);
    }

    // Parse cellMargin - handle attributes in any order
    const cellMarginTagMatch = xml.match(/<hp:cellMargin\s+([^>]*)\/?\s*>/);
    if (cellMarginTagMatch) {
      const attrs = cellMarginTagMatch[1];
      const leftMatch = attrs.match(/left="(\d+)"/);
      const rightMatch = attrs.match(/right="(\d+)"/);
      const topMatch = attrs.match(/top="(\d+)"/);
      const bottomMatch = attrs.match(/bottom="(\d+)"/);
      if (leftMatch) cell.marginLeft = parseInt(leftMatch[1]) / 100;
      if (rightMatch) cell.marginRight = parseInt(rightMatch[1]) / 100;
      if (topMatch) cell.marginTop = parseInt(topMatch[1]) / 100;
      if (bottomMatch) cell.marginBottom = parseInt(bottomMatch[1]) / 100;
    }

    const vertAlignMatch = xml.match(/vertAlign="([^"]*)"/);
    if (vertAlignMatch) {
      const align = vertAlignMatch[1].toLowerCase();
      if (align === 'center' || align === 'middle') {
        cell.verticalAlign = 'middle';
      } else if (align === 'bottom') {
        cell.verticalAlign = 'bottom';
      } else {
        cell.verticalAlign = 'top';
      }
    }

    const borderFillRefMatch = xml.match(/borderFillIDRef="(\d+)"/);
    if (borderFillRefMatch) {
      const borderFillId = parseInt(borderFillRefMatch[1]);
      cell.borderFillId = borderFillId;
      const borderFill = this.styles.borderFills.get(borderFillId);
      if (borderFill) {
        if (borderFill.fillColor) {
          cell.backgroundColor = borderFill.fillColor;
        }
        // Add gradation support
        if (borderFill.gradation && borderFill.gradation.colors.length > 0) {
          cell.backgroundGradation = {
            type: borderFill.gradation.type,
            angle: borderFill.gradation.angle,
            colors: borderFill.gradation.colors,
          };
        }
        if (borderFill.leftBorder) {
          cell.borderLeft = borderFill.leftBorder;
        }
        if (borderFill.rightBorder) {
          cell.borderRight = borderFill.rightBorder;
        }
        if (borderFill.topBorder) {
          cell.borderTop = borderFill.topBorder;
        }
        if (borderFill.bottomBorder) {
          cell.borderBottom = borderFill.bottomBorder;
        }
      }
    }

    if (!cell.backgroundColor) {
      const bgColorMatch = xml.match(/faceColor="([^"]*)"/);
      if (bgColorMatch && bgColorMatch[1] !== 'none') {
        cell.backgroundColor = bgColorMatch[1];
      }
    }

    const subLists = this.extractBalancedTags(xml, 'hp:subList');
    const contentXml = subLists.length > 0 
      ? subLists[0].replace(/^<hp:subList[^>]*>/, '').replace(/<\/hp:subList>$/, '')
      : xml;

    this.parseCellContent(contentXml, cell);

    return cell;
  }

  private static parseCellContent(contentXml: string, cell: TableCell): void {
    // Remove MEMO fieldBegin content from cell content to prevent memo text appearing in cell
    const cleanedXml = contentXml.replace(/<hp:fieldBegin[^>]*type="MEMO"[^>]*>[\s\S]*?<\/hp:fieldBegin>/gi, '');

    const nestedTables = this.extractBalancedTags(cleanedXml, 'hp:tbl');

    if (nestedTables.length > 0) {
      cell.nestedTables = [];
      cell.elements = [];

      let remainingXml = cleanedXml;
      for (const tableXml of nestedTables) {
        const tableIndex = remainingXml.indexOf(tableXml);
        if (tableIndex > 0) {
          const beforeTable = remainingXml.substring(0, tableIndex);
          const paragraphs = this.extractBalancedTags(beforeTable, 'hp:p');
          for (const pXml of paragraphs) {
            const paragraph = this.parseParagraph(pXml);
            cell.paragraphs.push(paragraph);
            cell.elements.push({ type: 'paragraph', data: paragraph });
          }
        }

        const nestedTable = this.parseTable(tableXml);
        cell.nestedTables.push(nestedTable);
        cell.elements.push({ type: 'table', data: nestedTable });

        remainingXml = remainingXml.substring(tableIndex + tableXml.length);
      }

      if (remainingXml) {
        const paragraphs = this.extractBalancedTags(remainingXml, 'hp:p');
        for (const pXml of paragraphs) {
          const paragraph = this.parseParagraph(pXml);
          cell.paragraphs.push(paragraph);
          cell.elements.push({ type: 'paragraph', data: paragraph });
        }
      }
    } else {
      const paragraphs = this.extractBalancedTags(cleanedXml, 'hp:p');
      for (const pXml of paragraphs) {
        const paragraph = this.parseParagraph(pXml);
        cell.paragraphs.push(paragraph);
      }
    }
  }

  private static parseImageElement(xml: string, content: HwpxContent): HwpxImage | null {
    let binaryRefMatch = xml.match(/<hc:img[^>]*binaryItemIDRef="([^"]*)"/);
    if (!binaryRefMatch) {
      binaryRefMatch = xml.match(/binaryItemIDRef="([^"]*)"/);
    }
    if (!binaryRefMatch) return null;

    const imageId = binaryRefMatch[1];
    const existingImage = content.images.get(imageId);

    // Use existing image from BinData if available, or create new
    const image: HwpxImage = existingImage ? {
      ...existingImage,
      id: existingImage.id || generateId(),
      width: existingImage.width || 100,
      height: existingImage.height || 100,
    } : {
      id: generateId(),
      binaryId: imageId,
      width: 100,
      height: 100,
    };

    // Try to get size from various tags (in order of priority)
    // 1. hp:sz - standard size
    // 2. hp:curSz - current display size
    // 3. hp:orgSz - original size (fallback)
    const szMatch = xml.match(/<hp:sz\s+width="(\d+)"[^>]*height="(\d+)"/);
    if (szMatch) {
      image.width = parseInt(szMatch[1]) / 100;
      image.height = parseInt(szMatch[2]) / 100;
    }

    const curSzMatch = xml.match(/<hp:curSz\s+width="(\d+)"[^>]*height="(\d+)"/);
    if (curSzMatch) {
      // curSz takes priority over sz if both exist
      image.width = parseInt(curSzMatch[1]) / 100;
      image.height = parseInt(curSzMatch[2]) / 100;
    }

    const orgSzMatch = xml.match(/<hp:orgSz\s+width="(\d+)"[^>]*height="(\d+)"/);
    if (orgSzMatch) {
      image.orgWidth = parseInt(orgSzMatch[1]) / 100;
      image.orgHeight = parseInt(orgSzMatch[2]) / 100;
      // Use orgSz as fallback if neither sz nor curSz found
      if (!szMatch && !curSzMatch) {
        image.width = image.orgWidth;
        image.height = image.orgHeight;
      }
    }

    const picTagMatch = xml.match(/<hp:pic[^>]*>/);
    if (picTagMatch) {
      const picAttrs = picTagMatch[0];

      const zOrderMatch = picAttrs.match(/zOrder="(\d+)"/);
      if (zOrderMatch) image.zOrder = parseInt(zOrderMatch[1]);

      const numTypeMatch = picAttrs.match(/numberingType="([^"]*)"/);
      if (numTypeMatch) {
        const map: Record<string, import('./types').NumberingType> = {
          'NONE': 'none', 'PICTURE': 'picture', 'TABLE': 'table', 'EQUATION': 'equation'
        };
        image.numberingType = map[numTypeMatch[1].toUpperCase()] || 'none';
      }

      const textWrapMatch = picAttrs.match(/textWrap="([^"]*)"/);
      if (textWrapMatch) {
        const map: Record<string, import('./types').TextWrapType> = {
          'SQUARE': 'square', 'TIGHT': 'tight', 'THROUGH': 'through',
          'TOP_AND_BOTTOM': 'topAndBottom', 'BEHIND_TEXT': 'behindText', 'IN_FRONT_OF_TEXT': 'inFrontOfText'
        };
        image.textWrap = map[textWrapMatch[1].toUpperCase()] || 'square';
      }

      const textFlowMatch = picAttrs.match(/textFlow="([^"]*)"/);
      if (textFlowMatch) {
        const map: Record<string, import('./types').TextFlowType> = {
          'BOTH_SIDES': 'bothSides', 'LEFT_ONLY': 'leftOnly', 'RIGHT_ONLY': 'rightOnly', 'LARGEST_ONLY': 'largestOnly'
        };
        image.textFlow = map[textFlowMatch[1].toUpperCase()] || 'bothSides';
      }
    }

    const posMatch = xml.match(/<hp:pos[^>]*>/);
    if (posMatch) {
      const pos = posMatch[0];
      image.position = {};

      if (pos.includes('treatAsChar="1"') || pos.includes('treatAsChar="true"')) {
        image.position.treatAsChar = true;
      }
      if (pos.includes('affectLSpacing="1"') || pos.includes('affectLSpacing="true"')) {
        image.position.affectLSpacing = true;
      }
      if (pos.includes('flowWithText="1"') || pos.includes('flowWithText="true"')) {
        image.position.flowWithText = true;
      }
      if (pos.includes('allowOverlap="1"') || pos.includes('allowOverlap="true"')) {
        image.position.allowOverlap = true;
      }
      if (pos.includes('holdAnchorAndSO="1"') || pos.includes('holdAnchorAndSO="true"')) {
        image.position.holdAnchorAndSO = true;
      }

      const vertRelMatch = pos.match(/vertRelTo="([^"]*)"/);
      if (vertRelMatch) {
        const map: Record<string, import('./types').VertRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'PARA': 'para'
        };
        image.position.vertRelTo = map[vertRelMatch[1].toUpperCase()];
      }

      const horzRelMatch = pos.match(/horzRelTo="([^"]*)"/);
      if (horzRelMatch) {
        const map: Record<string, import('./types').HorzRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'COLUMN': 'column', 'PARA': 'para'
        };
        image.position.horzRelTo = map[horzRelMatch[1].toUpperCase()];
      }

      const vertAlignMatch = pos.match(/vertAlign="([^"]*)"/);
      if (vertAlignMatch) {
        const map: Record<string, import('./types').VertAlign> = {
          'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom', 'INSIDE': 'inside', 'OUTSIDE': 'outside'
        };
        image.position.vertAlign = map[vertAlignMatch[1].toUpperCase()];
      }

      const horzAlignMatch = pos.match(/horzAlign="([^"]*)"/);
      if (horzAlignMatch) {
        const map: Record<string, import('./types').HorzAlign> = {
          'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'INSIDE': 'inside', 'OUTSIDE': 'outside'
        };
        image.position.horzAlign = map[horzAlignMatch[1].toUpperCase()];
      }

      const vertOffsetMatch = pos.match(/vertOffset="(-?\d+)"/);
      if (vertOffsetMatch) {
        image.position.vertOffset = parseInt(vertOffsetMatch[1]) / 100;
      }

      const horzOffsetMatch = pos.match(/horzOffset="(-?\d+)"/);
      if (horzOffsetMatch) {
        image.position.horzOffset = parseInt(horzOffsetMatch[1]) / 100;
      }
    }

    const outMarginMatch = xml.match(/<hp:outMargin[^>]*left="(\d+)"[^>]*right="(\d+)"[^>]*top="(\d+)"[^>]*bottom="(\d+)"/);
    if (outMarginMatch) {
      image.outMargin = {
        left: parseInt(outMarginMatch[1]) / 100,
        right: parseInt(outMarginMatch[2]) / 100,
        top: parseInt(outMarginMatch[3]) / 100,
        bottom: parseInt(outMarginMatch[4]) / 100
      };
    }

    const flipMatch = xml.match(/<hc:flip[^>]*horizontal="([^"]*)"[^>]*vertical="([^"]*)"/);
    if (flipMatch) {
      image.flip = {
        horizontal: flipMatch[1] === '1' || flipMatch[1] === 'true',
        vertical: flipMatch[2] === '1' || flipMatch[2] === 'true'
      };
    }

    const rotationMatch = xml.match(/<hp:rotationInfo[^>]*angle="(-?\d+)"(?:[^>]*centerX="(\d+)")?(?:[^>]*centerY="(\d+)")?/);
    if (rotationMatch) {
      image.rotation = {
        angle: parseInt(rotationMatch[1]),
        centerX: rotationMatch[2] ? parseInt(rotationMatch[2]) / 100 : undefined,
        centerY: rotationMatch[3] ? parseInt(rotationMatch[3]) / 100 : undefined
      };
    }

    const imgEffectMatch = xml.match(/<hc:imgEffect[^>]*>/);
    if (imgEffectMatch) {
      const effect = imgEffectMatch[0];
      
      const brightnessMatch = effect.match(/brightness="(-?\d+)"/);
      if (brightnessMatch) {
        image.brightness = parseInt(brightnessMatch[1]);
      }

      const contrastMatch = effect.match(/contrast="(-?\d+)"/);
      if (contrastMatch) {
        image.contrast = parseInt(contrastMatch[1]);
      }
    }

    const alphaMatch = xml.match(/<hc:img[^>]*alpha="(\d+)"/);
    if (alphaMatch) {
      image.alpha = parseInt(alphaMatch[1]) / 255;
    }

    const shapeCommentMatch = xml.match(/<hp:shapeComment>([^<]*)<\/hp:shapeComment>/);
    if (shapeCommentMatch) {
      image.shapeComment = shapeCommentMatch[1];
    }

    // Update content.images with parsed width/height
    // This ensures getImages() returns correct dimensions
    content.images.set(imageId, image);

    return image;
  }

  private static decodeXmlEntities(text: string): string {
    return text
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  static async updateZip(zip: JSZip, content: HwpxContent): Promise<void> {
    for (let sectionIndex = 0; sectionIndex < content.sections.length; sectionIndex++) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const existingXml = await this.readXmlFile(zip, sectionPath);

      if (existingXml) {
        const updatedXml = this.updateSectionXml(existingXml, content.sections[sectionIndex]);
        zip.file(sectionPath, updatedXml);
      }
    }
  }

  private static updateSectionXml(xml: string, section: HwpxSection): string {
    let updatedXml = xml;
    let elementIndex = 0;

    const paragraphElements = section.elements.filter((e) => e.type === 'paragraph');

    const paragraphRegex = /<hp:p[^>]*>([\s\S]*?)<\/hp:p>/g;

    updatedXml = xml.replace(paragraphRegex, (match) => {
      if (elementIndex < paragraphElements.length) {
        const paragraph = paragraphElements[elementIndex].data as HwpxParagraph;
        elementIndex++;
        return this.updateParagraphXml(match, paragraph);
      }
      elementIndex++;
      return match;
    });

    return updatedXml;
  }

  private static updateParagraphXml(xml: string, paragraph: HwpxParagraph): string {
    const fullText = paragraph.runs.map((r) => r.text).join('');
    const textTagRegex = /(<hp:t[^>]*>)[^<]*(<\/hp:t>)/;

    if (textTagRegex.test(xml)) {
      return xml.replace(textTagRegex, `$1${this.escapeXml(fullText)}$2`);
    }

    return xml;
  }

  private static escapeXml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  private static parseLine(xml: string): HwpxLine {
    const x1 = this.parseNumber(xml, /x1="([^"]*)"/) || this.parseNumber(xml, /startX="([^"]*)"/);
    const y1 = this.parseNumber(xml, /y1="([^"]*)"/) || this.parseNumber(xml, /startY="([^"]*)"/);
    const x2 = this.parseNumber(xml, /x2="([^"]*)"/) || this.parseNumber(xml, /endX="([^"]*)"/);
    const y2 = this.parseNumber(xml, /y2="([^"]*)"/) || this.parseNumber(xml, /endY="([^"]*)"/);
    const strokeColor = xml.match(/(?:stroke|lineColor)="([^"]*)"/)?.[1];
    const strokeWidth = this.parseNumber(xml, /(?:strokeWidth|lineWidth)="([^"]*)"/);
    
    return {
      id: generateId(),
      x1: x1 || 0,
      y1: y1 || 0,
      x2: x2 || 100,
      y2: y2 || 0,
      strokeColor: strokeColor || '#000000',
      strokeWidth: strokeWidth || 1,
      strokeStyle: 'solid',
    };
  }

  private static parseRect(xml: string): HwpxRect {
    const x = this.parseNumber(xml, /(?:x|left)="([^"]*)"/);
    const y = this.parseNumber(xml, /(?:y|top)="([^"]*)"/);
    const width = this.parseNumber(xml, /width="([^"]*)"/);
    const height = this.parseNumber(xml, /height="([^"]*)"/);
    const fillColor = xml.match(/(?:fill|fillColor)="([^"]*)"/)?.[1];
    const strokeColor = xml.match(/(?:stroke|lineColor)="([^"]*)"/)?.[1];
    const cornerRadius = this.parseNumber(xml, /(?:rx|cornerRadius)="([^"]*)"/);

    return {
      id: generateId(),
      x: x || 0,
      y: y || 0,
      width: width || 100,
      height: height || 50,
      fillColor,
      strokeColor: strokeColor || '#000000',
      strokeWidth: 1,
      cornerRadius,
    };
  }

  private static parseEllipse(xml: string): HwpxEllipse {
    const cx = this.parseNumber(xml, /(?:cx|centerX)="([^"]*)"/);
    const cy = this.parseNumber(xml, /(?:cy|centerY)="([^"]*)"/);
    const rx = this.parseNumber(xml, /(?:rx|radiusX)="([^"]*)"/);
    const ry = this.parseNumber(xml, /(?:ry|radiusY)="([^"]*)"/);
    const fillColor = xml.match(/(?:fill|fillColor)="([^"]*)"/)?.[1];
    const strokeColor = xml.match(/(?:stroke|lineColor)="([^"]*)"/)?.[1];

    return {
      id: generateId(),
      cx: cx || 50,
      cy: cy || 50,
      rx: rx || 50,
      ry: ry || 50,
      fillColor,
      strokeColor: strokeColor || '#000000',
      strokeWidth: 1,
    };
  }

  private static parseTextBox(xml: string): HwpxTextBox {
    const x = this.parseNumber(xml, /(?:x|left)="([^"]*)"/);
    const y = this.parseNumber(xml, /(?:y|top)="([^"]*)"/);
    const width = this.parseNumber(xml, /width="([^"]*)"/);
    const height = this.parseNumber(xml, /height="([^"]*)"/);
    const fillColor = xml.match(/(?:fill|fillColor)="([^"]*)"/)?.[1];
    const strokeColor = xml.match(/(?:stroke|lineColor)="([^"]*)"/)?.[1];

    const paragraphs: HwpxParagraph[] = [];
    const paragraphRegex = /<hp:p[^>]*>([\s\S]*?)<\/hp:p>/g;
    let match;
    while ((match = paragraphRegex.exec(xml)) !== null) {
      paragraphs.push(this.parseParagraph(match[0]));
    }

    return {
      id: generateId(),
      x: x || 0,
      y: y || 0,
      width: width || 200,
      height: height || 100,
      paragraphs,
      fillColor,
      strokeColor,
      strokeWidth: strokeColor ? 1 : 0,
    };
  }

  private static parseHorizontalRules(xml: string, section: HwpxSection): void {
    const hrOnlyPatterns = [
      /^[\s]*[─]{10,}[\s]*$/,
      /^[\s]*[━]{10,}[\s]*$/,
      /^[\s]*[═]{10,}[\s]*$/,
      /^[\s]*[▬]{10,}[\s]*$/,
      /^[\s]*[-]{20,}[\s]*$/,
    ];

    for (let i = 0; i < section.elements.length; i++) {
      const el = section.elements[i];
      if (el.type === 'paragraph') {
        const text = el.data.runs.map(r => r.text).join('').trim();
        const isHrOnly = hrOnlyPatterns.some(pattern => pattern.test(text));
        if (isHrOnly) {
          section.elements[i] = {
            type: 'hr',
            data: {
              id: generateId(),
              width: 'full',
              height: 1,
              color: '#000000',
              style: 'solid',
              align: 'center',
            },
          };
        }
      }
    }
  }

  private static parseNumber(xml: string, regex: RegExp): number | undefined {
    const match = xml.match(regex);
    if (match) {
      const val = parseFloat(match[1]);
      return isNaN(val) ? undefined : val / 100;
    }
    return undefined;
  }

  private static parseShapeObject(xml: string): ShapeObject | undefined {
    const szMatch = xml.match(/<hp:sz\s+width="(\d+)"[^>]*height="(\d+)"/);
    const posMatch = xml.match(/<hp:pos[^>]*>/);
    
    if (!szMatch && !posMatch) return undefined;
    
    const shapeObject: ShapeObject = {};
    
    if (szMatch) {
      shapeObject.size = {
        width: parseInt(szMatch[1]) / 100,
        height: parseInt(szMatch[2]) / 100,
      };
    }
    
    if (posMatch) {
      const pos = posMatch[0];
      shapeObject.position = {};
      
      if (pos.includes('treatAsChar="1"') || pos.includes('treatAsChar="true"')) {
        shapeObject.position.treatAsChar = true;
      }
      
      const vertRelMatch = pos.match(/vertRelTo="([^"]*)"/);
      if (vertRelMatch) {
        const map: Record<string, import('./types').VertRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'PARA': 'para'
        };
        shapeObject.position.vertRelTo = map[vertRelMatch[1].toUpperCase()];
      }
      
      const horzRelMatch = pos.match(/horzRelTo="([^"]*)"/);
      if (horzRelMatch) {
        const map: Record<string, import('./types').HorzRelTo> = {
          'PAPER': 'paper', 'PAGE': 'page', 'COLUMN': 'column', 'PARA': 'para'
        };
        shapeObject.position.horzRelTo = map[horzRelMatch[1].toUpperCase()];
      }
      
      const vertOffsetMatch = pos.match(/vertOffset="(-?\d+)"/);
      if (vertOffsetMatch) {
        shapeObject.position.vertOffset = parseInt(vertOffsetMatch[1]) / 100;
      }
      
      const horzOffsetMatch = pos.match(/horzOffset="(-?\d+)"/);
      if (horzOffsetMatch) {
        shapeObject.position.horzOffset = parseInt(horzOffsetMatch[1]) / 100;
      }
    }
    
    const instIdMatch = xml.match(/instId="([^"]*)"/);
    if (instIdMatch) shapeObject.instId = instIdMatch[1];
    
    const zOrderMatch = xml.match(/zOrder="(\d+)"/);
    if (zOrderMatch) shapeObject.zOrder = parseInt(zOrderMatch[1]);
    
    return shapeObject;
  }

  private static parseDrawingObject(xml: string): DrawingObject | undefined {
    const drawingObject: DrawingObject = {};
    
    const lineShapeMatch = xml.match(/<hc:lineShape[^>]*(?:\/>|>([\s\S]*?)<\/hc:lineShape>)/);
    if (lineShapeMatch) {
      const content = lineShapeMatch[0];
      drawingObject.lineShape = {
        color: content.match(/color="([^"]*)"/)?.[1],
        width: this.parseNumber(content, /width="([^"]*)"/),
      };
      
      const styleMatch = content.match(/style="([^"]*)"/);
      if (styleMatch) {
        const styleMap: Record<string, import('./types').LineType1> = {
          'SOLID': 'Solid', 'DASH': 'Dash', 'DOT': 'Dot', 'DASH_DOT': 'DashDot'
        };
        drawingObject.lineShape.style = styleMap[styleMatch[1].toUpperCase()] || 'Solid';
      }
    }
    
    const fillBrushMatch = xml.match(/<hc:fillBrush[^>]*>([\s\S]*?)<\/hc:fillBrush>/);
    if (fillBrushMatch) {
      drawingObject.fillBrush = {};
      const fillContent = fillBrushMatch[1];
      
      const winBrushMatch = fillContent.match(/<hc:winBrush[^>]*faceColor="([^"]*)"/);
      if (winBrushMatch) {
        drawingObject.fillBrush.windowBrush = { faceColor: winBrushMatch[1] };
      }
    }
    
    return Object.keys(drawingObject).length > 0 ? drawingObject : undefined;
  }

  private static parseArc(xml: string): HwpxArc {
    const arc: HwpxArc = {
      id: generateId(),
      centerX: 0,
      centerY: 0,
    };
    
    const typeMatch = xml.match(/\btype="([^"]*)"/);
    if (typeMatch) {
      const typeMap: Record<string, ArcType> = {
        'NORMAL': 'Normal', 'PIE': 'Pie', 'CHORD': 'Chord'
      };
      arc.type = typeMap[typeMatch[1].toUpperCase()] || 'Normal';
    }
    
    const centerXMatch = xml.match(/centerX="(-?\d+)"/);
    if (centerXMatch) arc.centerX = parseInt(centerXMatch[1]) / 100;
    
    const centerYMatch = xml.match(/centerY="(-?\d+)"/);
    if (centerYMatch) arc.centerY = parseInt(centerYMatch[1]) / 100;
    
    const axis1XMatch = xml.match(/axis1X="(-?\d+)"/);
    if (axis1XMatch) arc.axis1X = parseInt(axis1XMatch[1]) / 100;
    
    const axis1YMatch = xml.match(/axis1Y="(-?\d+)"/);
    if (axis1YMatch) arc.axis1Y = parseInt(axis1YMatch[1]) / 100;
    
    const axis2XMatch = xml.match(/axis2X="(-?\d+)"/);
    if (axis2XMatch) arc.axis2X = parseInt(axis2XMatch[1]) / 100;
    
    const axis2YMatch = xml.match(/axis2Y="(-?\d+)"/);
    if (axis2YMatch) arc.axis2Y = parseInt(axis2YMatch[1]) / 100;
    
    arc.shapeObject = this.parseShapeObject(xml);
    arc.drawingObject = this.parseDrawingObject(xml);
    
    return arc;
  }

  private static parsePolygon(xml: string): HwpxPolygon {
    const polygon: HwpxPolygon = {
      id: generateId(),
      points: [],
    };
    
    const pointRegex = /<(?:hp:|hc:)?pt[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/gi;
    let pointMatch;
    while ((pointMatch = pointRegex.exec(xml)) !== null) {
      polygon.points.push({
        x: parseInt(pointMatch[1]) / 100,
        y: parseInt(pointMatch[2]) / 100,
      });
    }
    
    if (polygon.points.length === 0) {
      const altPointRegex = /<(?:hp:|hc:)?point[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/gi;
      while ((pointMatch = altPointRegex.exec(xml)) !== null) {
        polygon.points.push({
          x: parseInt(pointMatch[1]) / 100,
          y: parseInt(pointMatch[2]) / 100,
        });
      }
    }
    
    polygon.shapeObject = this.parseShapeObject(xml);
    polygon.drawingObject = this.parseDrawingObject(xml);
    
    return polygon;
  }

  private static parseCurve(xml: string): HwpxCurve {
    const curve: HwpxCurve = {
      id: generateId(),
      segments: [],
    };
    
    const segmentRegex = /<(?:hp:|hc:)?seg[^>]*type="([^"]*)"[^>]*x1="(-?\d+)"[^>]*y1="(-?\d+)"[^>]*x2="(-?\d+)"[^>]*y2="(-?\d+)"/gi;
    let segMatch;
    while ((segMatch = segmentRegex.exec(xml)) !== null) {
      const segment: CurveSegment = {
        type: segMatch[1].toUpperCase() === 'CURVE' ? 'Curve' : 'Line',
        x1: parseInt(segMatch[2]) / 100,
        y1: parseInt(segMatch[3]) / 100,
        x2: parseInt(segMatch[4]) / 100,
        y2: parseInt(segMatch[5]) / 100,
      };
      curve.segments.push(segment);
    }
    
    curve.shapeObject = this.parseShapeObject(xml);
    curve.drawingObject = this.parseDrawingObject(xml);
    
    return curve;
  }

  private static parseConnectLine(xml: string): HwpxConnectLine {
    const connectLine: HwpxConnectLine = {
      id: generateId(),
    };
    
    const typeMatch = xml.match(/\btype="([^"]*)"/);
    if (typeMatch) connectLine.type = typeMatch[1];
    
    const startXMatch = xml.match(/startX="(-?\d+)"/);
    if (startXMatch) connectLine.startX = parseInt(startXMatch[1]) / 100;
    
    const startYMatch = xml.match(/startY="(-?\d+)"/);
    if (startYMatch) connectLine.startY = parseInt(startYMatch[1]) / 100;
    
    const endXMatch = xml.match(/endX="(-?\d+)"/);
    if (endXMatch) connectLine.endX = parseInt(endXMatch[1]) / 100;
    
    const endYMatch = xml.match(/endY="(-?\d+)"/);
    if (endYMatch) connectLine.endY = parseInt(endYMatch[1]) / 100;
    
    const startSubjectIDMatch = xml.match(/startSubjectID="([^"]*)"/);
    if (startSubjectIDMatch) connectLine.startSubjectID = startSubjectIDMatch[1];
    
    const startSubjectIndexMatch = xml.match(/startSubjectIndex="(\d+)"/);
    if (startSubjectIndexMatch) connectLine.startSubjectIndex = parseInt(startSubjectIndexMatch[1]);
    
    const endSubjectIDMatch = xml.match(/endSubjectID="([^"]*)"/);
    if (endSubjectIDMatch) connectLine.endSubjectID = endSubjectIDMatch[1];
    
    const endSubjectIndexMatch = xml.match(/endSubjectIndex="(\d+)"/);
    if (endSubjectIndexMatch) connectLine.endSubjectIndex = parseInt(endSubjectIndexMatch[1]);
    
    connectLine.shapeObject = this.parseShapeObject(xml);
    connectLine.drawingObject = this.parseDrawingObject(xml);
    
    return connectLine;
  }

  private static parseContainer(xml: string, content: HwpxContent): HwpxContainer {
    const container: HwpxContainer = {
      id: generateId(),
      children: [],
    };
    
    container.shapeObject = this.parseShapeObject(xml);
    
    const lineMatches = xml.matchAll(/<hp:line\b[^>]*(?:\/>|>[\s\S]*?<\/hp:line>)/g);
    for (const match of lineMatches) {
      container.children.push(this.parseLine(match[0]));
    }
    
    const rectMatches = xml.matchAll(/<hp:rect\b[^>]*(?:\/>|>[\s\S]*?<\/hp:rect>)/g);
    for (const match of rectMatches) {
      container.children.push(this.parseRect(match[0]));
    }
    
    const ellipseMatches = xml.matchAll(/<hp:ellipse\b[^>]*(?:\/>|>[\s\S]*?<\/hp:ellipse>)/g);
    for (const match of ellipseMatches) {
      container.children.push(this.parseEllipse(match[0]));
    }
    
    const arcMatches = xml.matchAll(/<hp:arc\b[^>]*(?:\/>|>[\s\S]*?<\/hp:arc>)/g);
    for (const match of arcMatches) {
      container.children.push(this.parseArc(match[0]));
    }
    
    const polygonMatches = xml.matchAll(/<hp:polygon\b[^>]*(?:\/>|>[\s\S]*?<\/hp:polygon>)/g);
    for (const match of polygonMatches) {
      container.children.push(this.parsePolygon(match[0]));
    }
    
    const curveMatches = xml.matchAll(/<hp:curve\b[^>]*(?:\/>|>[\s\S]*?<\/hp:curve>)/g);
    for (const match of curveMatches) {
      container.children.push(this.parseCurve(match[0]));
    }
    
    const picMatches = xml.matchAll(/<hp:pic\b[^>]*>[\s\S]*?<\/hp:pic>/g);
    for (const match of picMatches) {
      const image = this.parseImageElement(match[0], content);
      if (image) container.children.push(image);
    }
    
    const nestedContainerMatches = xml.matchAll(/<hp:container\b[^>]*>[\s\S]*?<\/hp:container>/g);
    for (const match of nestedContainerMatches) {
      if (match[0] !== xml) {
        container.children.push(this.parseContainer(match[0], content));
      }
    }
    
    return container;
  }

  private static parseOle(xml: string): HwpxOle {
    const ole: HwpxOle = {
      id: generateId(),
    };
    
    const objectTypeMatch = xml.match(/objectType="([^"]*)"/);
    if (objectTypeMatch) {
      const typeMap: Record<string, OleObjectType> = {
        'UNKNOWN': 'Unknown', 'EMBEDDED': 'Embedded', 'LINK': 'Link',
        'STATIC': 'Static', 'EQUATION': 'Equation'
      };
      ole.objectType = typeMap[objectTypeMatch[1].toUpperCase()] || 'Unknown';
    }
    
    const extentXMatch = xml.match(/extentX="(\d+)"/);
    if (extentXMatch) ole.extentX = parseInt(extentXMatch[1]) / 100;
    
    const extentYMatch = xml.match(/extentY="(\d+)"/);
    if (extentYMatch) ole.extentY = parseInt(extentYMatch[1]) / 100;
    
    const binItemMatch = xml.match(/binaryItemIDRef="([^"]*)"/);
    if (binItemMatch) ole.binItem = binItemMatch[1];
    
    const drawAspectMatch = xml.match(/drawAspect="([^"]*)"/);
    if (drawAspectMatch) {
      const aspectMap: Record<string, DrawAspect> = {
        'CONTENT': 'Content', 'THUMBNAIL': 'ThumbNail', 'ICON': 'Icon', 'DOCPRINT': 'DocPrint'
      };
      ole.drawAspect = aspectMap[drawAspectMatch[1].toUpperCase()] || 'Content';
    }
    
    const hasMonikerMatch = xml.match(/hasMoniker="([^"]*)"/);
    if (hasMonikerMatch) {
      ole.hasMoniker = hasMonikerMatch[1] === '1' || hasMonikerMatch[1] === 'true';
    }
    
    const eqBaseLineMatch = xml.match(/eqBaseLine="(-?\d+)"/);
    if (eqBaseLineMatch) ole.eqBaseLine = parseInt(eqBaseLineMatch[1]) / 100;
    
    ole.shapeObject = this.parseShapeObject(xml);
    
    return ole;
  }

  private static parseEquation(xml: string): HwpxEquation {
    const equation: HwpxEquation = {
      id: generateId(),
    };
    
    const lineModeMatch = xml.match(/lineMode="([^"]*)"/);
    if (lineModeMatch) {
      equation.lineMode = lineModeMatch[1] === '1' || lineModeMatch[1] === 'true';
    }
    
    const baseUnitMatch = xml.match(/baseUnit="(\d+)"/);
    if (baseUnitMatch) equation.baseUnit = parseInt(baseUnitMatch[1]);
    
    const textColorMatch = xml.match(/textColor="([^"]*)"/);
    if (textColorMatch) equation.textColor = textColorMatch[1];
    
    const baseLineMatch = xml.match(/baseLine="(-?\d+)"/);
    if (baseLineMatch) equation.baseLine = parseInt(baseLineMatch[1]) / 100;
    
    const versionMatch = xml.match(/version="([^"]*)"/);
    if (versionMatch) equation.version = versionMatch[1];
    
    const scriptMatch = xml.match(/<hp:script[^>]*>([^<]*)<\/hp:script>/i);
    if (scriptMatch) {
      equation.script = this.decodeXmlEntities(scriptMatch[1]);
    }
    
    equation.shapeObject = this.parseShapeObject(xml);
    
    return equation;
  }

  private static parseTextArt(xml: string): HwpxTextArt {
    const textArt: HwpxTextArt = {
      id: generateId(),
    };
    
    const textMatch = xml.match(/<hp:textArt[^>]*>[\s\S]*?<hp:text>([^<]*)<\/hp:text>/i);
    if (textMatch) {
      textArt.text = this.decodeXmlEntities(textMatch[1]);
    }
    
    const x0Match = xml.match(/x0="(-?\d+)"/);
    if (x0Match) textArt.x0 = parseInt(x0Match[1]) / 100;
    
    const y0Match = xml.match(/y0="(-?\d+)"/);
    if (y0Match) textArt.y0 = parseInt(y0Match[1]) / 100;
    
    const x1Match = xml.match(/x1="(-?\d+)"/);
    if (x1Match) textArt.x1 = parseInt(x1Match[1]) / 100;
    
    const y1Match = xml.match(/y1="(-?\d+)"/);
    if (y1Match) textArt.y1 = parseInt(y1Match[1]) / 100;
    
    const x2Match = xml.match(/x2="(-?\d+)"/);
    if (x2Match) textArt.x2 = parseInt(x2Match[1]) / 100;
    
    const y2Match = xml.match(/y2="(-?\d+)"/);
    if (y2Match) textArt.y2 = parseInt(y2Match[1]) / 100;
    
    const x3Match = xml.match(/x3="(-?\d+)"/);
    if (x3Match) textArt.x3 = parseInt(x3Match[1]) / 100;
    
    const y3Match = xml.match(/y3="(-?\d+)"/);
    if (y3Match) textArt.y3 = parseInt(y3Match[1]) / 100;
    
    const shapeMatch = xml.match(/<hp:textArtShape[^>]*>([\s\S]*?)<\/hp:textArtShape>/i);
    if (shapeMatch) {
      textArt.shape = {};
      const shapeContent = shapeMatch[0];
      
      const fontNameMatch = shapeContent.match(/fontName="([^"]*)"/);
      if (fontNameMatch) textArt.shape.fontName = fontNameMatch[1];
      
      const fontStyleMatch = shapeContent.match(/fontStyle="([^"]*)"/);
      if (fontStyleMatch) textArt.shape.fontStyle = fontStyleMatch[1];
      
      const textShapeMatch = shapeContent.match(/textShape="(\d+)"/);
      if (textShapeMatch) textArt.shape.textShape = parseInt(textShapeMatch[1]);
      
      const lineSpacingMatch = shapeContent.match(/lineSpacing="(\d+)"/);
      if (lineSpacingMatch) textArt.shape.lineSpacing = parseInt(lineSpacingMatch[1]);
      
      const charSpacingMatch = shapeContent.match(/charSpacing="(-?\d+)"/);
      if (charSpacingMatch) textArt.shape.charSpacing = parseInt(charSpacingMatch[1]);
    }
    
    const outlineDataMatch = xml.match(/<hp:outlineData[^>]*>([\s\S]*?)<\/hp:outlineData>/i);
    if (outlineDataMatch) {
      textArt.outlineData = [];
      const pointRegex = /<(?:hp:|hc:)?pt[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/gi;
      let pointMatch;
      while ((pointMatch = pointRegex.exec(outlineDataMatch[1])) !== null) {
        textArt.outlineData.push({
          x: parseInt(pointMatch[1]) / 100,
          y: parseInt(pointMatch[2]) / 100,
        });
      }
    }
    
    return textArt;
  }

  private static parseUnknownObject(xml: string): HwpxUnknownObject {
    const unknownObj: HwpxUnknownObject = {
      id: generateId(),
    };
    
    const ctrlIdMatch = xml.match(/ctrlId="([^"]*)"/);
    if (ctrlIdMatch) unknownObj.ctrlId = ctrlIdMatch[1];
    
    const x0Match = xml.match(/x0="(-?\d+)"/);
    if (x0Match) unknownObj.x0 = parseInt(x0Match[1]) / 100;
    
    const y0Match = xml.match(/y0="(-?\d+)"/);
    if (y0Match) unknownObj.y0 = parseInt(y0Match[1]) / 100;
    
    const x1Match = xml.match(/x1="(-?\d+)"/);
    if (x1Match) unknownObj.x1 = parseInt(x1Match[1]) / 100;
    
    const y1Match = xml.match(/y1="(-?\d+)"/);
    if (y1Match) unknownObj.y1 = parseInt(y1Match[1]) / 100;
    
    const x2Match = xml.match(/x2="(-?\d+)"/);
    if (x2Match) unknownObj.x2 = parseInt(x2Match[1]) / 100;
    
    const y2Match = xml.match(/y2="(-?\d+)"/);
    if (y2Match) unknownObj.y2 = parseInt(y2Match[1]) / 100;
    
    const x3Match = xml.match(/x3="(-?\d+)"/);
    if (x3Match) unknownObj.x3 = parseInt(x3Match[1]) / 100;
    
    const y3Match = xml.match(/y3="(-?\d+)"/);
    if (y3Match) unknownObj.y3 = parseInt(y3Match[1]) / 100;
    
    unknownObj.shapeObject = this.parseShapeObject(xml);
    unknownObj.drawingObject = this.parseDrawingObject(xml);
    
    return unknownObj;
  }

  private static parseFormObject(xml: string): import('./types').FormObject {
    const formObject: import('./types').FormObject = {};
    
    const nameMatch = xml.match(/\bname="([^"]*)"/);
    if (nameMatch) formObject.name = nameMatch[1];
    
    const foreColorMatch = xml.match(/foreColor="([^"]*)"/);
    if (foreColorMatch) formObject.foreColor = foreColorMatch[1];
    
    const backColorMatch = xml.match(/backColor="([^"]*)"/);
    if (backColorMatch) formObject.backColor = backColorMatch[1];
    
    const groupNameMatch = xml.match(/groupName="([^"]*)"/);
    if (groupNameMatch) formObject.groupName = groupNameMatch[1];
    
    const tabStopMatch = xml.match(/tabStop="([^"]*)"/);
    if (tabStopMatch) {
      formObject.tabStop = tabStopMatch[1] === '1' || tabStopMatch[1] === 'true';
    }
    
    const tabOrderMatch = xml.match(/tabOrder="(\d+)"/);
    if (tabOrderMatch) formObject.tabOrder = parseInt(tabOrderMatch[1]);
    
    const enabledMatch = xml.match(/enabled="([^"]*)"/);
    if (enabledMatch) {
      formObject.enabled = enabledMatch[1] === '1' || enabledMatch[1] === 'true';
    }
    
    const borderTypeMatch = xml.match(/borderType="(\d+)"/);
    if (borderTypeMatch) formObject.borderType = parseInt(borderTypeMatch[1]);
    
    const drawFrameMatch = xml.match(/drawFrame="([^"]*)"/);
    if (drawFrameMatch) {
      formObject.drawFrame = drawFrameMatch[1] === '1' || drawFrameMatch[1] === 'true';
    }
    
    const printableMatch = xml.match(/printable="([^"]*)"/);
    if (printableMatch) {
      formObject.printable = printableMatch[1] === '1' || printableMatch[1] === 'true';
    }
    
    const formCharShapeMatch = xml.match(/<(?:hp:|hc:)?formCharShape[^>]*>/i);
    if (formCharShapeMatch) {
      const fcs = formCharShapeMatch[0];
      formObject.formCharShape = {};
      
      const charShapeMatch = fcs.match(/charPrIDRef="(\d+)"/);
      if (charShapeMatch) formObject.formCharShape.charShape = parseInt(charShapeMatch[1]);
      
      const followContextMatch = fcs.match(/followContext="([^"]*)"/);
      if (followContextMatch) {
        formObject.formCharShape.followContext = followContextMatch[1] === '1' || followContextMatch[1] === 'true';
      }
      
      const autoSizeMatch = fcs.match(/autoSize="([^"]*)"/);
      if (autoSizeMatch) {
        formObject.formCharShape.autoSize = autoSizeMatch[1] === '1' || autoSizeMatch[1] === 'true';
      }
      
      const wordWrapMatch = fcs.match(/wordWrap="([^"]*)"/);
      if (wordWrapMatch) {
        formObject.formCharShape.wordWrap = wordWrapMatch[1] === '1' || wordWrapMatch[1] === 'true';
      }
    }
    
    const buttonSetMatch = xml.match(/<(?:hp:|hc:)?buttonSet[^>]*>([\s\S]*?)<\/(?:hp:|hc:)?buttonSet>/i);
    if (buttonSetMatch) {
      formObject.buttonSet = {};
      const bsContent = buttonSetMatch[0];
      
      const captionMatch = bsContent.match(/caption="([^"]*)"/);
      if (captionMatch) formObject.buttonSet.caption = captionMatch[1];
      
      const valueMatch = bsContent.match(/\bvalue="([^"]*)"/);
      if (valueMatch) formObject.buttonSet.value = valueMatch[1];
      
      const radioGroupNameMatch = bsContent.match(/radioGroupName="([^"]*)"/);
      if (radioGroupNameMatch) formObject.buttonSet.radioGroupName = radioGroupNameMatch[1];
      
      const triStateMatch = bsContent.match(/triState="([^"]*)"/);
      if (triStateMatch) {
        formObject.buttonSet.triState = triStateMatch[1] === '1' || triStateMatch[1] === 'true';
      }
      
      const backStyleMatch = bsContent.match(/backStyle="([^"]*)"/);
      if (backStyleMatch) formObject.buttonSet.backStyle = backStyleMatch[1];
    }
    
    return formObject;
  }

  private static parseButton(xml: string): HwpxButton {
    const button: HwpxButton = {
      id: generateId(),
    };
    
    button.shapeObject = this.parseShapeObject(xml);
    button.formObject = this.parseFormObject(xml);
    
    return button;
  }

  private static parseRadioButton(xml: string): HwpxRadioButton {
    const radioButton: HwpxRadioButton = {
      id: generateId(),
    };
    
    radioButton.shapeObject = this.parseShapeObject(xml);
    radioButton.formObject = this.parseFormObject(xml);
    
    return radioButton;
  }

  private static parseCheckButton(xml: string): HwpxCheckButton {
    const checkButton: HwpxCheckButton = {
      id: generateId(),
    };
    
    checkButton.shapeObject = this.parseShapeObject(xml);
    checkButton.formObject = this.parseFormObject(xml);
    
    return checkButton;
  }

  private static parseComboBox(xml: string): HwpxComboBox {
    const comboBox: HwpxComboBox = {
      id: generateId(),
    };
    
    const listBoxRowsMatch = xml.match(/listBoxRows="(\d+)"/);
    if (listBoxRowsMatch) comboBox.listBoxRows = parseInt(listBoxRowsMatch[1]);
    
    const listBoxWidthMatch = xml.match(/listBoxWidth="(\d+)"/);
    if (listBoxWidthMatch) comboBox.listBoxWidth = parseInt(listBoxWidthMatch[1]) / 100;
    
    const textMatch = xml.match(/\btext="([^"]*)"/);
    if (textMatch) comboBox.text = textMatch[1];
    
    const editEnableMatch = xml.match(/editEnable="([^"]*)"/);
    if (editEnableMatch) {
      comboBox.editEnable = editEnableMatch[1] === '1' || editEnableMatch[1] === 'true';
    }
    
    comboBox.shapeObject = this.parseShapeObject(xml);
    comboBox.formObject = this.parseFormObject(xml);
    
    return comboBox;
  }

  private static parseEdit(xml: string): HwpxEdit {
    const edit: HwpxEdit = {
      id: generateId(),
    };
    
    const multiLineMatch = xml.match(/multiLine="([^"]*)"/);
    if (multiLineMatch) {
      edit.multiLine = multiLineMatch[1] === '1' || multiLineMatch[1] === 'true';
    }
    
    const passwordCharMatch = xml.match(/passwordChar="([^"]*)"/);
    if (passwordCharMatch) edit.passwordChar = passwordCharMatch[1];
    
    const maxLengthMatch = xml.match(/maxLength="(\d+)"/);
    if (maxLengthMatch) edit.maxLength = parseInt(maxLengthMatch[1]);
    
    const scrollBarsMatch = xml.match(/scrollBars="([^"]*)"/);
    if (scrollBarsMatch) {
      edit.scrollBars = scrollBarsMatch[1] === '1' || scrollBarsMatch[1] === 'true';
    }
    
    const tabKeyBehaviorMatch = xml.match(/tabKeyBehavior="([^"]*)"/);
    if (tabKeyBehaviorMatch) edit.tabKeyBehavior = tabKeyBehaviorMatch[1];
    
    const numberMatch = xml.match(/\bnumber="([^"]*)"/);
    if (numberMatch) {
      edit.number = numberMatch[1] === '1' || numberMatch[1] === 'true';
    }
    
    const readOnlyMatch = xml.match(/readOnly="([^"]*)"/);
    if (readOnlyMatch) {
      edit.readOnly = readOnlyMatch[1] === '1' || readOnlyMatch[1] === 'true';
    }
    
    const alignTextMatch = xml.match(/alignText="([^"]*)"/);
    if (alignTextMatch) edit.alignText = alignTextMatch[1];
    
    const editTextMatch = xml.match(/<(?:hp:|hc:)?editText[^>]*>([^<]*)<\/(?:hp:|hc:)?editText>/i);
    if (editTextMatch) {
      edit.text = this.decodeXmlEntities(editTextMatch[1]);
    }
    
    edit.shapeObject = this.parseShapeObject(xml);
    edit.formObject = this.parseFormObject(xml);
    
    return edit;
  }

  private static parseListBox(xml: string): HwpxListBox {
    const listBox: HwpxListBox = {
      id: generateId(),
    };
    
    const textMatch = xml.match(/\btext="([^"]*)"/);
    if (textMatch) listBox.text = textMatch[1];
    
    const itemHeightMatch = xml.match(/itemHeight="(\d+)"/);
    if (itemHeightMatch) listBox.itemHeight = parseInt(itemHeightMatch[1]) / 100;
    
    const topIndexMatch = xml.match(/topIndex="(\d+)"/);
    if (topIndexMatch) listBox.topIndex = parseInt(topIndexMatch[1]);
    
    listBox.shapeObject = this.parseShapeObject(xml);
    listBox.formObject = this.parseFormObject(xml);
    
    return listBox;
  }

  private static parseScrollBar(xml: string): HwpxScrollBar {
    const scrollBar: HwpxScrollBar = {
      id: generateId(),
    };
    
    const delayMatch = xml.match(/delay="(\d+)"/);
    if (delayMatch) scrollBar.delay = parseInt(delayMatch[1]);
    
    const largeChangeMatch = xml.match(/largeChange="(\d+)"/);
    if (largeChangeMatch) scrollBar.largeChange = parseInt(largeChangeMatch[1]);
    
    const smallChangeMatch = xml.match(/smallChange="(\d+)"/);
    if (smallChangeMatch) scrollBar.smallChange = parseInt(smallChangeMatch[1]);
    
    const minMatch = xml.match(/\bmin="(\d+)"/);
    if (minMatch) scrollBar.min = parseInt(minMatch[1]);
    
    const maxMatch = xml.match(/\bmax="(\d+)"/);
    if (maxMatch) scrollBar.max = parseInt(maxMatch[1]);
    
    const pageMatch = xml.match(/\bpage="(\d+)"/);
    if (pageMatch) scrollBar.page = parseInt(pageMatch[1]);
    
    const valueMatch = xml.match(/\bvalue="(\d+)"/);
    if (valueMatch) scrollBar.value = parseInt(valueMatch[1]);
    
    const typeMatch = xml.match(/\btype="([^"]*)"/);
    if (typeMatch) scrollBar.type = typeMatch[1];
    
    scrollBar.shapeObject = this.parseShapeObject(xml);
    scrollBar.formObject = this.parseFormObject(xml);
    
    return scrollBar;
  }

  private static parseCompatibleDocument(xml: string): CompatibleDocument | undefined {
    const compatDocMatch = xml.match(/<hh:compatibleDocument[^>]*>([\s\S]*?)<\/hh:compatibleDocument>/i);
    if (!compatDocMatch) return undefined;
    
    const content = compatDocMatch[0];
    const compatDoc: CompatibleDocument = {};
    
    const targetProgramMatch = content.match(/targetProgram="([^"]*)"/);
    if (targetProgramMatch) {
      const progMap: Record<string, 'None' | 'Hwp70' | 'Word'> = {
        'NONE': 'None', 'HWP70': 'Hwp70', 'WORD': 'Word'
      };
      compatDoc.targetProgram = progMap[targetProgramMatch[1].toUpperCase()] || 'None';
    }
    
    const layoutCompatMatch = content.match(/<hh:layoutCompatibility[^>]*(?:\/>|([\s\S]*?)<\/hh:layoutCompatibility>)/i);
    if (layoutCompatMatch) {
      const lcContent = layoutCompatMatch[0];
      const lc: LayoutCompatibility = {};
      
      const boolFlags = [
        'applyFontWeightToBold', 'useInnerUnderline', 'fixedUnderlineWidth',
        'doNotApplyStrikeout', 'useLowercaseStrikeout', 'extendLineheightToOffset',
        'treatQuotationAsLatin', 'doNotAlignWhitespaceOnRight', 'doNotAdjustWordInJustify',
        'baseCharUnitOnEAsian', 'baseCharUnitOfIndentOnFirstChar', 'adjustLineheightToFont',
        'adjustBaselineInFixedLinespacing', 'excludeOverlappingParaSpacing',
        'applyNextspacingOfLastPara', 'applyAtLeastToPercent100Pct',
        'doNotApplyAutoSpaceEAsianEng', 'doNotApplyAutoSpaceEAsianNum',
        'adjustParaBorderfillToSpacing', 'connectParaBorderfillOfEqualBorder',
        'adjustParaBorderOffsetWithBorder', 'extendLineheightToParaBorderOffset',
        'applyParaBorderToOutside', 'baseLinespacingOnLinegrid', 'applyCharSpacingToCharGrid',
        'doNotApplyGridInHeaderfooter', 'extendHeaderfooterToBody',
        'adjustEndnotePositionToFootnote', 'doNotApplyImageEffect', 'doNotApplyShapeComment',
        'doNotAdjustEmptyAnchorLine', 'overlapBothAllowOverlap', 'doNotApplyVertOffsetOfForward',
        'extendVertLimitToPageMargins', 'doNotHoldAnchorOfTable', 'doNotFormattingAtBeneathAnchor',
        'doNotApplyExtensionCharCompose'
      ];
      
      for (const flag of boolFlags) {
        const regex = new RegExp(`${flag}="([^"]*)"`, 'i');
        const match = lcContent.match(regex);
        if (match) {
          (lc as Record<string, boolean>)[flag] = match[1] === '1' || match[1] === 'true';
        }
      }
      
      compatDoc.layoutCompatibility = lc;
    }
    
    return compatDoc;
  }

  private static async parseBinDataStorage(zip: JSZip, content: HwpxContent): Promise<void> {
    const binDataPath = 'Contents/content.hpf';
    const binDataXml = await this.readXmlFile(zip, binDataPath);
    
    if (binDataXml) {
      const binDataRegex = /<(?:hp:|hpf:)?binData[^>]*id="([^"]*)"[^>]*>([\s\S]*?)<\/(?:hp:|hpf:)?binData>/gi;
      let match;
      
      while ((match = binDataRegex.exec(binDataXml)) !== null) {
        const binData: BinData = {
          id: match[1],
          data: match[2].trim(),
        };
        
        const sizeMatch = match[0].match(/size="(\d+)"/);
        if (sizeMatch) binData.size = parseInt(sizeMatch[1]);
        
        const encodingMatch = match[0].match(/encoding="([^"]*)"/);
        if (encodingMatch && encodingMatch[1].toUpperCase() === 'BASE64') {
          binData.encoding = 'Base64';
        }
        
        const compressMatch = match[0].match(/compress="([^"]*)"/);
        if (compressMatch) {
          binData.compress = compressMatch[1] === '1' || compressMatch[1] === 'true';
        }
        
        content.binData.set(binData.id, binData);
      }
    }
    
    const binDataFolder = zip.folder('BinData');
    if (binDataFolder) {
      const binFiles = Object.keys(zip.files).filter(
        (f) => f.startsWith('BinData/') && !f.endsWith('/')
      );
      
      for (const binPath of binFiles) {
        const file = zip.file(binPath);
        if (!file) continue;
        
        const data = await file.async('base64');
        const fileName = binPath.split('/').pop() || '';
        const fileId = fileName.replace(/\.[^.]+$/, '');
        
        if (!content.binData.has(fileId)) {
          content.binData.set(fileId, {
            id: fileId,
            data: data,
            encoding: 'Base64',
          });
        }
      }
    }
  }
}
