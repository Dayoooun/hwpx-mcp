// ============================================================
// HWPML/HWPX Type Definitions
// Based on: 한글문서파일형식3.0_HWPML_revision1.2.md
// hwpunit: 10pt = 1000 hwpunit
// Colors: RGB as decimal (0x00bbggrr format)
// ============================================================

// ============================================================
// Section 2.2: Basic Attribute Value Types (기본 속성 값)
// ============================================================

// Line Type 1 (LineType1) - 선 종류 1
export type LineType1 =
  | 'Solid'           // 실선
  | 'Dash'            // 긴 점선
  | 'Dot'             // 점선
  | 'DashDot'         // -.-.-.-.
  | 'DashDotDot'      // -..-..-..
  | 'LongDash'        // Dash보다 긴 선분의 반복
  | 'CircleDot'       // 보다 큰 동그라미의 반복
  | 'DoubleSlim'      // 2중선
  | 'SlimThick'       // 가는 선 + 굵은 선 2중선
  | 'ThickSlim'       // 굵은 선 + 가는 선 2중선
  | 'SlimThickSlim'   // 가는 선 + 굵은 선 + 가는 선 3중선
  | 'None'            // 선 없음
  | 'solid' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot' | 'longDash' | 'circleDot'
  | 'doubleSlim' | 'slimThick' | 'thickSlim' | 'slimThickSlim' | 'none';

// Line Type 2 (LineType2) - same as LineType1 for most uses
export type LineType2 = LineType1;

// Line Type 3 (LineType3) - 선 종류 3 (simpler set)
export type LineType3 = 'Solid' | 'Dot' | 'Thick' | 'Dash' | 'DashDot' | 'DashDotDot';

// Line Width (LineWidth) - 선 두께
export type LineWidth =
  | '0.1mm' | '0.12mm' | '0.15mm' | '0.2mm' | '0.25mm'
  | '0.3mm' | '0.4mm' | '0.5mm' | '0.6mm' | '0.7mm'
  | '1.0mm' | '1.5mm' | '2.0mm' | '3.0mm' | '4.0mm' | '5.0mm';

// Number Type 1 (NumberType1) - 번호 모양 1
export type NumberType1 =
  | 'Digit'               // 1, 2, 3
  | 'CircledDigit'        // 동그라미 쳐진 1, 2, 3
  | 'RomanCapital'        // I, II, III
  | 'RomanSmall'          // i, ii, iii
  | 'LatinCapital'        // A, B, C
  | 'LatinSmall'          // a, b, c
  | 'CircledLatinCapital' // 동그라미 쳐진 A, B, C
  | 'CircledLatinSmall'   // 동그라미 쳐진 a, b, c
  | 'HangulSyllable'      // 가, 나, 다
  | 'CircledHangulSyllable' // 동그라미 쳐진 가, 나, 다
  | 'HangulJamo'          // ㄱ, ㄴ, ㄷ
  | 'CircledHangulJamo'   // 동그라미 쳐진 ㄱ, ㄴ, ㄷ
  | 'HangulPhonetic'      // 일, 이, 삼
  | 'Ideograph'           // 一, 二, 三
  | 'CircledIdeograph'    // 동그라미 쳐진 一, 二, 三
  | 'DecimalEnclosedInParentheses'
  | 'digit' | 'circledDigit' | 'romanCapital' | 'romanSmall' | 'latinCapital' | 'latinSmall'
  | 'circledLatinCapital' | 'circledLatinSmall' | 'hangulSyllable' | 'circledHangulSyllable'
  | 'hangulJamo' | 'circledHangulJamo' | 'hangulPhonetic' | 'ideograph' | 'circledIdeograph'
  | 'decimalEnclosedInParentheses';

// Number Type 2 - includes additional options
export type NumberType2 = NumberType1 | 'DecagonCircle' | 'DecagonCircleHanja' | 'Symbol' | 'UserChar';

export type AlignmentType1 = 'Justify' | 'Left' | 'Right' | 'Center' | 'Distribute' | 'DistributeSpace'
  | 'justify' | 'left' | 'right' | 'center' | 'distribute' | 'distributeSpace';

// Alignment Type 2 (AlignmentType2) - 정렬 방식 2
export type AlignmentType2 = 'Left' | 'Center' | 'Right';

// Arrow Type (ArrowType) - 화살표 시작/끝 모양
export type ArrowType =
  | 'Normal'        // 모양 없음
  | 'Arrow'         // 화살 모양
  | 'Spear'         // 작살 모양
  | 'ConcaveArrow'  // 오목한 화살모양
  | 'EmptyDiamond'  // 속이 빈 다이아몬드 모양
  | 'EmptyCircle'   // 속이 빈 원 모양
  | 'EmptyBox'      // 속이 빈 사각 모양
  | 'FilledDiamond' // 속이 채워진 다이아몬드 모양
  | 'FilledCircle'  // 속이 채워진 원 모양
  | 'FilledBox';    // 속이 채워진 사각 모양

// Arrow Size (ArrowSize) - 화살표 시작/끝 크기
export type ArrowSize =
  | 'SmallSmall' | 'SmallMedium' | 'SmallLarge'
  | 'MediumSmall' | 'MediumMedium' | 'MediumLarge'
  | 'LargeSmall' | 'LargeMedium' | 'LargeLarge';

// Language Type (LangType) - 언어 종류
export type LangType = 'Hangul' | 'Latin' | 'Hanja' | 'Japanese' | 'Other' | 'Symbol' | 'User';

// Hatch Style (HatchStyle) - 무늬 종류
export type HatchStyle = 'Horizontal' | 'Vertical' | 'BackSlash' | 'Slash' | 'Cross' | 'CrossDiagonal';

// Infill Mode (InfillMode) - 채우기 유형
export type InfillMode =
  | 'Tile' | 'TileHorzTop' | 'TileHorzBottom' | 'TileVertLeft' | 'TileVertRight'
  | 'Total' | 'Center' | 'CenterTop' | 'CenterBottom'
  | 'LeftCenter' | 'LeftTop' | 'LeftBottom'
  | 'RightCenter' | 'RightTop' | 'RightBottom' | 'Zoom';

// Line Wrap Type (LineWrapType) - 한줄로 입력
export type LineWrapType = 'Break' | 'Squeeze' | 'Keep' | 'break' | 'squeeze' | 'keep';

// Text Wrap Type (TextWrapType) - 글 배치
export type TextWrapType =
  | 'Square'        // bound rect를 따라
  | 'Tight'         // 오브젝트의 outline을 따라
  | 'Through'       // 오브젝트 내부의 빈 공간까지
  | 'TopAndBottom'  // 좌/우에는 텍스트를 배치하지 않음
  | 'BehindText'    // 글과 겹치게 하여 글 뒤로
  | 'InFrontOfText' // 글과 겹치게 하여 글 앞으로
  | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behindText' | 'inFrontOfText';

// Text Flow Type
export type TextFlowType = 'BothSides' | 'LeftOnly' | 'RightOnly' | 'LargestOnly'
  | 'bothSides' | 'leftOnly' | 'rightOnly' | 'largestOnly';

export type FieldType =
  | 'Clickhere' | 'Hyperlink' | 'Bookmark' | 'Formula' | 'Summery'
  | 'UserInfo' | 'Date' | 'DocDate' | 'Path' | 'Crossref' | 'Mailmerge' | 'Memo'
  | 'RevisionChange' | 'RevisionSign' | 'RevisionDelete' | 'RevisionAttach'
  | 'RevisionClipping' | 'RevisionSawtooth' | 'RevisionThinking' | 'RevisionPraise'
  | 'RevisionLine' | 'RevisionSimpleChange' | 'RevisionHyperlink'
  | 'RevisionLineAttach' | 'RevisionLineLink' | 'RevisionLineTransfer'
  | 'RevisionRightmove' | 'RevisionLeftmove' | 'RevisionTransfer' | 'RevisionSplit'
  | 'unknown' | 'hyperlink' | 'bookmark' | 'formula' | 'memo' | 'date' | 'docDate' | 'path' | 'mailMerge' | 'crossRef' | 'clickHere'
  | 'summary' | 'userInfo' | 'revisionSign' | 'privateTxt' | 'tableOfContents';

// Underline Type
export type UnderlineType = 'Bottom' | 'Center' | 'Top' | 'None' | 'bottom' | 'center' | 'top' | 'none';

// Strikeout Type
export type StrikeoutType = 'None' | 'Continuous';

// Shadow Type
export type ShadowType = 'None' | 'Drop' | 'Cont';

// Emphasis Mark (SymMark)
export type EmphasisMark = 'None' | 'Dot' | 'Circle' | 'Ring' | 'Caron' | 'UnderDot' | 'UnderLine' | 'Triangle'
  | 'none' | 'dot' | 'circle' | 'ring' | 'caron' | 'underDot' | 'underLine' | 'triangle';

export type HeadingType = 'None' | 'Outline' | 'Number' | 'Bullet' | 'none' | 'outline' | 'number' | 'bullet';

export type GradationType = 'Linear' | 'Radial' | 'Conical' | 'Square' | 'linear' | 'radial' | 'conical' | 'square';

export type PageBreakType = 'Table' | 'Cell' | 'None' | 'table' | 'cell' | 'none';

// Vertical Align
export type VertAlign = 'Top' | 'Center' | 'Bottom' | 'Inside' | 'Outside' | 'Para'
  | 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'para';

// Horizontal Align
export type HorzAlign = 'Left' | 'Center' | 'Right' | 'Inside' | 'Outside'
  | 'left' | 'center' | 'right' | 'inside' | 'outside';

// Vertical Relative To
export type VertRelTo = 'Paper' | 'Page' | 'Para' | 'paper' | 'page' | 'para';

// Horizontal Relative To
export type HorzRelTo = 'Paper' | 'Page' | 'Column' | 'Para' | 'paper' | 'page' | 'column' | 'para';

// Size Relative To
export type SizeRelTo = 'Paper' | 'Page' | 'Column' | 'Para' | 'Absolute';

// Gutter Type
export type GutterType = 'LeftOnly' | 'LeftRight' | 'TopBottom';

export type PageStartsOn = 'Both' | 'Even' | 'Odd' | 'both' | 'even' | 'odd';

// Arc Type
export type ArcType = 'Normal' | 'Pie' | 'Chord';

// OLE Object Type
export type OleObjectType = 'Unknown' | 'Embedded' | 'Link' | 'Static' | 'Equation';

// Draw Aspect
export type DrawAspect = 'Content' | 'ThumbNail' | 'Icon' | 'DocPrint';

// Column Type
export type ColumnType = 'Newspaper' | 'BalancedNewspaper' | 'Parallel' | 'newspaper' | 'balanced' | 'parallel';

// Column Layout
export type ColumnLayout = 'Left' | 'Right' | 'Mirror' | 'left' | 'right' | 'mirror';

// Column Info
export interface ColumnInfo {
  width: number;
  gap: number;
}

// Note Numbering Type
export type NoteNumberingType = 'Continuous' | 'OnSection' | 'OnPage';

// Note Placement
export type NotePlacement = 'EachColumn' | 'MergedColumn' | 'RightMostColumn' | 'EndOfDocument' | 'EndOfSection';

export type TabType = 'Left' | 'Right' | 'Center' | 'Decimal' | 'left' | 'right' | 'center' | 'decimal';

export type BreakLatinWord = 'KeepWord' | 'Hyphenation' | 'BreakWord' | 'keepWord' | 'normal' | 'hyphenation' | 'breakWord';

// ============================================================
// Section 4: Header Elements (헤더 엘리먼트)
// ============================================================

// 4.3.2 Font Info
export interface FontInfo {
  id: number;
  type?: 'rep' | 'ttf' | 'hft';
  name: string;
  substFont?: {
    type?: 'rep' | 'ttf' | 'hft';
    name: string;
  };
  typeInfo?: {
    familyType?: number;
    serifStyle?: number;
    weight?: number;
    proportion?: number;
    contrast?: number;
    strokeVariation?: number;
    armStyle?: number;
    letterform?: number;
    midline?: number;
    xHeight?: number;
  };
}

// 4.3.3 Border/Fill Info
export interface BorderStyle {
  type?: LineType1;
  width?: LineWidth | string | number;
  color?: string;
}

export interface WindowBrush {
  faceColor?: string;
  hatchColor?: string;
  hatchStyle?: HatchStyle;
  alpha?: number;
}

export interface GradationFill {
  type: GradationType;
  angle?: number;
  centerX?: number;
  centerY?: number;
  step?: number;
  colorNum?: number;
  stepCenter?: number;
  alpha?: number;
  colors: string[];
}

export interface ImageBrush {
  mode: InfillMode;
  bright?: number;
  contrast?: number;
  effect?: 'RealPic' | 'GrayScale' | 'BlackWhite';
  binItem?: string;
  alpha?: number;
}

export interface FillBrush {
  windowBrush?: WindowBrush;
  gradation?: GradationFill;
  imageBrush?: ImageBrush;
}

export interface BorderFillStyle {
  id: number;
  threeD?: boolean;
  shadow?: boolean;
  slash?: number;
  backSlash?: number;
  crookedSlash?: number;
  centerLine?: boolean;
  leftBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  rightBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  topBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  bottomBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  diagonal?: BorderStyle;
  diagonalBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  antiDiagonalBorder?: BorderStyle | { style?: string; width?: number | string; color?: string };
  fillBrush?: FillBrush;
  fillColor?: string;
  fillType?: FillType;
  gradation?: {
    type: GradationType;
    angle?: number;
    centerX?: number;
    centerY?: number;
    step?: number;
    colors: string[];
  };
  imageFill?: {
    mode: ImageFillMode;
    alpha?: number;
    binaryItemId?: string;
  };
}

// 4.3.4 Character Shape Info
export interface CharShape {
  id: number;
  tagName?: 'charShape' | 'charPr';  // Original XML tag name (default: charPr)
  height?: number;          // Font size in hwpunit (1000 = 10pt)
  textColor?: string;       // RGB-Color
  shadeColor?: string;      // RGB-Color
  useFontSpace?: boolean;
  useKerning?: boolean;
  symMark?: EmphasisMark;
  borderFillId?: number;
  // Language-specific font references (font IDs)
  fontRefs?: {
    hangul?: number;
    latin?: number;
    hanja?: number;
    japanese?: number;
    other?: number;
    symbol?: number;
    user?: number;
  };
  // Language-specific font names (resolved from fontRefs)
  fontNames?: {
    hangul?: string;
    latin?: string;
    hanja?: string;
    japanese?: string;
    other?: string;
    symbol?: string;
    user?: string;
  };
  // Language-specific ratios (장평, 50%~200%)
  ratio?: {
    hangul?: number;
    latin?: number;
    hanja?: number;
    japanese?: number;
    other?: number;
    symbol?: number;
    user?: number;
  };
  // Language-specific char spacing (자간, -50%~50%)
  charSpacing?: {
    hangul?: number;
    latin?: number;
    hanja?: number;
    japanese?: number;
    other?: number;
    symbol?: number;
    user?: number;
  };
  // Language-specific relative size (상대크기, 10%~250%)
  relSize?: {
    hangul?: number;
    latin?: number;
    hanja?: number;
    japanese?: number;
    other?: number;
    symbol?: number;
    user?: number;
  };
  // Language-specific char offset (글자위치, -100%~100%)
  charOffset?: {
    hangul?: number;
    latin?: number;
    hanja?: number;
    japanese?: number;
    other?: number;
    symbol?: number;
    user?: number;
  };
  italic?: boolean;
  bold?: boolean;
  underline?: boolean | {
    type: UnderlineType;
    shape: LineType2;
    color: string;
  };
  underlineType?: UnderlineType | 'bottom' | 'center' | 'top' | 'none';
  underlineShape?: UnderlineShape;
  underlineColor?: string;
  strikethrough?: boolean;
  strikeout?: boolean | {
    type: StrikeoutType;
    shape: LineType2;
    color: string;
  };
  strikeoutShape?: StrikeoutShape;
  strikeoutColor?: string;
  outline?: LineType3 | {
    type: LineType3;
  };
  shadow?: ShadowType | {
    type: ShadowType;
    color?: string;
    offsetX?: number;
    offsetY?: number;
    alpha?: number;
  };
  shadowX?: number;
  shadowY?: number;
  shadowColor?: string;
  emboss?: boolean;
  engrave?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
  emphasisMark?: EmphasisMark | 'none' | 'dot' | 'circle' | 'ring' | 'caron' | 'underDot' | 'underLine' | 'triangle';
  relativeSize?: number;
  fontName?: string;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
}

// 4.3.5 Tab Info
export interface TabItem {
  pos: number;    // hwpunit
  type: TabType;
  leader: LineType2;
}

export interface TabDef {
  id: number;
  autoTabLeft?: boolean;
  autoTabRight?: boolean;
  items: TabItem[];
}

// 4.3.6 Numbering/Bullet Info
export interface ParaHeadInfo {
  level: number;           // 1~7
  alignment?: AlignmentType2;
  useInstWidth?: boolean;
  autoIndent?: boolean;
  widthAdjust?: number;    // hwpunit
  textOffsetType?: 'percent' | 'hwpunit';
  textOffset?: number;
  numFormat?: NumberType1;
  charShape?: number;
  text?: string;           // 문단 머리 문자열 포맷
}

export interface NumberingDef {
  id: number;
  start?: number;
  paraHeads: ParaHeadInfo[];
}

export interface BulletDef {
  id: number;
  char?: string;
  image?: boolean;
  useImage?: boolean;
  paraHead?: ParaHeadInfo;
}

// 4.3.7 Paragraph Shape Info
export interface ParaShape {
  id: number;
  align?: AlignmentType1;
  verAlign?: 'Baseline' | 'Top' | 'Center' | 'Bottom';
  headingType?: HeadingType;
  heading?: number;
  level?: number;
  tabDef?: number;
  breakLatinWord?: BreakLatinWord;
  breakNonLatinWord?: boolean;
  condense?: number;
  widowOrphan?: boolean;
  keepWithNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  fontLineHeight?: boolean;
  snapToGrid?: boolean;
  lineWrap?: LineWrapType;
  autoSpaceEAsianEng?: boolean;
  autoSpaceEAsianNum?: boolean;
  borderFillId?: number;
  // Paragraph margin (PARAMARGIN)
  margin?: {
    indent?: number;
    left?: number;
    right?: number;
    prev?: number;
    next?: number;
    lineSpacingType?: 'Percent' | 'Fixed' | 'BetweenLines' | 'AtLeast';
    lineSpacing?: number;
  };
  // Paragraph border (PARABORDER)
  border?: {
    borderFill?: number;
    offsetLeft?: number;
    offsetRight?: number;
    offsetTop?: number;
    offsetBottom?: number;
    connect?: boolean;
    ignoreMargin?: boolean;
  };
  lineSpacing?: number;
  lineSpacingType?: string;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  firstLineIndent?: number;
  tabDefId?: number;
  suppressLineNumbers?: boolean;
  headingLevel?: number;
  widowControl?: boolean;
}

export interface StyleDef {
  id: number;
  type?: 'Para' | 'Char' | 'para' | 'char';
  name?: string;
  engName?: string;
  paraShape?: number;
  charShape?: number;
  nextStyle?: number;
  langId?: number;
  lockForm?: boolean;
  paraPrIdRef?: number;
  charPrIdRef?: number;
  nextStyleIdRef?: number;
}

// 4.3.9 Memo Info
export interface MemoShape {
  id: number;
  width?: number;
  lineType?: LineType1;
  lineColor?: string;
  fillColor?: string;
  activeColor?: string;
  memoType?: string;
}

// ============================================================
// Section 5: Body Elements (본문 엘리먼트)
// ============================================================

// 5.1 Character Elements
export interface CharacterStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | { type: UnderlineType; shape: LineType2; color: string };
  underlineType?: UnderlineType;
  underlineShape?: LineType2 | UnderlineShape;
  underlineColor?: string;
  strikethrough?: boolean;
  strikeoutShape?: LineType2 | StrikeoutShape;
  strikeoutColor?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  backgroundColor?: string;
  superscript?: boolean;
  subscript?: boolean;
  charSpacing?: number | FontRef;
  relativeSize?: number | FontRef;
  charOffset?: number | FontRef;
  emphasisMark?: EmphasisMark;
  useFontSpace?: boolean;
  useKerning?: boolean;
  outline?: LineType3 | { type: LineType3 };
  shadow?: ShadowType | { type: ShadowType; color?: string; offsetX?: number; offsetY?: number; alpha?: number };
  shadowX?: number;
  shadowY?: number;
  shadowColor?: string;
  emboss?: boolean;
  engrave?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
}

export interface TabInfo {
  width: number;
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'solid' | 'dash' | 'dashDot' | 'dashDotDot';
}

export interface HyperlinkField {
  fieldType: 'Hyperlink' | 'hyperlink';
  url: string;
  name?: string;
  command?: string;
}

export interface BookmarkField {
  fieldType: 'Bookmark' | 'bookmark';
  bookmarkName: string;
}

export interface MemoField {
  fieldType: 'Memo' | 'memo';
  memoContent?: string;
  author?: string;
  date?: string;
}

export interface FormulaField {
  fieldType: 'Formula' | 'formula';
  formulaScript?: string;
}

export interface FieldControl {
  fieldType: FieldType;
  name?: string;
  instId?: string;
  editable?: boolean;
  dirty?: boolean;
  property?: string;
  command?: string;
  text?: string;
}

export interface TextRun {
  text: string;
  charPrIDRef?: number;  // Reference to character property in header.xml
  charStyle?: CharacterStyle;
  tab?: TabInfo;
  hyperlink?: HyperlinkField;
  field?: FieldControl;
  markPen?: {
    color: string;
  };
  hasMemo?: boolean;
  memoId?: string;
  footnoteRef?: number;
  endnoteRef?: number;
}

// 5.2 Section Definition Elements
export interface PageDef {
  landscape?: 0 | 1;
  width?: number;      // hwpunit (default 59528 = A4)
  height?: number;     // hwpunit (default 84188 = A4)
  gutterType?: GutterType;
  margin?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
    gutter?: number;
  };
}

export interface StartNumber {
  pageStartsOn?: PageStartsOn;
  page?: number;
  figure?: number;
  table?: number;
  equation?: number;
}

export interface HideOptions {
  header?: boolean;
  footer?: boolean;
  masterPage?: boolean;
  border?: boolean;
  fill?: boolean;
  pageNumPos?: boolean;
  emptyLine?: boolean;
}

export interface AutoNumFormat {
  type?: NumberType2;
  userChar?: string;
  prefixChar?: string;
  suffixChar?: string;
  superscript?: boolean;
}

export interface NoteLine {
  length?: string;
  type?: LineType1;
  width?: LineWidth | string;
  color?: string;
}

export interface NoteSpacing {
  aboveLine?: number;
  belowLine?: number;
  betweenNotes?: number;
}

export interface NoteNumbering {
  type?: NoteNumberingType;
  newNumber?: number;
}

export interface NotePlacementInfo {
  place?: NotePlacement;
  beneathText?: boolean;
}

export interface FootnoteShape {
  autoNumFormat?: AutoNumFormat;
  noteLine?: NoteLine;
  noteSpacing?: NoteSpacing;
  noteNumbering?: NoteNumbering;
  notePlacement?: NotePlacementInfo;
}

export interface PageBorderFill {
  type?: PageStartsOn;
  borderFill?: number;
  textBorder?: boolean;
  headerInside?: boolean;
  footerInside?: boolean;
  fillArea?: 'Paper' | 'Page' | 'Border';
  offset?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  };
}

export interface MasterPage {
  type?: PageStartsOn;
  textWidth?: number;
  textHeight?: number;
  hasTextRef?: boolean;
  hasNumRef?: boolean;
  paragraphs?: HwpxParagraph[];
  // Extended master page properties
  isExtended?: boolean;
  pageNumber?: number;
  pageDuplicate?: boolean;
  pageFront?: boolean;
}

export interface SectionDef {
  textDirection?: 0 | 1;
  spaceColumns?: number;
  tabStop?: number;
  outlineShape?: number;
  lineGrid?: number;
  charGrid?: number;
  firstBorder?: boolean;
  firstFill?: boolean;
  extMasterpageCount?: number;
  memoShapeId?: number;
  pageDef?: PageDef;
  startNumber?: StartNumber;
  hide?: HideOptions;
  footnoteShape?: FootnoteShape;
  endnoteShape?: FootnoteShape;
  pageBorderFill?: PageBorderFill[];
  masterPage?: MasterPage[];
}

// 5.3 Column Definition
export interface ColumnLine {
  type?: LineType1;
  width?: LineWidth | string;
  color?: string;
}

export interface Column {
  width?: number;
  gap?: number;
}

export interface ColumnDef {
  type?: ColumnType;
  count?: number;
  layout?: ColumnLayout;
  sameSize?: boolean;
  sameGap?: number;
  columnLine?: ColumnLine;
  columns?: Column[];
}

// ============================================================
// 5.4 Table (표)
// ============================================================

export interface ShapeSize {
  width: number;
  height: number;
  widthRelTo?: SizeRelTo;
  heightRelTo?: SizeRelTo;
  protect?: boolean;
}

export interface ShapePosition {
  treatAsChar?: boolean;
  affectLSpacing?: boolean;
  vertRelTo?: VertRelTo;
  vertAlign?: VertAlign;
  horzRelTo?: HorzRelTo;
  horzAlign?: HorzAlign;
  vertOffset?: number;
  horzOffset?: number;
  flowWithText?: boolean;
  allowOverlap?: boolean;
  holdAnchorAndSO?: boolean;
}

export interface ObjectMargin {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

export interface Caption {
  side?: 'Left' | 'Right' | 'Top' | 'Bottom';
  fullSize?: boolean;
  width?: number;
  gap?: number;
  lastWidth?: number;
  paragraphs?: HwpxParagraph[];
}

export interface ShapeObject {
  instId?: string;
  zOrder?: number;
  numberingType?: NumberingType;
  textWrap?: TextWrapType;
  textFlow?: TextFlowType;
  lock?: boolean;
  size?: ShapeSize;
  position?: ShapePosition;
  outMargin?: ObjectMargin;
  caption?: Caption;
  shapeComment?: string;
}

export interface CellZone {
  startRowAddr: number;
  startColAddr: number;
  endRowAddr: number;
  endColAddr: number;
  borderFill?: number;
}

export interface CellMargin {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

export type CellElement =
  | { type: 'paragraph'; data: HwpxParagraph }
  | { type: 'table'; data: HwpxTable };

export interface TableCell {
  name?: string;
  colAddr?: number;
  rowAddr?: number;
  colSpan?: number;
  rowSpan?: number;
  width?: number;
  height?: number;
  header?: boolean;
  hasMargin?: boolean; // If false, use table's inMargin instead of cell's margin
  protect?: boolean;
  editable?: boolean;
  dirty?: boolean;
  borderFillId?: number;
  // Content
  paragraphs: HwpxParagraph[];
  nestedTables?: HwpxTable[];
  elements?: CellElement[];
  // Parsed from borderFill
  backgroundColor?: string;
  backgroundGradation?: {
    type: GradationType;
    angle?: number;
    colors: string[];
  };
  borderTop?: BorderStyle;
  borderBottom?: BorderStyle;
  borderLeft?: BorderStyle;
  borderRight?: BorderStyle;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  textDirection?: 'horizontal' | 'vertical';
  lineWrap?: LineWrapType;
}

export interface TableRow {
  cells: TableCell[];
  height?: number;
}

export interface HwpxTable {
  id: string;
  pageBreak?: PageBreakType;
  repeatHeader?: boolean;
  rowCount?: number;
  colCount?: number;
  rowCnt?: number;
  colCnt?: number;
  cellSpacing?: number;
  borderFillId?: number;
  shapeObject?: ShapeObject;
  inMargin?: ObjectMargin;
  cellZoneList?: CellZone[];
  rows: TableRow[];
  width?: number;
  height?: number;
  columnWidths?: number[];
  borderCollapse?: boolean;
  zOrder?: number;
  numberingType?: NumberingType;
  textWrap?: TextWrapType;
  textFlow?: TextFlowType;
  position?: ShapePosition;
  outMargin?: ObjectMargin;
  lock?: boolean;
  linesegs?: LineSeg[];
}

// ============================================================
// 5.5 Picture (그림)
// ============================================================

export interface ShapeComponent {
  hRef?: string;
  xPos?: number;
  yPos?: number;
  groupLevel?: number;
  oriWidth?: number;
  oriHeight?: number;
  curWidth?: number;
  curHeight?: number;
  horzFlip?: boolean;
  vertFlip?: boolean;
  instId?: string;
  rotationInfo?: {
    angle: number;
    centerX?: number;
    centerY?: number;
  };
  renderingInfo?: {
    transMatrix?: number[];
    scaMatrix?: number[];
    rotMatrix?: number[];
  };
}

export interface LineShape {
  color?: string;
  width?: number;
  style?: LineType1;
  endCap?: 'Round' | 'Flat';
  headStyle?: ArrowType;
  tailStyle?: ArrowType;
  headSize?: ArrowSize;
  tailSize?: ArrowSize;
  outlineStyle?: 'Normal' | 'Outer' | 'Inner';
  alpha?: number;
}

export interface ImageRect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  x3?: number;
  y3?: number;
}

export interface ImageClip {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface ShadowEffect {
  style?: string;
  alpha?: number;
  radius?: number;
  direction?: number;
  distance?: number;
  alignStyle?: string;
  skewX?: number;
  skewY?: number;
  scaleX?: number;
  scaleY?: number;
  rotationStyle?: string;
  color?: string;
}

export interface GlowEffect {
  alpha?: number;
  radius?: number;
  color?: string;
}

export interface SoftEdgeEffect {
  radius?: number;
}

export interface ReflectionEffect {
  alignStyle?: string;
  radius?: number;
  direction?: number;
  distance?: number;
  skewX?: number;
  skewY?: number;
  scaleX?: number;
  scaleY?: number;
  rotationStyle?: string;
  startAlpha?: number;
  startPos?: number;
  endAlpha?: number;
  endPos?: number;
  fadeDirection?: number;
}

export interface ImageEffects {
  shadow?: ShadowEffect;
  glow?: GlowEffect;
  softEdge?: SoftEdgeEffect;
  reflection?: ReflectionEffect;
}

export interface HwpxImage {
  id: string;
  binaryId: string;
  width: number;
  height: number;
  orgWidth?: number;
  orgHeight?: number;
  reverse?: boolean;
  shapeObject?: ShapeObject;
  shapeComponent?: ShapeComponent;
  lineShape?: LineShape;
  imageRect?: ImageRect;
  imageClip?: ImageClip;
  effects?: ImageEffects;
  inMargin?: ObjectMargin;
  // Image data
  data?: string;
  mimeType?: string;
  alt?: string;
  zOrder?: number;
  numberingType?: NumberingType;
  textWrap?: TextWrapType;
  textFlow?: TextFlowType;
  position?: ShapePosition;
  outMargin?: ObjectMargin;
  flip?: { horizontal?: boolean; vertical?: boolean };
  rotation?: { angle?: number; centerX?: number; centerY?: number };
  brightness?: number;
  contrast?: number;
  alpha?: number;
  effect?: string;
  shapeComment?: string;
}

// ============================================================
// 5.6 Drawing Objects (그리기 개체)
// ============================================================

export interface DrawingObject {
  shapeComponent?: ShapeComponent;
  lineShape?: LineShape;
  fillBrush?: FillBrush;
  drawText?: {
    lastWidth?: number;
    name?: string;
    editable?: boolean;
    textMargin?: ObjectMargin;
    paragraphs?: HwpxParagraph[];
  };
  shadow?: {
    type?: ShadowType;
    color?: string;
    offsetX?: number;
    offsetY?: number;
    alpha?: number;
  };
}

export interface HwpxLine {
  id: string;
  startX?: number;
  startY?: number;
  endX?: number;
  endY?: number;
  isReverseHV?: boolean;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  strokeColor?: string;
  strokeWidth?: number;
  strokeStyle?: 'solid' | 'dashed' | 'dotted';
}

export interface HwpxRect {
  id: string;
  ratio?: number;
  x0?: number;
  y0?: number;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  x3?: number;
  y3?: number;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
  cornerRadius?: number;
}

export interface HwpxEllipse {
  id: string;
  intervalDirty?: boolean;
  hasArcProperty?: boolean;
  arcType?: ArcType;
  centerX?: number;
  centerY?: number;
  axis1X?: number;
  axis1Y?: number;
  axis2X?: number;
  axis2Y?: number;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
  cx?: number;
  cy?: number;
  rx?: number;
  ry?: number;
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
}

export interface HwpxArc {
  id: string;
  type?: ArcType;
  centerX: number;
  centerY: number;
  axis1X?: number;
  axis1Y?: number;
  axis2X?: number;
  axis2Y?: number;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
}

export interface HwpxPolygon {
  id: string;
  points: Array<{ x: number; y: number }>;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
}

export interface CurveSegment {
  type: 'Line' | 'Curve';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface HwpxCurve {
  id: string;
  segments: CurveSegment[];
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
}

export interface HwpxConnectLine {
  id: string;
  type?: string;
  startX?: number;
  startY?: number;
  endX?: number;
  endY?: number;
  startSubjectID?: string;
  startSubjectIndex?: number;
  endSubjectID?: string;
  endSubjectIndex?: number;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
}

export interface HwpxTextBox {
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
  paragraphs: HwpxParagraph[];
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
}

export interface HwpxHorizontalRule {
  id: string;
  width: number | 'full';
  height: number;
  color?: string;
  style?: 'solid' | 'dashed' | 'dotted' | 'double';
  align?: 'left' | 'center' | 'right';
}

// ============================================================
// 5.8 Form Objects (양식 객체)
// ============================================================

export interface FormCharShape {
  charShape?: number;
  followContext?: boolean;
  autoSize?: boolean;
  wordWrap?: boolean;
}

export interface ButtonSet {
  caption?: string;
  value?: string;
  radioGroupName?: string;
  triState?: boolean;
  backStyle?: string;
}

export interface FormObject {
  name?: string;
  foreColor?: string;
  backColor?: string;
  groupName?: string;
  tabStop?: boolean;
  tabOrder?: number;
  enabled?: boolean;
  borderType?: number;
  drawFrame?: boolean;
  printable?: boolean;
  formCharShape?: FormCharShape;
  buttonSet?: ButtonSet;
}

export interface HwpxButton {
  id: string;
  shapeObject?: ShapeObject;
  formObject?: FormObject;
}

export interface HwpxRadioButton extends HwpxButton {}
export interface HwpxCheckButton extends HwpxButton {}

export interface HwpxComboBox {
  id: string;
  listBoxRows?: number;
  listBoxWidth?: number;
  text?: string;
  editEnable?: boolean;
  shapeObject?: ShapeObject;
  formObject?: FormObject;
}

export interface HwpxEdit {
  id: string;
  multiLine?: boolean;
  passwordChar?: string;
  maxLength?: number;
  scrollBars?: boolean;
  tabKeyBehavior?: string;
  number?: boolean;
  readOnly?: boolean;
  alignText?: string;
  text?: string;
  shapeObject?: ShapeObject;
  formObject?: FormObject;
}

export interface HwpxListBox {
  id: string;
  text?: string;
  itemHeight?: number;
  topIndex?: number;
  shapeObject?: ShapeObject;
  formObject?: FormObject;
}

export interface HwpxScrollBar {
  id: string;
  delay?: number;
  largeChange?: number;
  smallChange?: number;
  min?: number;
  max?: number;
  page?: number;
  value?: number;
  type?: string;
  shapeObject?: ShapeObject;
  formObject?: FormObject;
}

// ============================================================
// 5.7 Unknown Object
// ============================================================

export interface HwpxUnknownObject {
  id: string;
  ctrlId?: string;
  x0?: number;
  y0?: number;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  x3?: number;
  y3?: number;
  shapeObject?: ShapeObject;
  drawingObject?: DrawingObject;
}

// ============================================================
// Parameter Set (Section 5.2)
// ============================================================

export type ParameterItemType = 'Bstr' | 'Integer' | 'Set' | 'Array' | 'BinData';

export interface ParameterItem {
  itemId: string;
  type: ParameterItemType;
  value?: string | number;
  parameterSet?: ParameterSet;
  parameterArray?: ParameterArray;
}

export interface ParameterSet {
  setId?: string;
  count?: number;
  items: ParameterItem[];
}

export interface ParameterArray {
  count?: number;
  items: ParameterItem[];
}

// ============================================================
// 5.9-5.12 Other Objects
// ============================================================

export interface HwpxContainer {
  id: string;
  shapeObject?: ShapeObject;
  shapeComponent?: ShapeComponent;
  children: Array<HwpxLine | HwpxRect | HwpxEllipse | HwpxArc | HwpxPolygon | HwpxCurve | HwpxImage | HwpxContainer>;
}

export interface HwpxOle {
  id: string;
  objectType?: OleObjectType;
  extentX?: number;
  extentY?: number;
  binItem?: string;
  drawAspect?: DrawAspect;
  hasMoniker?: boolean;
  eqBaseLine?: number;
  shapeObject?: ShapeObject;
  shapeComponent?: ShapeComponent;
  lineShape?: LineShape;
}

export interface HwpxEquation {
  id: string;
  lineMode?: boolean;
  baseUnit?: number;
  textColor?: string;
  baseLine?: number;
  version?: string;
  script?: string;
  shapeObject?: ShapeObject;
}

export interface HwpxTextArt {
  id: string;
  text?: string;
  x0?: number;
  y0?: number;
  x1?: number;
  y1?: number;
  x2?: number;
  y2?: number;
  x3?: number;
  y3?: number;
  shape?: {
    fontName?: string;
    fontStyle?: string;
    fontType?: 'ttf' | 'htf';
    textShape?: number;
    lineSpacing?: number;
    charSpacing?: number;
    align?: AlignmentType1;
    shadow?: {
      type?: ShadowType;
      color?: string;
      offsetX?: number;
      offsetY?: number;
      alpha?: number;
    };
  };
  outlineData?: Array<{ x: number; y: number }>;
}

// ============================================================
// 5.13-5.25 Field and Control Elements
// ============================================================

export interface HwpxHyperlink {
  url: string;
  text: string;
}

export interface Bookmark {
  name: string;
}

export interface AutoNum {
  number?: number;
  numberType?: 'Page' | 'Footnote' | 'Endnote' | 'Figure' | 'Table' | 'Equation' | 'TotalPage';
  format?: AutoNumFormat;
}

export interface NewNum extends AutoNum {}

export interface PageNumCtrl {
  pageStartsOn?: PageStartsOn;
}

export interface PageHiding {
  hideHeader?: boolean;
  hideFooter?: boolean;
  hideMasterPage?: boolean;
  hideBorder?: boolean;
  hideFill?: boolean;
  hidePageNum?: boolean;
}

export interface PageNum {
  pos?: 'None' | 'TopLeft' | 'TopCenter' | 'TopRight' | 'BottomLeft' | 'BottomCenter' | 'BottomRight' | 'OutsideTop' | 'OutsideBottom' | 'InsideTop' | 'InsideBottom';
  formatType?: NumberType1;
  sideChar?: string;
}

export interface IndexMark {
  keyFirst?: string;
  keySecond?: string;
}

export interface Compose {
  circleType?: number;
  charSize?: number;
  composeType?: number;
  charShapeIds?: number[];
}

export interface Dutmal {
  posType?: 'Top' | 'Bottom';
  sizeRatio?: number;
  option?: number;
  styleNo?: number;
  align?: AlignmentType1;
  mainText?: string;
  subText?: string;
}

export interface HiddenComment {
  paragraphs: HwpxParagraph[];
}

// ============================================================
// Paragraph Style
// ============================================================

export interface ParagraphStyle {
  align?: 'left' | 'center' | 'right' | 'justify' | 'distribute';
  lineSpacing?: number;
  lineSpacingType?: 'percent' | 'fixed' | 'betweenLines' | 'atLeast';
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  firstLineIndent?: number;
  tabDefId?: number;
  condense?: number;
  breakLatinWord?: BreakLatinWord;
  breakNonLatinWord?: boolean;
  snapToGrid?: boolean;
  suppressLineNumbers?: boolean;
  headingType?: HeadingType;
  headingLevel?: number;
  borderFillId?: number;
  autoSpaceEAsianEng?: boolean;
  autoSpaceEAsianNum?: boolean;
  keepWithNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
}

// ============================================================
// Core Document Elements
// ============================================================

export interface LineSeg {
  vertpos: number;   // Vertical position from page/section top (pt)
  vertsize: number;  // Line height (pt)
  textheight: number; // Text height (pt)
  baseline: number;  // Baseline position (pt)
  spacing: number;   // Line spacing (pt)
}

export interface HwpxParagraph {
  id: string;
  paraPrId?: number;
  style?: number;
  instId?: string;
  pageBreak?: boolean;
  columnBreak?: boolean;
  runs: TextRun[];
  paraStyle?: ParagraphStyle;
  listType?: 'none' | 'bullet' | 'number';
  listLevel?: number;
  linesegs?: LineSeg[];  // Pre-calculated layout info from HWPX
}

export interface HeaderFooter {
  applyPageType?: PageStartsOn;
  seriesNum?: number;
  paragraphs: HwpxParagraph[];
}

export interface Footnote {
  id: string;
  number?: number;
  type?: 'footnote' | 'endnote';
  paragraphs: HwpxParagraph[];
}

export interface Endnote {
  id: string;
  number?: number;
  paragraphs: HwpxParagraph[];
}

export type SectionElement =
  | { type: 'paragraph'; data: HwpxParagraph }
  | { type: 'table'; data: HwpxTable }
  | { type: 'image'; data: HwpxImage }
  | { type: 'line'; data: HwpxLine }
  | { type: 'rect'; data: HwpxRect }
  | { type: 'ellipse'; data: HwpxEllipse }
  | { type: 'arc'; data: HwpxArc }
  | { type: 'polygon'; data: HwpxPolygon }
  | { type: 'curve'; data: HwpxCurve }
  | { type: 'connectline'; data: HwpxConnectLine }
  | { type: 'textbox'; data: HwpxTextBox }
  | { type: 'equation'; data: HwpxEquation }
  | { type: 'ole'; data: HwpxOle }
  | { type: 'container'; data: HwpxContainer }
  | { type: 'textart'; data: HwpxTextArt }
  | { type: 'unknownobject'; data: HwpxUnknownObject }
  | { type: 'hr'; data: HwpxHorizontalRule }
  | { type: 'button'; data: HwpxButton }
  | { type: 'radiobutton'; data: HwpxRadioButton }
  | { type: 'checkbutton'; data: HwpxCheckButton }
  | { type: 'combobox'; data: HwpxComboBox }
  | { type: 'edit'; data: HwpxEdit }
  | { type: 'listbox'; data: HwpxListBox }
  | { type: 'scrollbar'; data: HwpxScrollBar };

export interface PageSettings {
  width: number;
  height: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  headerMargin?: number;
  footerMargin?: number;
  gutterMargin?: number;
  orientation?: 'portrait' | 'landscape';
  gutterType?: GutterType;
}

export interface SectionProperties {
  textDirection?: 'horizontal' | 'vertical';
  spaceColumns?: number;
  tabStop?: number;
  outlineShapeIdRef?: number;
  memoShapeIdRef?: number;
  masterPageCnt?: number;
  masterPage?: MasterPage[];
  grid?: {
    lineGrid?: number;
    charGrid?: number;
    wonggojiFormat?: number;
  };
  startNum?: {
    pageStartsOn?: PageStartsOn;
    page?: number;
    pic?: number;
    tbl?: number;
    equation?: number;
  };
  visibility?: {
    hideFirstHeader?: boolean;
    hideFirstFooter?: boolean;
    hideFirstMasterPage?: boolean;
    border?: 'showAll' | 'hideAll' | 'showFirstPageOnly' | 'showAllButFirstPage';
    fill?: 'showAll' | 'hideAll' | 'showFirstPageOnly' | 'showAllButFirstPage';
    hideFirstPageNum?: boolean;
    hideFirstEmptyLine?: boolean;
    showLineNumber?: boolean;
  };
  pageBorderFill?: Array<{
    type?: PageStartsOn;
    borderFillIdRef?: number;
    textBorder?: 'paper' | 'page' | 'content';
    headerInside?: boolean;
    footerInside?: boolean;
    fillArea?: 'paper' | 'page' | 'content';
    offset?: ObjectMargin;
  }>;
}

export interface FootnoteEndnoteProperties {
  autoNumFormat?: {
    type?: NumberType2;
    userChar?: string;
    prefixChar?: string;
    suffixChar?: string;
    superscript?: boolean;
  };
  noteLine?: {
    length?: number;
    type?: LineType1;
    width?: LineWidth | string;
    color?: string;
  };
  noteSpacing?: {
    betweenNotes?: number;
    belowLine?: number;
    aboveLine?: number;
  };
  numbering?: {
    type?: NoteNumberingType;
    newNum?: number;
  };
  placement?: {
    place?: NotePlacement;
    beneathText?: boolean;
  };
}

export interface Memo {
  id: string;
  author: string;
  date: string;
  content: string[];
  linkedText?: string; // Text that the memo is attached to
}

export interface HwpxSection {
  id?: string;
  elements: SectionElement[];
  pageSettings?: PageSettings;
  sectionProperties?: SectionProperties;
  sectionDef?: SectionDef;
  columnDef?: ColumnDef;
  footnoteProperties?: FootnoteEndnoteProperties;
  endnoteProperties?: FootnoteEndnoteProperties;
  header?: HeaderFooter;
  footer?: HeaderFooter;
  memos?: Memo[];
}

// ============================================================
// Document Metadata
// ============================================================

export interface HwpxMetadata {
  title?: string;
  subject?: string;
  creator?: string;
  createdDate?: string;
  modifiedDate?: string;
  description?: string;
  keywords?: string[];
  comments?: string;
  forbiddenStrings?: string[];
}

export interface DocSetting {
  beginNumber?: {
    page?: number;
    footnote?: number;
    endnote?: number;
    picture?: number;
    table?: number;
    equation?: number;
    totalPage?: number;
  };
  caretPos?: {
    list?: string;
    para?: string;
    pos?: string;
  };
}

// ============================================================
// Styles Collection
// ============================================================

export interface HwpxStyles {
  charShapes: Map<number, CharShape>;
  paraShapes: Map<number, ParaShape>;
  fonts: Map<number, string>;
  fontsByLang: Map<string, string>;  // key: "lang_id" (e.g., "hangul_13"), value: font name
  borderFills: Map<number, BorderFillStyle>;
  tabDefs: Map<number, TabDef>;
  numberings: Map<number, NumberingDef>;
  bullets: Map<number, BulletDef>;
  styles: Map<number, StyleDef>;
  memoShapes: Map<number, MemoShape>;
}

// ============================================================
// Binary Data
// ============================================================

export interface BinItem {
  type: 'Link' | 'Embedding' | 'Storage';
  aPath?: string;
  rPath?: string;
  binData?: string;
  format?: 'jpg' | 'bmp' | 'gif' | 'png' | 'ole';
}

export interface BinData {
  id: string;
  size?: number;
  encoding?: 'Base64';
  compress?: boolean;
  data: string;
}

// ============================================================
// Section 3: Root Element (HWPML)
// ============================================================

export type HwpmlStyle = 'embed' | 'export';

export interface HwpmlRoot {
  version?: string;        // HWPML version, default "2.8"
  subVersion?: string;     // default "8.0.0.0"
  style2?: HwpmlStyle;     // embed | export, default "embed"
  head?: HwpmlHead;
  body?: HwpmlBody;
  tail?: HwpmlTail;
}

export interface HwpmlHead {
  secCnt?: number;         // Number of sections
  docSummary?: HwpxMetadata;
  docSetting?: DocSetting;
  mappingTable?: HwpxStyles;
  compatibleDocument?: CompatibleDocument;
}

export interface HwpmlBody {
  sections: HwpxSection[];
}

export interface HwpmlTail {
  binDataStorage?: BinDataStorage;
  scriptCode?: ScriptCode;
  xmlTemplate?: XmlTemplate;
}

// ============================================================
// Section 6: Tail Elements
// ============================================================

export interface BinDataStorage {
  binData: BinData[];
}

export interface ScriptCode {
  type?: 'JScript';
  version?: string;
  header?: string;
  source?: string;
  preScript?: ScriptFunction[];
  postScript?: ScriptFunction[];
}

export interface ScriptFunction {
  name?: string;
  code: string;
}

export interface XmlTemplate {
  schema?: string;
  instance?: string;
}

export interface CompatibleDocument {
  targetProgram?: 'None' | 'Hwp70' | 'Word';
  layoutCompatibility?: LayoutCompatibility;
}

export interface LayoutCompatibility {
  applyFontWeightToBold?: boolean;
  useInnerUnderline?: boolean;
  fixedUnderlineWidth?: boolean;
  doNotApplyStrikeout?: boolean;
  useLowercaseStrikeout?: boolean;
  extendLineheightToOffset?: boolean;
  treatQuotationAsLatin?: boolean;
  doNotAlignWhitespaceOnRight?: boolean;
  doNotAdjustWordInJustify?: boolean;
  baseCharUnitOnEAsian?: boolean;
  baseCharUnitOfIndentOnFirstChar?: boolean;
  adjustLineheightToFont?: boolean;
  adjustBaselineInFixedLinespacing?: boolean;
  excludeOverlappingParaSpacing?: boolean;
  applyNextspacingOfLastPara?: boolean;
  applyAtLeastToPercent100Pct?: boolean;
  doNotApplyAutoSpaceEAsianEng?: boolean;
  doNotApplyAutoSpaceEAsianNum?: boolean;
  adjustParaBorderfillToSpacing?: boolean;
  connectParaBorderfillOfEqualBorder?: boolean;
  adjustParaBorderOffsetWithBorder?: boolean;
  extendLineheightToParaBorderOffset?: boolean;
  applyParaBorderToOutside?: boolean;
  baseLinespacingOnLinegrid?: boolean;
  applyCharSpacingToCharGrid?: boolean;
  doNotApplyGridInHeaderfooter?: boolean;
  extendHeaderfooterToBody?: boolean;
  adjustEndnotePositionToFootnote?: boolean;
  doNotApplyImageEffect?: boolean;
  doNotApplyShapeComment?: boolean;
  doNotAdjustEmptyAnchorLine?: boolean;
  overlapBothAllowOverlap?: boolean;
  doNotApplyVertOffsetOfForward?: boolean;
  extendVertLimitToPageMargins?: boolean;
  doNotHoldAnchorOfTable?: boolean;
  doNotFormattingAtBeneathAnchor?: boolean;
  doNotApplyExtensionCharCompose?: boolean;
}

// ============================================================
// Main Content Interface
// ============================================================

export interface HwpxContent {
  metadata: HwpxMetadata;
  docSetting?: DocSetting;
  sections: HwpxSection[];
  images: Map<string, HwpxImage>;
  binItems: Map<string, BinItem>;
  binData: Map<string, BinData>;
  footnotes: Footnote[];
  endnotes: Endnote[];
  styles?: HwpxStyles;
  // HWPML root attributes
  hwpmlVersion?: string;
  hwpmlSubVersion?: string;
  hwpmlStyle?: HwpmlStyle;
  compatibleDocument?: CompatibleDocument;
}

export type UnderlineShape = 'solid' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot' | 'long' | 'thick' | 'double' | 'wave' | 'doubleWave' | 'thickWave';
export type StrikeoutShape = 'none' | 'continuous' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot' | 'double' | '3D';
export type OutlineType = 'none' | 'solid' | 'dot' | 'dash' | 'dashDot' | 'dashDotDot' | 'thick' | 'double' | 'triple' | 'thin';
export type BreakWordType = 'normal' | 'hyphenation' | 'breakWord' | 'keepWord';
export type FillType = 'none' | 'color' | 'gradation' | 'image';
export type ImageFillMode = 'tile' | 'tileHorz' | 'tileVert' | 'totalFit' | 'fit' | 'center' | 'onceAbsoluteScale';
export type TabLeader = 'none' | 'solid' | 'dash' | 'dot' | 'dashDot' | 'dashDotDot';
export type NumFormat = 'digit' | 'romanCapital' | 'romanSmall' | 'latinCapital' | 'latinSmall' | 'hangulSyllable' | 'hangulJamo' | 'circledDigit' | 'decimalEnclosedInParentheses';

// Numbering Type for objects (Figure, Table, Equation)
export type NumberingType = 'None' | 'Figure' | 'Table' | 'Equation' | 'Picture'
  | 'none' | 'figure' | 'table' | 'equation' | 'picture';

export interface FontRef {
  hangul?: number;
  latin?: number;
  hanja?: number;
  japanese?: number;
  other?: number;
  symbol?: number;
  user?: number;
}
