/**
 * Hanging Indent Calculator (내어쓰기 자동 계산기)
 *
 * 마커 기반 룩업 테이블 + 폰트 크기 스케일링으로
 * 80% 이상의 정확성을 목표로 함
 *
 * 개선 v2:
 * - 앞 공백 포함 계산
 * - 한글 폰트 보정 계수 적용
 * - 공백 너비 조정
 */

export interface MarkerInfo {
  marker: string;
  type: MarkerType;
  leadingSpaces: number; // 앞 공백 개수
}

export type MarkerType =
  | 'bullet'
  | 'number'
  | 'korean'
  | 'parenthesized'
  | 'parenthesized_korean'
  | 'circled'
  | 'roman'
  | 'alpha'
  | 'article';

/**
 * 문자 너비 테이블 (em 단위).
 *
 * 값은 Windows 한/글이 함초롬바탕 10pt 로 그린 PDF 에서 **글자 전진폭**
 * (다음 글자 x − 이 글자 x) 을 잰 것이다. 추정치가 아니다.
 * 측정 원본: src/fixtures/hancom-marker-widths.json
 *
 * 이전 표는 전각 1.0em·공백 0.5em 에 폰트 보정 1.25 를 곱해, 한/글보다
 * 평균 4.14pt(최대 6.30pt) 넓게 계산했다. 한/글은 전각 기호를 0.996em,
 * 공백을 0.498em 로 그리므로 보정 계수는 1.0 이 맞다.
 *
 * 표에 없는 글자는 getCharWidth 의 범주별 기본값을 쓴다.
 */
const CHAR_WIDTH_TABLE: Record<string, number> = {
  // 전각 기호 (실측 0.996em: ○ ● □ ■ ▶ ①)
  '○': 0.996,
  '●': 0.996,
  '◆': 0.996,
  '◇': 0.996,
  '■': 0.996,
  '□': 0.996,
  '★': 0.996,
  '☆': 0.996,
  '◎': 0.996,
  '◉': 0.996,
  '▶': 0.996,
  '▷': 0.996,
  '►': 0.996,
  '➢': 0.996,
  '➣': 0.996,
  '➤': 0.996,
  '→': 0.996,
  '⇒': 0.996,
  '▣': 0.996,
  '▤': 0.996,
  '▥': 0.996,

  // 폭이 좁은 기호 (실측: ※ 0.756, • 0.456)
  '※': 0.756,
  '•': 0.456,
  '▪': 0.456,
  '▻': 0.456,
  '▸': 0.456,
  '▹': 0.456,
  // 체크·별 기호 (실측 ✓ 0.996)
  '✓': 0.996,
  '✔': 0.996,
  '✗': 0.996,
  '✘': 0.996,
  '✦': 0.996,
  '✧': 0.996,

  // 대시 (실측 - 0.834, — 0.852)
  '-': 0.834,
  '–': 0.834,
  '—': 0.852,

  // 숫자 (실측 1 0.576, 0 0.588)
  '0': 0.588,
  '1': 0.576,
  '2': 0.588,
  '3': 0.588,
  '4': 0.588,
  '5': 0.588,
  '6': 0.588,
  '7': 0.588,
  '8': 0.588,
  '9': 0.588,

  // 구두점·공백 (실측 . 0.294, ( 0.498, ) 0.504, 공백 0.498)
  '.': 0.294,
  ')': 0.504,
  '(': 0.498,
  ':': 0.294,
  ' ': 0.498,

  // 알파벳 (실측 A 0.744, B 0.66, I 0.372, a 0.492)
  'I': 0.372,
  'V': 0.744,
  'X': 0.744,
  'L': 0.6,
  'C': 0.744,
  'D': 0.744,
  'M': 0.9,
  'A': 0.744,
  'B': 0.66,
  'E': 0.66,
  'F': 0.6,
  'G': 0.744,
  'H': 0.744,
  'a': 0.492,
  'b': 0.492,
  'c': 0.492,
  'd': 0.492,
  'e': 0.492,
  'f': 0.35,
  'g': 0.492,
  'h': 0.492,
};

/** 한글 음절 전진폭 (실측 가·본 0.972em). */
const HANGUL_SYLLABLE_WIDTH = 0.972;
/** 원문자·전각 기호 기본 전진폭 (실측 ① 0.996em). */
const FULL_WIDTH_SYMBOL = 0.996;

/**
 * 폰트별 보정 계수 — CHAR_WIDTH_TABLE(함초롬바탕 실측) 대비 배율.
 *
 * 함초롬 계열만 한/글에서 실측했다(1.0). 나머지는 측정하지 않았으므로
 * 종전 표가 함초롬(1.25) 대비 두던 **상대 비율**만 옮겼다
 * (예: 맑은 고딕 1.35/1.25 = 1.08). 새 폰트를 쓰는 양식이 생기면
 * scripts 쪽 측정 문서로 다시 재서 이 값을 교체한다.
 */
const FONT_FACTOR_TABLE: Record<string, number> = {
  // 기본값 — 한/글 새 문서 기본 글꼴이 함초롬바탕이다
  'default': 1.0,

  // 한컴 폰트 (실측)
  '함초롬바탕': 1.0,
  '함초롬돋움': 1.0,
  'HCR Batang': 1.0,
  'HCR Dotum': 1.0,
  '한컴바탕': 1.04,
  '한컴돋움': 1.04,

  // 마이크로소프트 폰트 (미측정, 종전 상대비)
  '맑은 고딕': 1.08,
  '맑은고딕': 1.08,
  'Malgun Gothic': 1.08,
  '바탕': 1.04,
  '돋움': 1.04,
  '굴림': 1.04,
  '궁서': 1.08,

  // 나눔 폰트 (미측정, 종전 상대비)
  '나눔고딕': 1.04,
  '나눔명조': 1.04,
  'NanumGothic': 1.04,
  'NanumMyeongjo': 1.04,
  '나눔바른고딕': 1.024,

  // Adobe 폰트 (미측정, 종전 상대비)
  '본고딕': 1.0,
  '본명조': 1.0,
  'Noto Sans KR': 1.0,
  'Noto Serif KR': 1.0,

  // 영문 폰트 (미측정, 종전 상대비)
  'Arial': 0.8,
  'Times New Roman': 0.8,
};

/**
 * 폰트 보정 계수 가져오기
 */
function getFontFactor(fontName?: string | null): number {
  if (!fontName) return FONT_FACTOR_TABLE['default'];

  // 정확한 매칭 시도
  if (FONT_FACTOR_TABLE[fontName]) {
    return FONT_FACTOR_TABLE[fontName];
  }

  // 부분 매칭 시도 (공백/대소문자 무시)
  const normalizedName = fontName.toLowerCase().replace(/\s+/g, '');
  for (const [key, value] of Object.entries(FONT_FACTOR_TABLE)) {
    if (key.toLowerCase().replace(/\s+/g, '') === normalizedName) {
      return value;
    }
  }

  return FONT_FACTOR_TABLE['default'];
}

/**
 * 마커 패턴 정의 (순서 중요 - 더 구체적인 패턴이 먼저)
 * 앞 공백도 허용하도록 수정
 */
const MARKER_PATTERNS: Array<{
  regex: RegExp;
  type: MarkerType;
}> = [
  // 법률/공문서 마커 (더 구체적인 것이 먼저)
  // 제1조의2, 제1항의3 등
  { regex: /^(\s*)(제\d+[조항호목]의\d+)\s/, type: 'article' },

  // 제1조, 제2항, 제3호, 제4목 등
  { regex: /^(\s*)(제\d+[조항호목])\s/, type: 'article' },

  // 1호, 2목 등 (숫자 + 호/목)
  { regex: /^(\s*)(\d+[호목])\s/, type: 'article' },

  // 괄호 한글: (가), (나), ... (앞 공백 허용)
  { regex: /^(\s*)\(([가-힣])\)\s/, type: 'parenthesized_korean' },

  // 괄호 숫자: (1), (2), ... (앞 공백 허용)
  { regex: /^(\s*)\((\d+)\)\s/, type: 'parenthesized' },

  // 원문자: ①, ②, ... (앞 공백 허용)
  { regex: /^(\s*)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])\s/, type: 'circled' },

  // 로마 숫자: I., II., III., IV., ... (앞 공백 허용)
  { regex: /^(\s*)([IVXLCDM]+)\.\s/, type: 'roman' },

  // 알파벳 대문자 + 점: A., B., ... (앞 공백 허용)
  { regex: /^(\s*)([A-Z])\.\s/, type: 'alpha' },

  // 알파벳 소문자 + 괄호: a), b), ... (앞 공백 허용)
  { regex: /^(\s*)([a-z])\)\s/, type: 'alpha' },

  // 한글 + 점: 가., 나., ... (앞 공백 허용)
  { regex: /^(\s*)([가나다라마바사아자차카타파하])\.\s/, type: 'korean' },

  // 숫자 + 점: 1., 2., 10., 99., ... (앞 공백 허용)
  { regex: /^(\s*)(\d+)\.\s/, type: 'number' },

  // 불릿 문자들 (앞 공백 허용)
  // 기본: ○◦●•▪◆◇■□※★☆
  // 화살표: ▶▷►▻▸▹➢➣➤→⇒
  // 체크/별: ✓✔✗✘✦✧
  // 박스: ▣▤▥
  // 이중원: ◎◉
  // 대시: -–—
  { regex: /^(\s*)([○◦●•▪◆◇■□※★☆◎◉▶▷►▻▸▹➢➣➤✓✔✗✘✦✧→⇒▣▤▥\-–—])\s/, type: 'bullet' },
];

export class HangingIndentCalculator {
  // 기본 폰트 크기 (pt) - 한글 문서 기본값은 보통 10pt 또는 12pt
  private static readonly DEFAULT_FONT_SIZE = 12;

  /**
   * 문자의 너비를 em 단위로 반환
   */
  private getCharWidth(char: string): number {
    if (CHAR_WIDTH_TABLE[char] !== undefined) {
      return CHAR_WIDTH_TABLE[char];
    }

    // 한글 음절 (가-힣)
    if (/[가-힣]/.test(char)) {
      return HANGUL_SYLLABLE_WIDTH;
    }

    const code = char.codePointAt(0) ?? 0;
    // 원문자 ①–⑳ 등 Enclosed Alphanumerics, 도형 기호 블록, 전각 형태
    if ((code >= 0x2460 && code <= 0x24FF) ||
        (code >= 0x25A0 && code <= 0x25FF) ||
        (code >= 0xFF00 && code <= 0xFFEF)) {
      return FULL_WIDTH_SYMBOL;
    }

    // 표에 없는 반각 문자 — 실측한 소문자 폭을 쓴다
    return 0.492;
  }

  /**
   * 마커 문자열의 너비를 em 단위로 계산
   */
  private calculateMarkerWidthInEm(marker: string): number {
    let totalWidth = 0;
    for (const char of marker) {
      totalWidth += this.getCharWidth(char);
    }
    return totalWidth;
  }

  /**
   * 마커 너비 계산 (points 단위)
   *
   * @param marker 마커 문자열 (예: "○ ", "1. ")
   * @param fontSize 폰트 크기 (pt)
   * @param fontName 폰트 이름 (선택적, 기본값 사용시 생략)
   * @returns 마커의 너비 (pt)
   */
  calculateMarkerWidth(marker: string, fontSize: number, fontName?: string | null): number {
    const widthInEm = this.calculateMarkerWidthInEm(marker);
    const fontFactor = getFontFactor(fontName);
    return widthInEm * fontSize * fontFactor;
  }

  /**
   * 텍스트에서 마커 감지 (앞 공백 포함)
   *
   * @param text 텍스트
   * @returns 마커 정보 또는 null
   */
  detectMarker(text: string): MarkerInfo | null {
    if (!text || text.length === 0) {
      return null;
    }

    for (const pattern of MARKER_PATTERNS) {
      const match = text.match(pattern.regex);
      if (match) {
        const leadingSpaces = match[1]?.length || 0;
        return {
          marker: match[0],  // 전체 매치 (앞 공백 + 마커 + 뒤 공백)
          type: pattern.type,
          leadingSpaces,
        };
      }
    }

    return null;
  }

  /**
   * 텍스트에서 내어쓰기 값 자동 계산 (points 단위)
   *
   * @param text 텍스트
   * @param fontSize 폰트 크기 (pt, 기본값 12pt)
   * @param fontName 폰트 이름 (선택적, 기본값 사용시 생략)
   * @returns 내어쓰기 값 (pt)
   */
  calculateHangingIndent(text: string, fontSize?: number, fontName?: string | null): number {
    const size = fontSize ?? HangingIndentCalculator.DEFAULT_FONT_SIZE;
    const markerInfo = this.detectMarker(text);

    if (!markerInfo) {
      return 0;
    }

    return this.calculateMarkerWidth(markerInfo.marker, size, fontName);
  }

  /**
   * points를 HWPUNIT으로 변환
   *
   * @param points 포인트 값
   * @returns HWPUNIT 값 (points × 100)
   */
  toHwpUnit(points: number): number {
    return Math.round(points * 100);
  }

  /**
   * 텍스트에서 내어쓰기 값 자동 계산 (HWPUNIT 단위)
   *
   * @param text 텍스트
   * @param fontSize 폰트 크기 (pt, 기본값 12pt)
   * @param fontName 폰트 이름 (선택적, 기본값 사용시 생략)
   * @returns 내어쓰기 값 (HWPUNIT)
   */
  calculateHangingIndentInHwpUnit(text: string, fontSize?: number, fontName?: string | null): number {
    const points = this.calculateHangingIndent(text, fontSize, fontName);
    return this.toHwpUnit(points);
  }
}
