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
  | 'alpha';

/**
 * 문자 너비 테이블 (em 단위, 기준 폰트 기준)
 * 실제 한글 문서에서 자주 사용되는 마커 문자들의 상대적 너비
 *
 * 한글 폰트(함초롬바탕, 맑은 고딕 등)에서 측정된 값 기준
 */
const CHAR_WIDTH_TABLE: Record<string, number> = {
  // 불릿 문자 (전각 = 1em)
  '○': 1.0,
  '●': 1.0,
  '•': 0.6,  // 중간 불릿
  '▪': 0.6,
  '◆': 1.0,
  '◇': 1.0,
  '■': 1.0,
  '□': 1.0,
  '※': 1.0,
  '★': 1.0,
  '☆': 1.0,
  '◎': 1.0,  // 이중 원
  '◉': 1.0,
  '▶': 1.0,  // 화살표
  '▷': 1.0,
  '►': 1.0,
  '▻': 0.6,
  '▸': 0.6,  // 작은 화살표
  '▹': 0.6,
  '➢': 1.0,
  '➣': 1.0,
  '➤': 1.0,
  '✓': 0.7,  // 체크마크
  '✔': 0.7,
  '✗': 0.7,
  '✘': 0.7,
  '✦': 0.7,  // 별
  '✧': 0.7,
  '→': 1.0,  // 화살표
  '⇒': 1.0,
  '▣': 1.0,  // 박스
  '▤': 1.0,
  '▥': 1.0,

  // 대시/하이픈
  '-': 0.5,
  '–': 0.7, // en-dash
  '—': 1.0, // em-dash

  // 숫자 (반각이지만 한글 폰트에서는 조금 넓음)
  '0': 0.6,
  '1': 0.6,
  '2': 0.6,
  '3': 0.6,
  '4': 0.6,
  '5': 0.6,
  '6': 0.6,
  '7': 0.6,
  '8': 0.6,
  '9': 0.6,

  // 구두점 (한글 폰트에서 조금 넓음)
  '.': 0.35,
  ')': 0.4,
  '(': 0.4,

  // 공백 (한글 폰트에서 더 넓음)
  ' ': 0.5,

  // 한글 자모/글자 (전각)
  '가': 1.0,
  '나': 1.0,
  '다': 1.0,
  '라': 1.0,
  '마': 1.0,
  '바': 1.0,
  '사': 1.0,
  '아': 1.0,
  '자': 1.0,
  '차': 1.0,
  '카': 1.0,
  '타': 1.0,
  '파': 1.0,
  '하': 1.0,

  // 원문자 (전각)
  '①': 1.0,
  '②': 1.0,
  '③': 1.0,
  '④': 1.0,
  '⑤': 1.0,
  '⑥': 1.0,
  '⑦': 1.0,
  '⑧': 1.0,
  '⑨': 1.0,
  '⑩': 1.0,
  '⑪': 1.0,
  '⑫': 1.0,
  '⑬': 1.0,
  '⑭': 1.0,
  '⑮': 1.0,
  '⑯': 1.0,
  '⑰': 1.0,
  '⑱': 1.0,
  '⑲': 1.0,
  '⑳': 1.0,

  // 로마 숫자/알파벳 대문자 (반각이지만 한글 폰트에서 조금 넓음)
  'I': 0.4,
  'V': 0.7,
  'X': 0.7,
  'L': 0.6,
  'C': 0.7,
  'D': 0.7,
  'M': 0.9,
  'A': 0.7,
  'B': 0.7,
  'E': 0.6,
  'F': 0.6,
  'G': 0.7,
  'H': 0.7,

  // 알파벳 소문자 (반각)
  'a': 0.55,
  'b': 0.55,
  'c': 0.55,
  'd': 0.55,
  'e': 0.55,
  'f': 0.35,
  'g': 0.55,
  'h': 0.55,

  // 콜론 (마커 뒤에 올 수 있음)
  ':': 0.35,
};

/**
 * 한글 폰트 보정 계수
 * 한글 문서에서 실제 렌더링되는 너비가 em 값보다 넓음
 */
const HANGUL_FONT_FACTOR = 1.3;

/**
 * 마커 패턴 정의 (순서 중요 - 더 구체적인 패턴이 먼저)
 * 앞 공백도 허용하도록 수정
 */
const MARKER_PATTERNS: Array<{
  regex: RegExp;
  type: MarkerType;
}> = [
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

    // 한글 범위 (가-힣) 체크 - 전각으로 처리
    if (/[가-힣]/.test(char)) {
      return 1.0;
    }

    // 전각 문자 범위 체크
    const code = char.charCodeAt(0);
    if (code >= 0xFF00 && code <= 0xFFEF) {
      return 1.0; // 전각 문자
    }

    // 기본값: 반각 문자로 가정
    return 0.55;
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
   * @returns 마커의 너비 (pt)
   */
  calculateMarkerWidth(marker: string, fontSize: number): number {
    const widthInEm = this.calculateMarkerWidthInEm(marker);
    // em을 pt로 변환하고 한글 폰트 보정 계수 적용
    return widthInEm * fontSize * HANGUL_FONT_FACTOR;
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
   * @returns 내어쓰기 값 (pt)
   */
  calculateHangingIndent(text: string, fontSize?: number): number {
    const size = fontSize ?? HangingIndentCalculator.DEFAULT_FONT_SIZE;
    const markerInfo = this.detectMarker(text);

    if (!markerInfo) {
      return 0;
    }

    return this.calculateMarkerWidth(markerInfo.marker, size);
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
   * @returns 내어쓰기 값 (HWPUNIT)
   */
  calculateHangingIndentInHwpUnit(text: string, fontSize?: number): number {
    const points = this.calculateHangingIndent(text, fontSize);
    return this.toHwpUnit(points);
  }
}
