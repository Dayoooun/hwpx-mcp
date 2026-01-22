/**
 * Tests for Hanging Indent Calculator (내어쓰기 자동 계산)
 *
 * TDD로 구현: RED → GREEN → REFACTOR
 *
 * 목표: 80% 이상의 정확성으로 마커 기반 내어쓰기 값 계산
 *
 * v3: 복합 법률 마커 패턴, 폰트별 보정, 앞 공백 포함 테스트 포괄
 */
import { describe, it, test, expect } from 'vitest';
import { HangingIndentCalculator } from './HangingIndentCalculator';

describe('HangingIndentCalculator (내어쓰기 자동 계산)', () => {
  // ============================================================
  // 기본 마커 너비 계산 테스트
  // 한글 폰트 보정 계수(1.3) 적용됨
  // ============================================================
  describe('calculateMarkerWidth (마커 너비 계산)', () => {
    it('should calculate width for bullet marker "○ "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('○ ', 10); // 10pt 폰트

      // ○ (1em) + 공백 (0.5em) = 1.5em × 10pt × 1.3 = 19.5pt
      expect(width).toBeGreaterThan(15);
      expect(width).toBeLessThan(25);
    });

    it('should calculate width for dash marker "- "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('- ', 10);

      // - (0.5em) + 공백 (0.5em) = 1em × 10pt × 1.3 = 13pt
      expect(width).toBeGreaterThan(10);
      expect(width).toBeLessThan(18);
    });

    it('should calculate width for numbered marker "1. "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('1. ', 10);

      // 1 (0.6em) + . (0.35em) + 공백 (0.5em) = 1.45em × 10pt × 1.3 = 18.85pt
      expect(width).toBeGreaterThan(15);
      expect(width).toBeLessThan(25);
    });

    it('should calculate width for Korean marker "가. "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('가. ', 10);

      // 가 (1em) + . (0.35em) + 공백 (0.5em) = 1.85em × 10pt × 1.3 = 24.05pt
      expect(width).toBeGreaterThan(20);
      expect(width).toBeLessThan(30);
    });

    it('should calculate width for parenthesized marker "(1) "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('(1) ', 10);

      // ( (0.4em) + 1 (0.6em) + ) (0.4em) + 공백 (0.5em) = 1.9em × 10pt × 1.3 = 24.7pt
      expect(width).toBeGreaterThan(20);
      expect(width).toBeLessThan(30);
    });

    it('should calculate width for circled number "① "', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('① ', 10);

      // ① (1em) + 공백 (0.5em) = 1.5em × 10pt × 1.3 = 19.5pt
      expect(width).toBeGreaterThan(15);
      expect(width).toBeLessThan(25);
    });
  });

  // ============================================================
  // 폰트 크기 스케일링 테스트
  // ============================================================
  describe('font size scaling (폰트 크기 스케일링)', () => {
    it('should scale marker width proportionally to font size', () => {
      const calc = new HangingIndentCalculator();

      const width10pt = calc.calculateMarkerWidth('○ ', 10);
      const width20pt = calc.calculateMarkerWidth('○ ', 20);

      // 20pt 폰트는 10pt 폰트의 2배 너비
      expect(width20pt).toBeCloseTo(width10pt * 2, 1);
    });

    it('should handle small font sizes', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('- ', 8);

      // 1em × 8pt × 1.3 = 10.4pt
      expect(width).toBeGreaterThan(8);
      expect(width).toBeLessThan(15);
    });

    it('should handle large font sizes', () => {
      const calc = new HangingIndentCalculator();
      const width = calc.calculateMarkerWidth('1. ', 24);

      // 1.45em × 24pt × 1.3 = 45.24pt
      expect(width).toBeGreaterThan(35);
      expect(width).toBeLessThan(55);
    });
  });

  // ============================================================
  // 마커 감지 테스트
  // ============================================================
  describe('detectMarker (마커 감지)', () => {
    it('should detect bullet marker at start of text', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('○ 첫 번째 항목');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('○ ');
      expect(result?.type).toBe('bullet');
    });

    it('should detect dash marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('- 항목 내용');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('- ');
      expect(result?.type).toBe('bullet');
    });

    it('should detect numbered marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('1. 첫 번째');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('1. ');
      expect(result?.type).toBe('number');
    });

    it('should detect two-digit numbered marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('12. 열두 번째');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('12. ');
      expect(result?.type).toBe('number');
    });

    it('should detect Korean alphabetic marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('가. 첫 번째 항목');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('가. ');
      expect(result?.type).toBe('korean');
    });

    it('should detect parenthesized number marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('(1) 괄호 번호');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('(1) ');
      expect(result?.type).toBe('parenthesized');
    });

    it('should detect circled number marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('① 원문자');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('① ');
      expect(result?.type).toBe('circled');
    });

    it('should return null for text without marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('일반 텍스트');

      expect(result).toBeNull();
    });

    it('should detect Roman numeral marker', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('I. 로마 숫자');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('I. ');
      expect(result?.type).toBe('roman');
    });
  });

  // ============================================================
  // 자동 계산 통합 테스트
  // ============================================================
  describe('calculateHangingIndent (자동 계산)', () => {
    it('should calculate hanging indent from text with marker', () => {
      const calc = new HangingIndentCalculator();
      const indent = calc.calculateHangingIndent('○ 항목 내용', 10);

      // 1.5em × 10pt × 1.3 = 19.5pt
      expect(indent).toBeGreaterThan(15);
      expect(indent).toBeLessThan(25);
    });

    it('should return 0 for text without marker', () => {
      const calc = new HangingIndentCalculator();
      const indent = calc.calculateHangingIndent('일반 텍스트', 10);

      expect(indent).toBe(0);
    });

    it('should use default font size when not provided', () => {
      const calc = new HangingIndentCalculator();
      const indent = calc.calculateHangingIndent('- 항목');

      // 기본 폰트 크기 12pt 기준
      expect(indent).toBeGreaterThan(0);
    });
  });

  // ============================================================
  // HWPUNIT 변환 테스트
  // ============================================================
  describe('toHwpUnit (HWPUNIT 변환)', () => {
    it('should convert points to HWPUNIT', () => {
      const calc = new HangingIndentCalculator();

      // HWPUNIT = points × 100
      expect(calc.toHwpUnit(10)).toBe(1000);
      expect(calc.toHwpUnit(15.5)).toBe(1550);
    });

    it('should calculate hanging indent in HWPUNIT', () => {
      const calc = new HangingIndentCalculator();
      const hwpUnit = calc.calculateHangingIndentInHwpUnit('○ 항목', 10);

      // 약 19.5pt → 1950 HWPUNIT
      expect(hwpUnit).toBeGreaterThan(1500);
      expect(hwpUnit).toBeLessThan(2500);
    });
  });

  // ============================================================
  // 다양한 마커 패턴 테스트
  // ============================================================
  describe('various marker patterns (다양한 마커 패턴)', () => {
    const testCases = [
      { text: '• 불릿', expectedType: 'bullet' },
      { text: '▪ 네모 불릿', expectedType: 'bullet' },
      { text: '◆ 다이아몬드', expectedType: 'bullet' },
      { text: '※ 참고', expectedType: 'bullet' },
      { text: 'A. 알파벳', expectedType: 'alpha' },
      { text: 'a) 소문자', expectedType: 'alpha' },
      { text: '나. 두번째', expectedType: 'korean' },
      { text: '(가) 괄호 한글', expectedType: 'parenthesized_korean' },
      { text: 'II. 로마 2', expectedType: 'roman' },
      { text: '② 원문자 2', expectedType: 'circled' },
    ];

    testCases.forEach(({ text, expectedType }) => {
      it(`should detect "${text.substring(0, 10)}..." as ${expectedType}`, () => {
        const calc = new HangingIndentCalculator();
        const result = calc.detectMarker(text);

        expect(result).not.toBeNull();
        expect(result?.type).toBe(expectedType);
      });
    });
  });

  // ============================================================
  // 엣지 케이스 테스트
  // ============================================================
  describe('edge cases (엣지 케이스)', () => {
    it('should handle empty string', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('');

      expect(result).toBeNull();
    });

    it('should handle marker only (no content)', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('1. ');

      expect(result).not.toBeNull();
      expect(result?.marker).toBe('1. ');
    });

    it('should not detect marker in middle of text', () => {
      const calc = new HangingIndentCalculator();
      const result = calc.detectMarker('앞에 내용 1. 뒤에 내용');

      expect(result).toBeNull();
    });

    it('should detect marker with leading whitespace', () => {
      const calc = new HangingIndentCalculator();
      // v2: 앞 공백도 허용하고 너비에 포함
      const result = calc.detectMarker('  1. 들여쓰기된 항목');

      expect(result).not.toBeNull();
      expect(result?.leadingSpaces).toBe(2);
    });
  });
});

// ============================================================
// 복합 마커 패턴 테스트 (법률/공문서)
// ============================================================
describe('HangingIndentCalculator - 복합 마커 패턴', () => {
  const calculator = new HangingIndentCalculator();

  describe('detectMarker - 법률 마커', () => {
    test('제1조 감지', () => {
      const result = calculator.detectMarker('제1조 총칙');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('제1조 ');
      expect(result?.leadingSpaces).toBe(0);
    });

    test('제1조의2 감지', () => {
      const result = calculator.detectMarker('제1조의2 적용범위');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('제1조의2 ');
      expect(result?.leadingSpaces).toBe(0);
    });

    test('제10항 감지', () => {
      const result = calculator.detectMarker('제10항 벌칙');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('제10항 ');
    });

    test('제100조의99 감지 (3자리 숫자)', () => {
      const result = calculator.detectMarker('제100조의99 복잡한 법률');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('제100조의99 ');
    });

    test('1호 감지', () => {
      const result = calculator.detectMarker('1호 가입자');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('1호 ');
    });

    test('2목 감지', () => {
      const result = calculator.detectMarker('2목 세부사항');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('2목 ');
    });

    test('15호 감지', () => {
      const result = calculator.detectMarker('15호 기타사항');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');
      expect(result?.marker).toBe('15호 ');
    });
  });

  describe('calculateHangingIndent - 복합 마커 계산', () => {
    test('제1조 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('제1조 내용', 12);
      // "제1조 " = 제(1.0) + 1(0.6) + 조(1.0) + 공백(0.5) = 3.1em
      // 3.1 × 12 × 1.3 = 48.36pt
      expect(indent).toBeGreaterThan(40);
      expect(indent).toBeLessThan(60);
    });

    test('제1조의2 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('제1조의2 내용', 12);
      // "제1조의2 " = 제(1.0) + 1(0.6) + 조(1.0) + 의(1.0) + 2(0.6) + 공백(0.5) = 4.7em
      // 4.7 × 12 × 1.3 = 73.32pt
      expect(indent).toBeGreaterThan(65);
      expect(indent).toBeLessThan(85);
    });

    test('제10항 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('제10항 내용', 12);
      // "제10항 " = 제(1.0) + 1(0.6) + 0(0.6) + 항(1.0) + 공백(0.5) = 3.7em
      // 3.7 × 12 × 1.3 = 57.72pt
      expect(indent).toBeGreaterThan(50);
      expect(indent).toBeLessThan(70);
    });

    test('1호 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('1호 내용', 12);
      // "1호 " = 1(0.6) + 호(1.0) + 공백(0.5) = 2.1em
      // 2.1 × 12 × 1.3 = 32.76pt
      expect(indent).toBeGreaterThan(25);
      expect(indent).toBeLessThan(40);
    });

    test('15호 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('15호 내용', 12);
      // "15호 " = 1(0.6) + 5(0.6) + 호(1.0) + 공백(0.5) = 2.7em
      // 2.7 × 12 × 1.3 = 42.12pt
      expect(indent).toBeGreaterThan(35);
      expect(indent).toBeLessThan(50);
    });

    test('제100조의99 내어쓰기 계산 (12pt)', () => {
      const indent = calculator.calculateHangingIndent('제100조의99 내용', 12);
      // "제100조의99 " = 제(1.0) + 1(0.6) + 0(0.6) + 0(0.6) + 조(1.0) + 의(1.0) + 9(0.6) + 9(0.6) + 공백(0.5) = 6.5em
      // 6.5 × 12 × 1.3 = 101.4pt
      expect(indent).toBeGreaterThan(90);
      expect(indent).toBeLessThan(120);
    });
  });

  describe('폰트별 계산', () => {
    test('맑은 고딕 보정 계수 적용', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const malgunIndent = calculator.calculateHangingIndent('제1조 내용', 12, '맑은 고딕');

      // 맑은 고딕(1.35) > default(1.3)
      expect(malgunIndent).toBeGreaterThan(defaultIndent);

      // 대략 3.8% 차이 (1.35/1.3 ≈ 1.038)
      const ratio = malgunIndent / defaultIndent;
      expect(ratio).toBeCloseTo(1.038, 2);
    });

    test('Noto Sans KR 보정 계수 적용', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const notoIndent = calculator.calculateHangingIndent('제1조 내용', 12, 'Noto Sans KR');

      // Noto Sans KR(1.25) < default(1.3)
      expect(notoIndent).toBeLessThan(defaultIndent);

      // 대략 3.8% 차이 (1.25/1.3 ≈ 0.962)
      const ratio = notoIndent / defaultIndent;
      expect(ratio).toBeCloseTo(0.962, 2);
    });

    test('함초롬바탕 보정 계수 적용', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const hamchoIndent = calculator.calculateHangingIndent('제1조 내용', 12, '함초롬바탕');

      // 함초롬바탕(1.25) < default(1.3)
      expect(hamchoIndent).toBeLessThan(defaultIndent);
    });

    test('나눔바른고딕 보정 계수 적용', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const nanumIndent = calculator.calculateHangingIndent('제1조 내용', 12, '나눔바른고딕');

      // 나눔바른고딕(1.28) < default(1.3)
      expect(nanumIndent).toBeLessThan(defaultIndent);

      const ratio = nanumIndent / defaultIndent;
      expect(ratio).toBeCloseTo(0.985, 2);
    });

    test('영문 폰트 (Arial)', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const arialIndent = calculator.calculateHangingIndent('제1조 내용', 12, 'Arial');

      // Arial(1.0) < default(1.3)
      expect(arialIndent).toBeLessThan(defaultIndent);

      const ratio = arialIndent / defaultIndent;
      expect(ratio).toBeCloseTo(0.769, 2); // 1.0/1.3
    });
  });

  describe('앞 공백 포함 테스트', () => {
    test('앞 공백 없는 제1조', () => {
      const result = calculator.detectMarker('제1조 내용');
      expect(result).not.toBeNull();
      expect(result?.leadingSpaces).toBe(0);
      expect(result?.marker).toBe('제1조 ');
    });

    test('앞 공백 2개 포함 제1조', () => {
      const result = calculator.detectMarker('  제1조 내용');
      expect(result).not.toBeNull();
      expect(result?.leadingSpaces).toBe(2);
      expect(result?.marker).toBe('  제1조 '); // 공백 포함됨
    });

    test('앞 공백 4개 포함 1호', () => {
      const result = calculator.detectMarker('    1호 내용');
      expect(result).not.toBeNull();
      expect(result?.leadingSpaces).toBe(4);
      expect(result?.marker).toBe('    1호 ');
    });

    test('앞 공백 포함 내어쓰기 계산 - 제1조', () => {
      const indent = calculator.calculateHangingIndent('  제1조 내용', 12);
      // "  제1조 " = 공백(0.5) + 공백(0.5) + 제(1.0) + 1(0.6) + 조(1.0) + 공백(0.5) = 4.1em
      // 4.1 × 12 × 1.3 = 63.96pt
      expect(indent).toBeGreaterThan(55);
      expect(indent).toBeLessThan(75);
    });

    test('앞 공백 포함 내어쓰기 계산 - 1호', () => {
      const indent = calculator.calculateHangingIndent('    1호 내용', 12);
      // "    1호 " = 공백(0.5)*4 + 1(0.6) + 호(1.0) + 공백(0.5) = 4.1em
      // 4.1 × 12 × 1.3 = 63.96pt
      expect(indent).toBeGreaterThan(55);
      expect(indent).toBeLessThan(75);
    });
  });

  describe('HWPUNIT 변환 테스트', () => {
    test('points를 HWPUNIT으로 변환', () => {
      expect(calculator.toHwpUnit(10)).toBe(1000);
      expect(calculator.toHwpUnit(12.5)).toBe(1250);
      expect(calculator.toHwpUnit(48.36)).toBe(4836);
    });

    test('제1조 HWPUNIT 계산 (12pt)', () => {
      const hwpunit = calculator.calculateHangingIndentInHwpUnit('제1조 내용', 12);
      // 약 48.36pt × 100 = 4836 hwpunit
      expect(hwpunit).toBeGreaterThan(4000);
      expect(hwpunit).toBeLessThan(6000);
    });

    test('제1조의2 HWPUNIT 계산 (12pt)', () => {
      const hwpunit = calculator.calculateHangingIndentInHwpUnit('제1조의2 내용', 12);
      // 약 73.32pt × 100 = 7332 hwpunit
      expect(hwpunit).toBeGreaterThan(6500);
      expect(hwpunit).toBeLessThan(8500);
    });

    test('1호 HWPUNIT 계산 (맑은 고딕, 12pt)', () => {
      const hwpunit = calculator.calculateHangingIndentInHwpUnit('1호 내용', 12, '맑은 고딕');
      // "1호 " = 2.1em × 12 × 1.35 = 34.02pt → 3402 hwpunit
      expect(hwpunit).toBeGreaterThan(3000);
      expect(hwpunit).toBeLessThan(4000);
    });
  });

  describe('기타 마커 타입 테스트', () => {
    test('괄호 한글 - (가)', () => {
      const result = calculator.detectMarker('(가) 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('parenthesized_korean');
      expect(result?.marker).toBe('(가) ');
    });

    test('괄호 숫자 - (1)', () => {
      const result = calculator.detectMarker('(1) 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('parenthesized');
      expect(result?.marker).toBe('(1) ');
    });

    test('원문자 - ①', () => {
      const result = calculator.detectMarker('① 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('circled');
      expect(result?.marker).toBe('① ');
    });

    test('로마 숫자 - IV.', () => {
      const result = calculator.detectMarker('IV. 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('roman');
      expect(result?.marker).toBe('IV. ');
    });

    test('알파벳 대문자 - A.', () => {
      const result = calculator.detectMarker('A. 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('alpha');
      expect(result?.marker).toBe('A. ');
    });

    test('알파벳 소문자 괄호 - a)', () => {
      const result = calculator.detectMarker('a) 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('alpha');
      expect(result?.marker).toBe('a) ');
    });

    test('한글 점 - 가.', () => {
      const result = calculator.detectMarker('가. 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('korean');
      expect(result?.marker).toBe('가. ');
    });

    test('숫자 점 - 1.', () => {
      const result = calculator.detectMarker('1. 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('number');
      expect(result?.marker).toBe('1. ');
    });

    test('불릿 - ●', () => {
      const result = calculator.detectMarker('● 내용');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('bullet');
      expect(result?.marker).toBe('● ');
    });
  });

  describe('엣지 케이스', () => {
    test('마커 없는 텍스트', () => {
      const result = calculator.detectMarker('그냥 일반 텍스트');
      expect(result).toBeNull();
    });

    test('빈 문자열', () => {
      const result = calculator.detectMarker('');
      expect(result).toBeNull();
    });

    test('공백만', () => {
      const result = calculator.detectMarker('    ');
      expect(result).toBeNull();
    });

    test('마커 뒤 공백 없음 (매칭 실패)', () => {
      const result = calculator.detectMarker('제1조내용');
      expect(result).toBeNull();
    });

    test('마커 없는 경우 내어쓰기 0', () => {
      const indent = calculator.calculateHangingIndent('일반 텍스트');
      expect(indent).toBe(0);
    });

    test('폰트 크기 생략 시 기본값 12pt 사용', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 내용');
      const indent2 = calculator.calculateHangingIndent('제1조 내용', 12);
      expect(indent1).toBe(indent2);
    });

    test('폰트 이름 생략 시 기본 보정 계수 사용', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 내용', 12);
      const indent2 = calculator.calculateHangingIndent('제1조 내용', 12, undefined);
      expect(indent1).toBe(indent2);
    });
  });

  describe('실제 법률 문서 예시', () => {
    test('헌법 제1조', () => {
      const result = calculator.detectMarker('제1조 대한민국은 민주공화국이다.');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');

      const indent = calculator.calculateHangingIndent('제1조 대한민국은 민주공화국이다.', 12);
      expect(indent).toBeGreaterThan(40);
    });

    test('민법 제103조', () => {
      const result = calculator.detectMarker('제103조 선량한 풍속 기타 사회질서에 위반한 사항을 내용으로 하는 법률행위는 무효로 한다.');
      expect(result).not.toBeNull();
      expect(result?.type).toBe('article');

      const indent = calculator.calculateHangingIndent('제103조 선량한 풍속 기타 사회질서에 위반한 사항을 내용으로 하는 법률행위는 무효로 한다.', 12);
      expect(indent).toBeGreaterThan(50);
    });

    test('부칙 제1조', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 이 법은 공포한 날부터 시행한다.', 10);
      const indent2 = calculator.calculateHangingIndent('제1조 이 법은 공포한 날부터 시행한다.', 12);

      // 폰트 크기가 크면 내어쓰기도 증가
      expect(indent2).toBeGreaterThan(indent1);

      const ratio = indent2 / indent1;
      expect(ratio).toBeCloseTo(1.2, 1); // 12/10 = 1.2
    });
  });

  describe('폰트 이름 대소문자/공백 무시 매칭', () => {
    test('맑은고딕 (공백 없음)', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 내용', 12, '맑은 고딕');
      const indent2 = calculator.calculateHangingIndent('제1조 내용', 12, '맑은고딕');
      expect(indent1).toBe(indent2);
    });

    test('Malgun Gothic (영문명)', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 내용', 12, '맑은 고딕');
      const indent2 = calculator.calculateHangingIndent('제1조 내용', 12, 'Malgun Gothic');
      expect(indent1).toBe(indent2);
    });

    test('NanumGothic (공백 없음)', () => {
      const indent1 = calculator.calculateHangingIndent('제1조 내용', 12, '나눔고딕');
      const indent2 = calculator.calculateHangingIndent('제1조 내용', 12, 'NanumGothic');
      expect(indent1).toBe(indent2);
    });

    test('알 수 없는 폰트는 기본값 사용', () => {
      const defaultIndent = calculator.calculateHangingIndent('제1조 내용', 12);
      const unknownIndent = calculator.calculateHangingIndent('제1조 내용', 12, 'UnknownFont123');
      expect(defaultIndent).toBe(unknownIndent);
    });
  });

  describe('복잡한 법률 마커 우선순위', () => {
    test('제1조의2가 제1조보다 우선 매칭', () => {
      const result1 = calculator.detectMarker('제1조 내용');
      const result2 = calculator.detectMarker('제1조의2 내용');

      expect(result1?.marker).toBe('제1조 ');
      expect(result2?.marker).toBe('제1조의2 ');

      // 길이가 다름
      expect(result2!.marker.length).toBeGreaterThan(result1!.marker.length);
    });

    test('제1항의3이 제1항보다 우선 매칭', () => {
      const result1 = calculator.detectMarker('제1항 내용');
      const result2 = calculator.detectMarker('제1항의3 내용');

      expect(result1?.marker).toBe('제1항 ');
      expect(result2?.marker).toBe('제1항의3 ');
    });

    test('1호가 1. (숫자 점)보다 우선 매칭', () => {
      const result1 = calculator.detectMarker('1호 내용');
      const result2 = calculator.detectMarker('1. 내용');

      expect(result1?.type).toBe('article');
      expect(result2?.type).toBe('number');
    });
  });
});
