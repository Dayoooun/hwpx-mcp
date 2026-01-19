/**
 * Tests for Hanging Indent Calculator (내어쓰기 자동 계산)
 *
 * TDD로 구현: RED → GREEN → REFACTOR
 *
 * 목표: 80% 이상의 정확성으로 마커 기반 내어쓰기 값 계산
 *
 * v2: 한글 폰트 보정 계수(1.3), 앞 공백 허용 등 개선
 */
import { describe, it, expect } from 'vitest';
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
