# 표/이미지 이동 기능 설계

## ✅ 구현 완료 (2026-01-20)

- `moveTable()`, `copyTable()` 메서드 구현
- `validateTagBalance()`, `validateXmlEscaping()` 검증 유틸리티 구현
- `applyTableMovesToXml()` XML 처리 메서드 구현
- `regenerateIdsInXml()` ID 재생성 메서드 구현
- `move_table`, `copy_table` MCP 도구 등록
- 14개 테스트 작성 및 통과

## 개요

XML 태그를 직접 복사하여 원하는 위치로 이동하는 기능 구현.
기존 delete + recreate 방식보다 효율적이고 안전함.

## 범위

- **표(Table)**: `<hp:tbl>` 요소 이동
- **이미지(Image)**: `<hp:pic>` 요소 이동

## API 설계

### moveTable

```typescript
moveTable(
  sectionIndex: number,      // 원본 섹션 인덱스
  tableIndex: number,        // 이동할 표 인덱스 (섹션 내)
  targetSectionIndex: number,// 대상 섹션 인덱스
  targetAfterIndex: number   // 대상 요소 인덱스 (이 요소 뒤에 삽입)
): { success: boolean; error?: string }
```

### moveImage

```typescript
moveImage(
  sectionIndex: number,      // 원본 섹션 인덱스
  imageIndex: number,        // 이동할 이미지 인덱스 (섹션 내)
  targetSectionIndex: number,// 대상 섹션 인덱스
  targetAfterIndex: number   // 대상 요소 인덱스 (이 요소 뒤에 삽입)
): { success: boolean; error?: string }
```

### copyTable / copyImage

동일한 시그니처로 복사 기능 제공 (원본 유지)

## 엄격 검증 (Option B)

### 1. ID 중복 검사

이동 전 대상 섹션에 동일한 ID가 있는지 확인:
- 표: `id` 속성
- 셀 내 문단: `<hp:p id="...">`
- 이미지: `id` 속성, `binItem id`

중복 시 새 ID 생성하여 교체.

### 2. 태그 균형 검사

이동 후 XML 태그 균형 확인:
```typescript
validateTagBalance(xml: string): {
  balanced: boolean;
  mismatches: Array<{ tag: string; opens: number; closes: number }>
}
```

검사 대상 태그:
- `hp:tbl`, `hp:tr`, `hp:tc`
- `hp:p`, `hp:run`, `hp:t`
- `hp:pic`, `hp:imgRect`
- `hp:subList`

### 3. 참조 무결성

- `binItem` 참조가 유효한지 확인 (이미지)
- `charPrIDRef`, `paraPrIDRef` 스타일 참조 확인

### 4. 이스케이프 검증

텍스트 내용이 올바르게 이스케이프되었는지 확인:
- `<` → `&lt;`
- `>` → `&gt;`
- `&` → `&amp;`
- `"` → `&quot;`
- `'` → `&apos;`

### 5. XML 파싱 테스트

저장 전 전체 XML을 파싱하여 well-formed 여부 확인:
```typescript
try {
  new DOMParser().parseFromString(xml, 'text/xml');
} catch (e) {
  throw new Error('Invalid XML after move operation');
}
```

## 구현 순서

1. **테스트 작성** (RED)
   - `MoveTableImage.test.ts` 생성
   - 기본 이동 테스트
   - ID 중복 처리 테스트
   - 태그 균형 테스트
   - 참조 무결성 테스트
   - 에러 케이스 테스트

2. **검증 유틸리티 구현** (GREEN)
   - `validateTagBalance()`
   - `generateUniqueId()`
   - `validateXml()`

3. **이동 메서드 구현** (GREEN)
   - `extractElementXml()` - 요소 XML 추출
   - `moveTable()`
   - `moveImage()`
   - `copyTable()`
   - `copyImage()`

4. **MCP 도구 등록**
   - `move_table`, `move_image`
   - `copy_table`, `copy_image`

## 테스트 케이스

### 기본 기능
- [ ] 같은 섹션 내 표 이동
- [ ] 다른 섹션으로 표 이동
- [ ] 같은 섹션 내 이미지 이동
- [ ] 다른 섹션으로 이미지 이동
- [ ] 표 복사 (원본 유지)
- [ ] 이미지 복사 (원본 유지)

### 검증
- [ ] ID 중복 시 새 ID 생성
- [ ] 태그 균형 검사 통과
- [ ] 참조 무결성 유지
- [ ] 이스케이프 처리 정확성
- [ ] XML well-formed 검사

### 에러 케이스
- [ ] 존재하지 않는 섹션 인덱스
- [ ] 존재하지 않는 표/이미지 인덱스
- [ ] 잘못된 대상 위치
- [ ] 자기 자신 위치로 이동 (no-op)
