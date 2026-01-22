# 🔧 Solved Problems Log

2번 이상 시도해서 해결한 문제들의 기록 (최근 15개)

---


## [2026-01-22 16:13:00] 컴파일된 코드에 `tableId`가 제대로 포함되어 있습니다. 이제 테스트를 실행해서 버그가 수정되었는지 확

### 🔴 문제

```
컴파일된 코드에 `tableId`가 제대로 포함되어 있습니다. 이제 테스트를 실행해서 버그가 수정되었는지 확인합니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

XML 저장은 성공했지만 리로드 후 테이블 파싱에 문제가 있습니다. 리로드된 문서의 상태를 확인합니다.

**해결 단계:**
- XML 저장은 성공했지만 리로드 후 테이블 파싱에 문제가 있습니다. 리로드된 문서의 상태를 확인합니다.

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 16:51:01] 테스트 실행 결과, 두 가지 문제 발견:

### 🔴 문제

```
테스트 실행 결과, 두 가지 문제 발견:
2. **`fontName.toLowerCase is not a function`** - null/undefined 체크 필요
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

환경에 `cygpath` 관련 문제가 있어 파일 시스템 도구들이 작동하지 않습니다. executor 에이전트를 사용해서 파일을 생성하겠습니다.

**해결 단계:**
- 환경에 `cygpath` 관련 문제가 있어 파일 시스템 도구들이 작동하지 않습니다. executor 에이전트를 사용해서 파일을 생성하겠습니다.

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 16:58:25] 환경 문제로 소스 코드를 직접 읽을 수 없지만, 이전 대화와 CLAUDE.md 문서를 기반으로 잠재적 버그 

### 🔴 문제

```
환경 문제로 소스 코드를 직접 읽을 수 없지만, 이전 대화와 CLAUDE.md 문서를 기반으로 잠재적 버그 리포트를 작성하겠습니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

**해결 방안:**
사용자가 직접 터미널에서 다음을 실행해주세요:

**해결 단계:**
- **현재 상황:**
- - `cygpath` 관련 환경 문제로 모든 Bash/파일 시스템 작업이 실패
- - 테스트 파일 생성 및 실행 불가
- **해결 방안:**
- **레드팀 테스트 코드:**

### 📁 관련 파일

- `CLAUDE.md`
- `mcp-server/src/MasterRedTeam.comprehensive.test.ts`

### 📊 메타정보

- **시도 횟수**: 3번

---


## [2026-01-22 16:58:29] 환경 문제로 소스 코드를 직접 읽을 수 없지만, 이전 대화와 CLAUDE.md 문서를 기반으로 잠재적 버그 

### 🔴 문제

```
환경 문제로 소스 코드를 직접 읽을 수 없지만, 이전 대화와 CLAUDE.md 문서를 기반으로 잠재적 버그 리포트를 작성하겠습니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

**해결 방안:**
사용자가 직접 터미널에서 다음을 실행해주세요:

**해결 단계:**
- **현재 상황:**
- - `cygpath` 관련 환경 문제로 모든 Bash/파일 시스템 작업이 실패
- - 테스트 파일 생성 및 실행 불가
- **해결 방안:**
- **레드팀 테스트 코드:**

### 📁 관련 파일

- `mcp-server/src/MasterRedTeam.comprehensive.test.ts`
- `CLAUDE.md`

### 📊 메타정보

- **시도 횟수**: 3번

---


## [2026-01-22 21:35:24] hwpx-mcp의 문단 업데이트가 저장 시 유지되지 않는 버그가 있습니다. XML 직접 수정을 시도하겠습니다

### 🔴 문제

```
hwpx-mcp의 문단 업데이트가 저장 시 유지되지 않는 버그가 있습니다. XML 직접 수정을 시도하겠습니다.
⎿ Error: result (1,171,082 characters) exceeds maximum allowed tokens. Output has been saved to C:\Users\rlae
● XML이 1.1MB로 너무 큽니다. hwpx-mcp 도구에 문단 업데이트 저장 관련 버그가 있는 것으로 확인됩니다.
❌ hwpx-mcp 버그 발견:
오류기가있어
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

대화 기록 참조

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 21:43:22] ● hwpx-mcp 문제점 정리

### 🔴 문제

```
● hwpx-mcp 문제점 정리
🔴 문제 발생 함수
증상: (테스트 안함, 동일 문제 예상)
1. 메모리 vs 파일 동기화 문제: 메모리상 수정은 되지만 XML 파일에 반영되지 않음
3. HWPX 내부 XML 구조 문제: 문단 XML 업데이트가 실제 파일에 기록되지 않음
[ ] 사용자 직접 입력 필요 (문단 업데이트 버그)
```

### 🔍 원인

문제점을 명확하게 정리해주셨네요. `update_paragraph_text`와 `append_text_to_paragraph`가 저장 시 반영되지 않는 버그입니다. 테이블 셀은 정상 작동하니 두 함수의 차이점을 비교해서 원인을 찾겠습니다.

### ✅ 해결방안

문제점을 명확하게 정리해주셨네요. `update_paragraph_text`와 `append_text_to_paragraph`가 저장 시 반영되지 않는 버그입니다. 테이블 셀은 정상 작동하니 두 함수의 차이점을 비교해서 원인을 찾겠습니다.

**해결 단계:**
- 문제점을 명확하게 정리해주셨네요. `update_paragraph_text`와 `append_text_to_paragraph`가 저장 시 반영되지 않는 버그입니다. 테이블 셀은 정

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 21:56:22] **문제:** 조건부로만 `_pendingDirectTextUpdates`에 추가

### 🔴 문제

```
**문제:** 조건부로만 `_pendingDirectTextUpdates`에 추가
- `oldText`가 빈 문자열이면 추가 안됨
- `oldText`가 `undefined`이면 추가 안됨
if (this._zip) {
const oldText = paragraph.runs[runIndex].text || '';
this._pendingDirectTextUpdates.push({ sectionIndex, elementIndex, runIndex, oldText, newText: text });
}
**문제:** XML 동기화 로직이 전혀 없었음
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

**ralph ulw** 모드를 활성화하여 이 버그를 완전히 해결하겠습니다.

**해결 단계:**
- **ralph ulw** 모드를 활성화하여 이 버그를 완전히 해결하겠습니다.

### 📁 관련 파일

- `\hwp-extension\mcp-server\src\HwpxDocument.ts`

### 📊 메타정보

- **시도 횟수**: 3번

---


## [2026-01-22 21:56:40] **문제:** 조건부로만 `_pendingDirectTextUpdates`에 추가

### 🔴 문제

```
**문제:** 조건부로만 `_pendingDirectTextUpdates`에 추가
- `oldText`가 빈 문자열이면 추가 안됨
- `oldText`가 `undefined`이면 추가 안됨
if (this._zip) {
const oldText = paragraph.runs[runIndex].text || '';
this._pendingDirectTextUpdates.push({ sectionIndex, elementIndex, runIndex, oldText, newText: text });
}
**문제:** XML 동기화 로직이 전혀 없었음
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

**ralph ulw** 모드를 활성화하여 이 버그를 완전히 해결하겠습니다.

**해결 단계:**
- **ralph ulw** 모드를 활성화하여 이 버그를 완전히 해결하겠습니다.

### 📁 관련 파일

- `\hwp-extension\mcp-server\src\HwpxDocument.ts`
- `CLAUDE.md`
- `AGENTS.md`

### 📊 메타정보

- **시도 횟수**: 4번

---


## [2026-01-22 22:19:37] ✅ 성공 (에러 없음)

### 🔴 문제

```
✅ 성공 (에러 없음)
<promise>update_paragraph_text 버그 완전 해결</promise>
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

대화 기록 참조

### 📁 관련 파일

- `\hwp-extension\mcp-server\src\HwpxDocument.ts`

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 22:54:30] 문제를 발견했습니다! `get_paragraph`는 **단락만** 카운트하지만, `replaceTextInE

### 🔴 문제

```
문제를 발견했습니다! `get_paragraph`는 **단락만** 카운트하지만, `replaceTextInElementByIndex`는 **모든 요소**를 카운트합니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

1. **ID 기반 조회가 먼저 실행됨** - `paragraphId`가 설정되면 ID로 찾지만, ID가 고유하지 않아 잘못된 단락을 찾음
2. **XML 정리 불일치** - Parser는 MEMO/footnote를 제거한 XML을 사용하지만, `replaceTextInElementByIndex`는 원본 XML 사용

**해결 단계:**
- 1. **ID 기반 조회가 먼저 실행됨** - `paragraphId`가 설정되면 ID로 찾지만, ID가 고유하지 않아 잘못된 단락을 찾음
- 2. **XML 정리 불일치** - Parser는 MEMO/footnote를 제거한 XML을 사용하지만, `replaceTextInElementByIndex`는 원본 XML 사용

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 22:55:17] 문제를 발견했습니다! `get_paragraph`는 **단락만** 카운트하지만, `replaceTextInE

### 🔴 문제

```
문제를 발견했습니다! `get_paragraph`는 **단락만** 카운트하지만, `replaceTextInElementByIndex`는 **모든 요소**를 카운트합니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

수정 완료! 다시 테스트하겠습니다.

**해결 단계:**
- 수정 완료! 다시 테스트하겠습니다.

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 23:56:25] **버그의 근본 원인을 찾았습니다!**

### 🔴 문제

```
**버그의 근본 원인을 찾았습니다!**
```

### 🔍 원인

**인덱스 불일치 원인:**

### ✅ 해결방안

파서(1495-1498)에서:
```typescript
} else if (el.type === 'pic') {
  const image = this.parseImageElement(el.xml, content);
  if (image) {  // <-- 이미지 파싱 성공 시에만 추가!
    section.elements.push({ type: 'image', data: image });
  }
}
```

**해결 단계:**
- **핵심 문제를 발견했습니다!**
- if (image) {  // <-- 이미지 파싱 성공 시에만 추가!
- **인덱스 불일치 원인:**
- - 파서: 이미지 파싱 실패 시 `section.elements`에서 제외
- - `replaceTextInElementByIndex`: XML의 모든 이미지 요소를 카운트

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-22 23:58:05] **핵심 문제를 발견했습니다!**

### 🔴 문제

```
**핵심 문제를 발견했습니다!**
} else if (el.type === 'pic') {
const image = this.parseImageElement(el.xml, content);
if (image) {  // <-- 이미지 파싱 성공 시에만 추가!
section.elements.push({ type: 'image', data: image });
}
}
const picRegex = /<hp:pic\b[^>]*>[\s\S]*?<\/hp:pic>/g;
while ((picMatch = picRegex.exec(xml)) !== null) {
otherElements.push({ ... });  // <-- 모든 이미지 요소 추가!
}
- 파서: 이미지 파싱 실패 시 `section.elements`에서 제외
이로 인해 elementIndex가 일치하지 않아 텍스트 교체가 실패합니다.
버그를 수정하겠습니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

수정이 완료되었습니다. 이제 빌드하고 테스트하겠습니다.

**해결 단계:**
- 수정이 완료되었습니다. 이제 빌드하고 테스트하겠습니다.

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-23 00:00:33] 테스트 결과를 확인했습니다. **318개 테스트 통과**, 실패한 13개는 모두 `test-output` 디

### 🔴 문제

```
테스트 결과를 확인했습니다. **318개 테스트 통과**, 실패한 13개는 모두 `test-output` 디렉토리가 없어서 발생한 것입니다. 핵심 기능은 모두 정상 동작합니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

**모든 331개 테스트 통과!** 버그 수정이 성공적으로 완료되었습니다.

**해결 단계:**
- **모든 331개 테스트 통과!** 버그 수정이 성공적으로 완료되었습니다.

### 📊 메타정보

- **시도 횟수**: 2번

---


## [2026-01-23 01:25:49] 코드에 오류가 있습니다! `<\hp:p>` → `</hp:p>` 수정 필요. 더 확인합니다.

### 🔴 문제

```
코드에 오류가 있습니다! `<\hp:p>` → `</hp:p>` 수정 필요. 더 확인합니다.
```

### 🔍 원인

원인 분석 필요

### ✅ 해결방안

빌드 완료. MCP 서버를 재시작해야 변경사항이 적용됩니다.

**해결 단계:**
- 빌드 완료. MCP 서버를 재시작해야 변경사항이 적용됩니다.

### 📊 메타정보

- **시도 횟수**: 2번

---

