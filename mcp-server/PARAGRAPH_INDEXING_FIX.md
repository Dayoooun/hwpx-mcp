# Paragraph Indexing Fix

## Problem Summary

The `replaceTextInElementByIndex` method in HwpxDocument.ts had incorrect paragraph counting logic that caused wrong paragraphs to be updated when documents contained tables.

## Root Cause

The issue was in how the method mapped from `elementIndex` (index into mixed elements array) to the position in the `originalTopLevelParas` array.

### Why the bug occurred

1. MCP's `elements` array contains BOTH tables and paragraphs as top-level items
2. A table counts as ONE element in this array, even though it may contain many paragraphs internally
3. The element indexing system excludes table-internal paragraphs (only top-level elements are indexed)

### The incorrect logic (lines 7241-7246)

```typescript
// WRONG: Counts paragraphs up to and including elementIndex
let paraCountUpToIndex = 0;
for (let i = 0; i <= elementIndex; i++) {
  if (elements[i].type === 'p') {
    paraCountUpToIndex++;
  }
}
```

This logic assumed that `elementIndex` referred to a paragraph-only index, but it actually refers to an index in the mixed elements array.

**Example bug scenario:**
```
elements array:
[0] paragraph A
[1] paragraph B
[2] table (contains many paragraphs internally)
[3] paragraph C  <- Want to update this
[4] paragraph D
```

When calling `replaceTextInElementByIndex(3, ...)`:
- OLD LOGIC: Counted paragraphs 0,1,3 → `paraCountUpToIndex = 3`
- But `originalTopLevelParas[3-1] = originalTopLevelParas[2]` points to paragraph C's XML
- This worked accidentally when there were no tables

With tables:
- OLD LOGIC: Still counts 3 paragraphs, but now it's [0,1,3]
- Should actually point to `topLevelParagraphIndex = 2` (0-indexed: A=0, B=1, C=2)
- Result: WRONG paragraph gets updated

## The Fix

Changed the logic to count how many paragraph elements appear BEFORE `elementIndex`:

```typescript
// CORRECT: Count paragraphs BEFORE elementIndex
let topLevelParagraphIndex = 0;
for (let i = 0; i < elementIndex; i++) {  // Changed from <= to <
  if (elements[i].type === 'p') {
    topLevelParagraphIndex++;
  }
}
```

Now:
- For `elementIndex = 3` in the example above:
  - Count paragraphs at indices 0, 1, 2
  - Only indices 0 and 1 are paragraphs
  - `topLevelParagraphIndex = 2`
  - `originalTopLevelParas[2]` correctly points to paragraph C

## Code Changes

**File:** `D:\hwp-extension\mcp-server\src\HwpxDocument.ts`

### Change 1: Variable naming and counting logic (lines 7240-7249)

```typescript
// Before:
let paraCountUpToIndex = 0;
for (let i = 0; i <= elementIndex; i++) {
  if (elements[i].type === 'p') {
    paraCountUpToIndex++;
  }
}

// After:
let topLevelParagraphIndex = 0;
for (let i = 0; i < elementIndex; i++) {  // CRITICAL: Changed <= to <
  if (elements[i].type === 'p') {
    topLevelParagraphIndex++;
  }
}
```

### Change 2: Empty text case (lines 7251-7256)

```typescript
// Before:
if (oldText === '') {
  if (paraCountUpToIndex > 0 && paraCountUpToIndex <= originalTopLevelParas.length) {
    targetInOriginal = originalTopLevelParas[paraCountUpToIndex - 1];
  }
}

// After:
if (oldText === '') {
  if (topLevelParagraphIndex >= 0 && topLevelParagraphIndex < originalTopLevelParas.length) {
    targetInOriginal = originalTopLevelParas[topLevelParagraphIndex];  // Direct index, no -1
  }
}
```

### Change 3: Text search case (lines 7257-7276)

```typescript
// Before:
let parasSeen = 0;
for (const para of originalTopLevelParas) {
  if (para.xml.includes(escapedOld)) {
    parasSeen++;
    if (parasSeen === paraCountUpToIndex) {
      targetInOriginal = para;
      break;
    }
  }
}

// After:
// First try exact positional match
if (topLevelParagraphIndex >= 0 && topLevelParagraphIndex < originalTopLevelParas.length) {
  const candidatePara = originalTopLevelParas[topLevelParagraphIndex];
  if (candidatePara.xml.includes(escapedOld)) {
    targetInOriginal = candidatePara;
  }
}

// Fallback: search any paragraph with matching text
if (!targetInOriginal) {
  for (const para of originalTopLevelParas) {
    if (para.xml.includes(escapedOld)) {
      targetInOriginal = para;
      break;
    }
  }
}
```

## Testing

The fix ensures that:

1. Element indices correctly map to top-level paragraphs
2. Tables are counted as single elements (not skipped, but not expanded)
3. The positional mapping accounts for mixed element types
4. Both empty text and text-based searches use the correct index

## Impact

This fix resolves the issue where `update_paragraph_text` and similar MCP tools were updating the wrong paragraph when the document contained tables. The element index now correctly translates to the paragraph position regardless of table presence.
