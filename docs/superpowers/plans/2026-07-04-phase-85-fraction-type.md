# Phase 85 Fraction Type Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:fPr/m:type` support for `fraction` math nodes so stacked, skewed, linear, and no-bar fraction layouts survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `fraction` node with an optional `fractionType` enum that maps directly to OMML `m:type/@m:val`. Reuse the current `m:fPr` container so `fractionType` can coexist with existing `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `fractionType?: "bar" | "skewed" | "linear" | "noBar"` on `fraction`.
- Modify `src/docx-writer.ts`: emit `m:fPr/m:type` before `m:ctrlPr` when `fractionType` is present.
- Modify `src/docx-reader.ts`: parse `fPr.type.val` into `fractionType`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes fraction type` that builds a skewed fraction and expects `m:fPr/m:type`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes fraction type"
```

Expected: FAIL because `fraction.fractionType` is ignored and no `m:type` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `fraction` node:

```ts
  | { type: "fraction"; fractionType?: "bar" | "skewed" | "linear" | "noBar"; controlProperties?: MathControlProperties; numerator: MathNode[]; denominator: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add helper mapping:

```ts
function fractionTypeXml(type: Extract<MathNode, { type: "fraction" }>["fractionType"]): string {
  if (type === undefined) {
    return "";
  }
  const value = type === "skewed" ? "skw" : type;
  return `<m:type m:val="${value}"/>`;
}
```

Then build `m:fPr` from both `m:type` and `m:ctrlPr`.

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes fraction type"
```

Expected: PASS with the expected `m:fPr/m:type` XML while preserving numerator and denominator.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips fraction type` that builds, parses, and compares a linear fraction.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips fraction type"
```

Expected: FAIL because `parseMathNodes` ignores `fPr.type.val`.

- [x] **Step 3: Implement reader support**

Add helper mapping:

```ts
function fractionTypeValue(value: unknown): Extract<MathNode, { type: "fraction" }>["fractionType"] | undefined {
  if (value === "skw") {
    return "skewed";
  }
  return value === "bar" || value === "lin" || value === "noBar" ? (value === "lin" ? "linear" : value) : undefined;
}
```

Then include:

```ts
          ...(fractionType ? { fractionType } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips fraction type"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-85-fraction-type.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "fraction type"
```

Expected: PASS for both Phase 85 tests.

- [x] **Step 2: Run full test suite**

Run:

```bash
npm test
```

Expected: all tests pass.

- [x] **Step 3: Run build**

Run:

```bash
npm run build
```

Expected: build exits with code 0.

- [x] **Step 4: Run whitespace check**

Run:

```bash
git diff --check
```

Expected: exit code 0.

- [ ] **Step 5: Commit implementation**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-85-fraction-type.md
git commit -m "feat: add phase 85 fraction type"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-85-fraction-type
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-85-fraction-type.md
git commit -m "docs: mark phase 85 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 85 covers `m:fPr/m:type` writer and reader round-trip for fraction layout type while preserving numerator, denominator, and existing control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The JSON property is consistently named `fractionType` and maps to OMML values `bar`, `skw`, `lin`, and `noBar`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes fraction type"` failed because `m:fPr/m:type` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes fraction type"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips fraction type"` failed because parsed JSON did not include `fractionType`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips fraction type"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "fraction type"` passed 2 tests, 260 skipped.
- Full suite: `npm test` passed 262 tests.
- Build: first `npm run build` found a missing `MathNode` type import in `src/docx-reader.ts`; after adding the import, `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
