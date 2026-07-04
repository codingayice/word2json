# Phase 97 Matrix Column Spacing Rule Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:cSpRule` support for `matrix` math nodes so matrix column spacing rules survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `columnSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"`. Writer emits `m:cSpRule` inside `m:mPr` next to `columnSpacing`; reader reuses the existing matrix spacing rule parser to parse `m:cSpRule`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `columnSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mPr/m:cSpRule` when `columnSpacingRule` is present.
- Modify `src/docx-reader.ts`: parse `mPr.cSpRule` into `columnSpacingRule`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix column spacing rule` that builds a 1x2 matrix with `columnSpacing: 4` and `columnSpacingRule: "multiple"`, then expects `m:cSp` followed by `m:cSpRule`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column spacing rule"
```

Expected: FAIL because writer does not emit `m:cSpRule`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; columnSpacing?: number; columnSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; columnJustifications?: ("left" | "center" | "right")[]; columnCounts?: number[]; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Add `columnSpacingRule` to matrix properties:

```ts
node.columnSpacingRule !== undefined ? `<m:cSpRule m:val="${node.columnSpacingRule}"/>` : "",
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column spacing rule"
```

Expected: PASS with `m:cSpRule m:val="multiple"` while preserving rows and column spacing.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix column spacing rule` that builds, parses, and compares a matrix with `columnSpacing: 4` and `columnSpacingRule: "multiple"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing rule"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.cSpRule`.

- [x] **Step 3: Implement reader support**

Read `cSpRule` beside `cSp`, then include it when it maps to a supported value:

```ts
const columnSpacingRule = matrixSpacingRuleValue(asObject(matrixProperties.cSpRule).val);
```

```ts
...(columnSpacingRule ? { columnSpacingRule } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing rule"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-97-matrix-column-spacing-rule.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix column spacing rule"
```

Expected: PASS for both Phase 97 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-97-matrix-column-spacing-rule.md
git commit -m "feat: add phase 97 matrix column spacing rule"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-97-matrix-column-spacing-rule
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-97-matrix-column-spacing-rule.md
git commit -m "docs: mark phase 97 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 97 covers writer and reader round-trip for matrix column spacing rules while preserving existing rows, row spacing/rule, column spacing, column properties, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `columnSpacingRule` and maps to OMML `m:mPr/m:cSpRule`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column spacing rule"` failed because `m:mPr/m:cSpRule` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column spacing rule"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing rule"` failed because parsed JSON did not include `columnSpacingRule`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing rule"` passed.
- Reader regression: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix column spacing rule"` passed 2 tests, 284 skipped.
- Full suite: `npm test` passed 286 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
