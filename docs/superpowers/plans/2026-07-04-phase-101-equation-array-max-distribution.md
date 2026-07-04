# Phase 101 Equation Array Max Distribution Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:eqArrPr/m:maxDist` support for `equationArray` math nodes so equation-array maximum distribution settings survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `equationArray` node with optional `maxDistribution?: boolean`. Writer emits explicit `m:maxDist m:val="1"` or `m:maxDist m:val="0"` inside `m:eqArrPr` when the property is present; reader uses the existing optional math boolean parser so explicit false is preserved instead of dropped.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `maxDistribution?: boolean` on `equationArray`.
- Modify `src/docx-writer.ts`: emit `m:eqArrPr/m:maxDist` when `equationArray.maxDistribution` is present.
- Modify `src/docx-reader.ts`: parse `eqArrPr.maxDist` into `maxDistribution`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes equation array max distribution` that builds an equation array with `maxDistribution: false` and expects `m:maxDist m:val="0"` inside `m:eqArrPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array max distribution"
```

Expected: FAIL because writer does not emit `m:eqArrPr/m:maxDist`.

- [x] **Step 3: Extend schema**

Update the `equationArray` node:

```ts
  | { type: "equationArray"; controlProperties?: MathControlProperties; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; objectDistribution?: boolean; maxDistribution?: boolean; rows: MathNode[][] }
```

- [x] **Step 4: Implement writer support**

Add `maxDistribution` to equation array properties:

```ts
const equationArrayProperties = [
  node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
  node.rowSpacingRule !== undefined ? `<m:rSpRule m:val="${node.rowSpacingRule}"/>` : "",
  node.objectDistribution !== undefined ? `<m:objDist m:val="${node.objectDistribution ? "1" : "0"}"/>` : "",
  node.maxDistribution !== undefined ? `<m:maxDist m:val="${node.maxDistribution ? "1" : "0"}"/>` : "",
  properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array max distribution"
```

Expected: PASS with `m:maxDist m:val="0"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips equation array max distribution` that builds, parses, and compares an equation array with `maxDistribution: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array max distribution"
```

Expected: FAIL because `parseMathNodes` ignores `eqArrPr.maxDist`.

- [x] **Step 3: Implement reader support**

Read `maxDist` beside the other equation array properties:

```ts
...mathOptionalBooleanProperty(equationArrayProperties.maxDist, "maxDistribution"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array max distribution"
```

Expected: PASS and parsed JSON equals the source JSON, including explicit `maxDistribution: false`.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-101-equation-array-max-distribution.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "equation array max distribution"
```

Expected: PASS for both Phase 101 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-101-equation-array-max-distribution.md
git commit -m "feat: add phase 101 equation array max distribution"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-101-equation-array-max-distribution
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-101-equation-array-max-distribution.md
git commit -m "docs: mark phase 101 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 101 covers writer and reader round-trip for equation array maximum distribution while preserving explicit false values.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `maxDistribution` and maps to OMML `m:eqArrPr/m:maxDist`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes equation array max distribution"` failed because `m:eqArrPr/m:maxDist` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes equation array max distribution"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array max distribution"` failed because parsed JSON did not include explicit `maxDistribution: false`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array max distribution"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "equation array max distribution"` passed 2 tests, 292 skipped.
- Full suite: `npm test` passed 294 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
