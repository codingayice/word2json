# Phase 95 Matrix Column Spacing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:cSp` support for `matrix` math nodes so matrix column spacing survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `columnSpacing?: number`. Writer emits `m:cSp` inside `m:mPr` alongside existing row spacing, base justification, column properties, and control properties; reader parses `m:cSp` back into `columnSpacing`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `columnSpacing?: number` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mPr/m:cSp` when `columnSpacing` is present.
- Modify `src/docx-reader.ts`: parse `mPr.cSp` into `columnSpacing`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix column spacing` that builds a 1x2 matrix with `columnSpacing: 4` and expects `m:cSp m:val="4"` inside `m:mPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column spacing"
```

Expected: FAIL because writer does not emit `m:cSp`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rowSpacing?: number; columnSpacing?: number; columnJustifications?: ("left" | "center" | "right")[]; columnCounts?: number[]; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Add `columnSpacing` to matrix properties:

```ts
node.columnSpacing !== undefined ? `<m:cSp m:val="${node.columnSpacing}"/>` : "",
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column spacing"
```

Expected: PASS with `m:cSp m:val="4"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix column spacing` that builds, parses, and compares a matrix with `columnSpacing: 4`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.cSp`.

- [x] **Step 3: Implement reader support**

Read `cSp` beside `rSp`, then include it when present:

```ts
const columnSpacing = asObject(matrixProperties.cSp).val;
```

```ts
...(columnSpacing !== undefined ? { columnSpacing: parseNumber(columnSpacing) } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-95-matrix-column-spacing.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix column spacing"
```

Expected: PASS for both Phase 95 tests.

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

- [x] **Step 5: Commit implementation**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-95-matrix-column-spacing.md
git commit -m "feat: add phase 95 matrix column spacing"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-95-matrix-column-spacing
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-95-matrix-column-spacing.md
git commit -m "docs: mark phase 95 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 95 covers writer and reader round-trip for matrix column spacing while preserving existing rows, row spacing, base justification, column properties, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `columnSpacing` and maps to OMML `m:mPr/m:cSp`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column spacing"` failed because `m:mPr/m:cSp` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column spacing"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing"` failed because parsed JSON did not include `columnSpacing`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column spacing"` passed.
- Reader regression: `npm test -- tests/docx-core.test.ts -t "round-trips matrix row spacing"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix column spacing"` passed 2 tests, 280 skipped.
- Full suite: `npm test` passed 282 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-95-matrix-column-spacing`
- Implementation commit: `c061eb0 feat: add phase 95 matrix column spacing`
- Remote: `origin/phase-95-matrix-column-spacing`
