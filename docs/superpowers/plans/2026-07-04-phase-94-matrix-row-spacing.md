# Phase 94 Matrix Row Spacing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:rSp` support for `matrix` math nodes so matrix row spacing survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `rowSpacing?: number`. Writer emits `m:rSp` inside `m:mPr` alongside existing base justification, column properties, and control properties; reader parses `m:rSp` back into `rowSpacing`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `rowSpacing?: number` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mPr/m:rSp` when `rowSpacing` is present.
- Modify `src/docx-reader.ts`: parse `mPr.rSp` into `rowSpacing`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix row spacing` that builds a 2x1 matrix with `rowSpacing: 3` and expects `m:rSp m:val="3"` inside `m:mPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix row spacing"
```

Expected: FAIL because writer does not emit `m:rSp`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rowSpacing?: number; columnJustifications?: ("left" | "center" | "right")[]; columnCounts?: number[]; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Add `rowSpacing` to matrix properties:

```ts
node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix row spacing"
```

Expected: PASS with `m:rSp m:val="3"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix row spacing` that builds, parses, and compares a matrix with `rowSpacing: 3`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix row spacing"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.rSp`.

- [x] **Step 3: Implement reader support**

Read `rSp` beside `baseJc`, `mcs`, and `ctrlPr`, then include it when present:

```ts
const rowSpacing = asObject(matrixProperties.rSp).val;
```

```ts
...(rowSpacing !== undefined ? { rowSpacing: parseNumber(rowSpacing) } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix row spacing"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-94-matrix-row-spacing.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix row spacing"
```

Expected: PASS for both Phase 94 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-94-matrix-row-spacing.md
git commit -m "feat: add phase 94 matrix row spacing"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-94-matrix-row-spacing
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-94-matrix-row-spacing.md
git commit -m "docs: mark phase 94 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 94 covers writer and reader round-trip for matrix row spacing while preserving existing rows, base justification, column properties, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `rowSpacing` and maps to OMML `m:mPr/m:rSp`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix row spacing"` failed because `m:mPr/m:rSp` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix row spacing"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix row spacing"` failed because parsed JSON did not include `rowSpacing`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix row spacing"` passed.
- Reader regression: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column counts"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix row spacing"` passed 2 tests, 278 skipped.
- Full suite: `npm test` passed 280 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
