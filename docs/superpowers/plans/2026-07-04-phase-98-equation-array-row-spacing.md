# Phase 98 Equation Array Row Spacing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:eqArrPr/m:rSp` support for `equationArray` math nodes so multi-line equation row spacing survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `equationArray` node with optional `rowSpacing?: number`. Writer emits `m:rSp` inside `m:eqArrPr` alongside existing control properties; reader parses `m:rSp` back into `rowSpacing`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `rowSpacing?: number` on `equationArray`.
- Modify `src/docx-writer.ts`: emit `m:eqArrPr/m:rSp` when `rowSpacing` is present.
- Modify `src/docx-reader.ts`: parse `eqArrPr.rSp` into `rowSpacing`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes equation array row spacing` that builds an equation array with `rowSpacing: 3` and expects `m:eqArrPr/m:rSp m:val="3"`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array row spacing"
```

Expected: FAIL because writer does not emit `m:eqArrPr/m:rSp`.

- [x] **Step 3: Extend schema**

Update the `equationArray` node:

```ts
  | { type: "equationArray"; controlProperties?: MathControlProperties; rowSpacing?: number; rows: MathNode[][] }
```

- [x] **Step 4: Implement writer support**

Build equation array properties from row spacing and control properties:

```ts
const equationArrayProperties = [
  node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
  properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array row spacing"
```

Expected: PASS with `m:rSp m:val="3"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips equation array row spacing` that builds, parses, and compares an equation array with `rowSpacing: 3`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing"
```

Expected: FAIL because `parseMathNodes` ignores `eqArrPr.rSp`.

- [x] **Step 3: Implement reader support**

Read `rSp` from `eqArrPr`, then include it when present:

```ts
const equationArrayProperties = asObject(equationArrayNode.eqArrPr);
const rowSpacing = asObject(equationArrayProperties.rSp).val;
```

```ts
...(rowSpacing !== undefined ? { rowSpacing: parseNumber(rowSpacing) } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-98-equation-array-row-spacing.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "equation array row spacing"
```

Expected: PASS for both Phase 98 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-98-equation-array-row-spacing.md
git commit -m "feat: add phase 98 equation array row spacing"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-98-equation-array-row-spacing
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-98-equation-array-row-spacing.md
git commit -m "docs: mark phase 98 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 98 covers writer and reader round-trip for equation array row spacing while preserving existing rows and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `rowSpacing` and maps to OMML `m:eqArrPr/m:rSp`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes equation array row spacing"` failed because `m:eqArrPr/m:rSp` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes equation array row spacing"` passed.
- Writer regression: `npm test -- tests/docx-core.test.ts -t "writes equation array control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing"` failed because parsed JSON did not include `rowSpacing`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing"` passed.
- Reader regression: `npm test -- tests/docx-core.test.ts -t "round-trips equation array control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "equation array row spacing"` passed 2 tests, 286 skipped.
- Full suite: `npm test` passed 288 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-98-equation-array-row-spacing`
- Implementation commit: `6454fe2 feat: add phase 98 equation array row spacing`
- Remote: `origin/phase-98-equation-array-row-spacing`
