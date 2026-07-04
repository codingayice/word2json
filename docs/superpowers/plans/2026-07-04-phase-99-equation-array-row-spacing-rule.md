# Phase 99 Equation Array Row Spacing Rule Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:eqArrPr/m:rSpRule` support for `equationArray` math nodes so equation-array row spacing rules survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `equationArray` node with optional `rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"`. Writer emits `m:rSpRule` inside `m:eqArrPr` next to `rowSpacing`; reader reuses the existing spacing rule parser to parse known `m:rSpRule` values back into `rowSpacingRule`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"` on `equationArray`.
- Modify `src/docx-writer.ts`: emit `m:eqArrPr/m:rSpRule` when `equationArray.rowSpacingRule` is present.
- Modify `src/docx-reader.ts`: parse `eqArrPr.rSpRule` into `rowSpacingRule`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes equation array row spacing rule` that builds an equation array with `rowSpacing: 3` and `rowSpacingRule: "exactly"`, then expects `m:rSp` followed by `m:rSpRule` inside `m:eqArrPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array row spacing rule"
```

Expected: FAIL because writer does not emit `m:eqArrPr/m:rSpRule`.

- [x] **Step 3: Extend schema**

Update the `equationArray` node:

```ts
  | { type: "equationArray"; controlProperties?: MathControlProperties; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; rows: MathNode[][] }
```

- [x] **Step 4: Implement writer support**

Add `rowSpacingRule` to equation array properties:

```ts
const equationArrayProperties = [
  node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
  node.rowSpacingRule !== undefined ? `<m:rSpRule m:val="${node.rowSpacingRule}"/>` : "",
  properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array row spacing rule"
```

Expected: PASS with `m:rSpRule m:val="exactly"` while preserving rows and row spacing.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips equation array row spacing rule` that builds, parses, and compares an equation array with `rowSpacing: 3` and `rowSpacingRule: "exactly"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing rule"
```

Expected: FAIL because `parseMathNodes` ignores `eqArrPr.rSpRule`.

- [x] **Step 3: Implement reader support**

Read `rSpRule` beside `rSp`, then include it when it maps to a supported value:

```ts
const rowSpacingRule = matrixSpacingRuleValue(asObject(equationArrayProperties.rSpRule).val);
```

```ts
...(rowSpacingRule ? { rowSpacingRule } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing rule"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-99-equation-array-row-spacing-rule.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "equation array row spacing rule"
```

Expected: PASS for both Phase 99 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-99-equation-array-row-spacing-rule.md
git commit -m "feat: add phase 99 equation array row spacing rule"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-99-equation-array-row-spacing-rule
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-99-equation-array-row-spacing-rule.md
git commit -m "docs: mark phase 99 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 99 covers writer and reader round-trip for equation array row spacing rules while preserving existing row spacing, rows, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `rowSpacingRule` and maps to OMML `m:eqArrPr/m:rSpRule`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes equation array row spacing rule"` failed because `m:eqArrPr/m:rSpRule` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes equation array row spacing rule"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing rule"` failed because parsed JSON did not include `rowSpacingRule`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array row spacing rule"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "equation array row spacing rule"` passed 2 tests, 288 skipped.
- Full suite: `npm test` passed 290 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-99-equation-array-row-spacing-rule`
- Implementation commit: `b837dd4 feat: add phase 99 equation array row spacing rule`
- Remote: `origin/phase-99-equation-array-row-spacing-rule`
