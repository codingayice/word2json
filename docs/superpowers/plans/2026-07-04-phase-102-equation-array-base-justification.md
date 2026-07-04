# Phase 102 Equation Array Base Justification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:eqArrPr/m:baseJc` support for `equationArray` math nodes so equation-array baseline alignment survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `equationArray` node with optional `baseJustification?: "top" | "center" | "bottom"`. Writer reuses the existing matrix base-justification XML mapping so JSON `bottom` writes OMML `m:val="bot"`; reader reuses the existing parser so `bot` comes back as `bottom`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `baseJustification?: "top" | "center" | "bottom"` on `equationArray`.
- Modify `src/docx-writer.ts`: emit `m:eqArrPr/m:baseJc` when `equationArray.baseJustification` is present.
- Modify `src/docx-reader.ts`: parse `eqArrPr.baseJc` into `baseJustification`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes equation array base justification` that builds an equation array with `baseJustification: "bottom"` and expects `m:baseJc m:val="bot"` inside `m:eqArrPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array base justification"
```

Expected: FAIL because writer does not emit `m:eqArrPr/m:baseJc`.

- [x] **Step 3: Extend schema**

Update the `equationArray` node:

```ts
  | { type: "equationArray"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; objectDistribution?: boolean; maxDistribution?: boolean; rows: MathNode[][] }
```

- [x] **Step 4: Implement writer support**

Add `baseJustification` to equation array properties:

```ts
const equationArrayProperties = [
  node.baseJustification !== undefined ? `<m:baseJc m:val="${matrixBaseJustificationXml(node.baseJustification)}"/>` : "",
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
npm test -- tests/docx-core.test.ts -t "writes equation array base justification"
```

Expected: PASS with `m:baseJc m:val="bot"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips equation array base justification` that builds, parses, and compares an equation array with `baseJustification: "bottom"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array base justification"
```

Expected: FAIL because `parseMathNodes` ignores `eqArrPr.baseJc`.

- [x] **Step 3: Implement reader support**

Read `baseJc` beside the other equation array properties:

```ts
const baseJustification = matrixBaseJustificationValue(asObject(equationArrayProperties.baseJc).val);
```

```ts
...(baseJustification ? { baseJustification } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array base justification"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-102-equation-array-base-justification.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "equation array base justification"
```

Expected: PASS for both Phase 102 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-102-equation-array-base-justification.md
git commit -m "feat: add phase 102 equation array base justification"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-102-equation-array-base-justification
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-102-equation-array-base-justification.md
git commit -m "docs: mark phase 102 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 102 covers writer and reader round-trip for equation array base justification using the same semantic values as matrix base justification.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `baseJustification` and maps to OMML `m:eqArrPr/m:baseJc`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes equation array base justification"` failed because `m:eqArrPr/m:baseJc` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes equation array base justification"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array base justification"` failed because parsed JSON did not include `baseJustification`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array base justification"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "equation array base justification"` passed 2 tests, 294 skipped.
- Full suite: `npm test` passed 296 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
