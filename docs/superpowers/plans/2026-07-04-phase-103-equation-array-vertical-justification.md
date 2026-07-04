# Phase 103 Equation Array Vertical Justification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:eqArrPr/m:vertJc` support for `equationArray` math nodes so equation-array vertical justification survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `equationArray` node with optional `verticalJustification?: "top" | "bottom"`. Writer emits `m:vertJc` inside `m:eqArrPr` when the property is present; reader reuses the existing vertical justification parser to map supported OMML values back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `verticalJustification?: "top" | "bottom"` on `equationArray`.
- Modify `src/docx-writer.ts`: emit `m:eqArrPr/m:vertJc` when `equationArray.verticalJustification` is present.
- Modify `src/docx-reader.ts`: parse `eqArrPr.vertJc` into `verticalJustification`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes equation array vertical justification` that builds an equation array with `verticalJustification: "bottom"` and expects `m:vertJc m:val="bottom"` inside `m:eqArrPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes equation array vertical justification"
```

Expected: FAIL because writer does not emit `m:eqArrPr/m:vertJc`.

- [x] **Step 3: Extend schema**

Update the `equationArray` node:

```ts
  | { type: "equationArray"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; verticalJustification?: "top" | "bottom"; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; objectDistribution?: boolean; maxDistribution?: boolean; rows: MathNode[][] }
```

- [x] **Step 4: Implement writer support**

Add `verticalJustification` to equation array properties:

```ts
const equationArrayProperties = [
  node.baseJustification !== undefined ? `<m:baseJc m:val="${matrixBaseJustificationXml(node.baseJustification)}"/>` : "",
  node.verticalJustification !== undefined ? `<m:vertJc m:val="${node.verticalJustification}"/>` : "",
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
npm test -- tests/docx-core.test.ts -t "writes equation array vertical justification"
```

Expected: PASS with `m:vertJc m:val="bottom"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips equation array vertical justification` that builds, parses, and compares an equation array with `verticalJustification: "bottom"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array vertical justification"
```

Expected: FAIL because `parseMathNodes` ignores `eqArrPr.vertJc`.

- [x] **Step 3: Implement reader support**

Read `vertJc` beside the other equation array properties:

```ts
const verticalJustification = groupCharacterVerticalJustificationValue(asObject(equationArrayProperties.vertJc).val);
```

```ts
...(verticalJustification ? { verticalJustification } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips equation array vertical justification"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-103-equation-array-vertical-justification.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "equation array vertical justification"
```

Expected: PASS for both Phase 103 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-103-equation-array-vertical-justification.md
git commit -m "feat: add phase 103 equation array vertical justification"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-103-equation-array-vertical-justification
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-103-equation-array-vertical-justification.md
git commit -m "docs: mark phase 103 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 103 covers writer and reader round-trip for equation array vertical justification using the same semantic values as group-character vertical justification.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `verticalJustification` and maps to OMML `m:eqArrPr/m:vertJc`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes equation array vertical justification"` failed because `m:eqArrPr/m:vertJc` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes equation array vertical justification"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array vertical justification"` failed because parsed JSON did not include `verticalJustification`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips equation array vertical justification"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "equation array vertical justification"` passed 2 tests, 296 skipped.
- Full suite: `npm test` passed 298 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
