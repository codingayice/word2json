# Phase 91 Matrix Base Justification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:baseJc` support for `matrix` math nodes so matrix baseline alignment survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `baseJustification?: "top" | "center" | "bottom"`. Writer maps the semantic JSON value `bottom` to OMML `m:val="bot"` and reader maps `bot` back to `bottom`, while preserving existing rows and control properties.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `baseJustification?: "top" | "center" | "bottom"` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mPr/m:baseJc` when `baseJustification` is present.
- Modify `src/docx-reader.ts`: parse `mPr.baseJc` into `baseJustification`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix base justification` that builds a 1x2 matrix with `baseJustification: "bottom"` and expects `m:baseJc m:val="bot"` inside `m:mPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix base justification"
```

Expected: FAIL because writer does not emit `m:baseJc`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Build matrix properties from `baseJustification` and `controlProperties`:

```ts
const properties = [
  node.baseJustification !== undefined ? `<m:baseJc m:val="${matrixBaseJustificationXml(node.baseJustification)}"/>` : "",
  controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
].join("");
```

Add the mapping helper:

```ts
function matrixBaseJustificationXml(value: Extract<MathNode, { type: "matrix" }>["baseJustification"]): string {
  return value === "bottom" ? "bot" : value ?? "center";
}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix base justification"
```

Expected: PASS with `m:baseJc m:val="bot"` while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix base justification` that builds, parses, and compares a matrix with `baseJustification: "bottom"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix base justification"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.baseJc`.

- [x] **Step 3: Implement reader support**

Read `baseJc` beside `ctrlPr`, then include it when it maps to a supported value:

```ts
const baseJustification = matrixBaseJustificationValue(asObject(matrixProperties.baseJc).val);
```

```ts
...(baseJustification ? { baseJustification } : {}),
```

Add the mapping helper:

```ts
function matrixBaseJustificationValue(value: unknown): Extract<MathNode, { type: "matrix" }>["baseJustification"] | undefined {
  if (value === "top" || value === "center") {
    return value;
  }
  if (value === "bot") {
    return "bottom";
  }
  return undefined;
}
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix base justification"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-91-matrix-base-justification.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix base justification"
```

Expected: PASS for both Phase 91 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-91-matrix-base-justification.md
git commit -m "feat: add phase 91 matrix base justification"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-91-matrix-base-justification
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-91-matrix-base-justification.md
git commit -m "docs: mark phase 91 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 91 covers writer and reader round-trip for matrix base justification while preserving existing rows and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `baseJustification` and maps to OMML `m:mPr/m:baseJc`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix base justification"` failed because `m:mPr/m:baseJc` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix base justification"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix base justification"` failed because parsed JSON did not include `baseJustification`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix base justification"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix base justification"` passed 2 tests, 272 skipped.
- Full suite: `npm test` passed 274 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-91-matrix-base-justification`
- Implementation commit: `f9e9e27 feat: add phase 91 matrix base justification`
- Remote: `origin/phase-91-matrix-base-justification`
