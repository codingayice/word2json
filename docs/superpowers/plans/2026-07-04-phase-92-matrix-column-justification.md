# Phase 92 Matrix Column Justification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:mcs/m:mc/m:mcPr/m:mcJc` support for `matrix` math nodes so per-column matrix alignment survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `columnJustifications?: ("left" | "center" | "right")[]`. Writer maps each array entry to an OMML matrix column property entry, and reader maps `m:mcJc` values back into the same ordered array while preserving existing rows, base justification, and control properties.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `columnJustifications?: ("left" | "center" | "right")[]` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mPr/m:mcs/m:mc/m:mcPr/m:mcJc` when `columnJustifications` is present.
- Modify `src/docx-reader.ts`: parse `mPr.mcs.mc[].mcPr.mcJc` into `columnJustifications`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix column justifications` that builds a 1x3 matrix with `columnJustifications: ["left", "center", "right"]` and expects ordered `m:mcJc` entries inside `m:mcs`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column justifications"
```

Expected: FAIL because writer does not emit `m:mcs/m:mc/m:mcPr/m:mcJc`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; columnJustifications?: ("left" | "center" | "right")[]; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Build matrix column properties from `columnJustifications`:

```ts
const columnJustifications = matrixColumnJustificationsXml(node.columnJustifications);
```

Include it in matrix properties before `ctrlPr`:

```ts
columnJustifications,
```

Add the writer helper:

```ts
function matrixColumnJustificationsXml(values: Extract<MathNode, { type: "matrix" }>["columnJustifications"]): string {
  if (!values || values.length === 0) {
    return "";
  }
  return `<m:mcs>${values.map((value) => `<m:mc><m:mcPr><m:mcJc m:val="${value}"/></m:mcPr></m:mc>`).join("")}</m:mcs>`;
}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column justifications"
```

Expected: PASS with ordered left, center, and right column justification entries while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix column justifications` that builds, parses, and compares a matrix with `columnJustifications: ["left", "center", "right"]`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column justifications"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.mcs`.

- [x] **Step 3: Implement reader support**

Read `mcs.mc` beside `baseJc` and `ctrlPr`, then include it when at least one valid value is found:

```ts
const columnJustifications = matrixColumnJustificationValues(matrixProperties.mcs);
```

```ts
...(columnJustifications.length > 0 ? { columnJustifications } : {}),
```

Add the reader helpers:

```ts
function matrixColumnJustificationValues(node: unknown): NonNullable<Extract<MathNode, { type: "matrix" }>["columnJustifications"]> {
  return asArray(asObject(node).mc)
    .map((column) => matrixColumnJustificationValue(asObject(asObject(column).mcPr).mcJc))
    .filter((value): value is "left" | "center" | "right" => value !== undefined);
}

function matrixColumnJustificationValue(node: unknown): "left" | "center" | "right" | undefined {
  const value = asObject(node).val;
  return value === "left" || value === "center" || value === "right" ? value : undefined;
}
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column justifications"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-92-matrix-column-justification.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix column justifications"
```

Expected: PASS for both Phase 92 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-92-matrix-column-justification.md
git commit -m "feat: add phase 92 matrix column justification"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-92-matrix-column-justification
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-92-matrix-column-justification.md
git commit -m "docs: mark phase 92 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 92 covers writer and reader round-trip for matrix column justifications while preserving existing rows, base justification, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `columnJustifications` and maps to ordered OMML `m:mPr/m:mcs/m:mc/m:mcPr/m:mcJc` entries.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column justifications"` failed because `m:mPr/m:mcs` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column justifications"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column justifications"` failed because parsed JSON did not include `columnJustifications`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column justifications"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix column justifications"` passed 2 tests, 274 skipped.
- Full suite: `npm test` passed 276 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-92-matrix-column-justification`
- Implementation commit: `3e99c78 feat: add phase 92 matrix column justification`
- Remote: `origin/phase-92-matrix-column-justification`
