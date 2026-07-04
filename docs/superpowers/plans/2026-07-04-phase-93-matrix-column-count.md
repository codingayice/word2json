# Phase 93 Matrix Column Count Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:mPr/m:mcs/m:mc/m:mcPr/m:count` support for `matrix` math nodes so matrix column repeat counts survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `matrix` node with optional `columnCounts?: number[]`. Writer merges `columnCounts` and existing `columnJustifications` into ordered `m:mc` entries, and reader maps `m:count` values back into the same ordered count array while preserving rows, base justification, column justification, and control properties.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `columnCounts?: number[]` on `matrix`.
- Modify `src/docx-writer.ts`: emit `m:mcPr/m:count` when `columnCounts` is present.
- Modify `src/docx-reader.ts`: parse `mPr.mcs.mc[].mcPr.count` into `columnCounts`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes matrix column counts` that builds a 1x3 matrix with `columnCounts: [1, 2, 1]` and expects ordered `m:count` entries inside `m:mcs`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column counts"
```

Expected: FAIL because writer does not emit `m:mcPr/m:count`.

- [x] **Step 3: Extend schema**

Update the `matrix` node:

```ts
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; columnJustifications?: ("left" | "center" | "right")[]; columnCounts?: number[]; rows: MathNode[][][] }
```

- [x] **Step 4: Implement writer support**

Pass counts into the column property helper:

```ts
const columnProperties = matrixColumnPropertiesXml(node.columnJustifications, node.columnCounts);
```

Add the writer helper:

```ts
function matrixColumnPropertiesXml(
  justifications: Extract<MathNode, { type: "matrix" }>["columnJustifications"],
  counts: Extract<MathNode, { type: "matrix" }>["columnCounts"],
): string {
  const columnCount = Math.max(justifications?.length ?? 0, counts?.length ?? 0);
  if (columnCount === 0) {
    return "";
  }
  return `<m:mcs>${Array.from({ length: columnCount }, (_, index) => {
    const count = counts?.[index];
    const justification = justifications?.[index];
    const properties = [
      count !== undefined ? `<m:count m:val="${count}"/>` : "",
      justification !== undefined ? `<m:mcJc m:val="${justification}"/>` : "",
    ].join("");
    return `<m:mc><m:mcPr>${properties}</m:mcPr></m:mc>`;
  }).join("")}</m:mcs>`;
}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes matrix column counts"
```

Expected: PASS with ordered count entries while preserving rows.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips matrix column counts` that builds, parses, and compares a matrix with `columnCounts: [1, 2, 1]`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column counts"
```

Expected: FAIL because `parseMathNodes` ignores `mPr.mcs.mc[].mcPr.count`.

- [x] **Step 3: Implement reader support**

Read counts beside column justifications, then include them when at least one valid number is found:

```ts
const columnCounts = matrixColumnCountValues(matrixProperties.mcs);
```

```ts
...(columnCounts.length > 0 ? { columnCounts } : {}),
```

Add the reader helper:

```ts
function matrixColumnCountValues(node: unknown): number[] {
  return asArray(asObject(node).mc)
    .map((column) => integerValue(asObject(asObject(column).mcPr).count))
    .filter((value): value is number => value !== undefined);
}
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips matrix column counts"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-93-matrix-column-count.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "matrix column counts"
```

Expected: PASS for both Phase 93 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-93-matrix-column-count.md
git commit -m "feat: add phase 93 matrix column count"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-93-matrix-column-count
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-93-matrix-column-count.md
git commit -m "docs: mark phase 93 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 93 covers writer and reader round-trip for matrix column counts while preserving existing rows, base justification, column justification, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `columnCounts` and maps to ordered OMML `m:mPr/m:mcs/m:mc/m:mcPr/m:count` entries.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column counts"` failed because `m:mcPr/m:count` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes matrix column counts"` passed.
- Writer regression: `npm test -- tests/docx-core.test.ts -t "writes matrix column justifications"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column counts"` failed because parsed JSON did not include `columnCounts`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column counts"` passed.
- Reader regression: `npm test -- tests/docx-core.test.ts -t "round-trips matrix column justifications"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "matrix column counts"` passed 2 tests, 276 skipped.
- Full suite: `npm test` passed 278 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-93-matrix-column-count`
- Implementation commit: `31c8275 feat: add phase 93 matrix column count`
- Remote: `origin/phase-93-matrix-column-count`
