# Phase 86 N-Ary Operator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:naryPr/m:chr` support for multiple `nary.operator` values so common large operators survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `nary.operator` enum beyond `"sum"` and map each JSON value to the corresponding OMML character. Reuse the current `m:naryPr` container so operator characters continue to coexist with existing `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: extend `nary.operator` to `"sum" | "integral" | "product" | "coproduct" | "intersection" | "union"`.
- Modify `src/docx-writer.ts`: map `nary.operator` to `∑`, `∫`, `∏`, `∐`, `⋂`, or `⋃`.
- Modify `src/docx-reader.ts`: parse `naryPr.chr.val` back into the matching `nary.operator`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary operator` that builds a product n-ary node and expects `m:chr m:val="∏"`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary operator"
```

Expected: FAIL because writer currently always emits `m:chr m:val="∑"`.

- [x] **Step 3: Extend schema**

Update the `nary` node:

```ts
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add helper mapping:

```ts
function naryOperatorXml(operator: Extract<MathNode, { type: "nary" }>["operator"]): string {
  const values = {
    sum: "∑",
    integral: "∫",
    product: "∏",
    coproduct: "∐",
    intersection: "⋂",
    union: "⋃",
  } satisfies Record<Extract<MathNode, { type: "nary" }>["operator"], string>;
  return `<m:chr m:val="${values[operator]}"/>`;
}
```

Then use this helper instead of the hard-coded `∑`.

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary operator"
```

Expected: PASS with `m:chr m:val="∏"` while preserving lower limit, upper limit, and body.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary operator` that builds, parses, and compares an integral n-ary node.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary operator"
```

Expected: FAIL because `parseMathNodes` currently parses every n-ary operator as `"sum"`.

- [x] **Step 3: Implement reader support**

Update `naryOperatorValue`:

```ts
function naryOperatorValue(value: unknown): Extract<MathNode, { type: "nary" }>["operator"] {
  const values = {
    "∑": "sum",
    "∫": "integral",
    "∏": "product",
    "∐": "coproduct",
    "⋂": "intersection",
    "⋃": "union",
  } as const;
  return typeof value === "string" && value in values ? values[value as keyof typeof values] : "sum";
}
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary operator"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-86-nary-operator.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary operator"
```

Expected: PASS for both Phase 86 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-86-nary-operator.md
git commit -m "feat: add phase 86 nary operator"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-86-nary-operator
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-86-nary-operator.md
git commit -m "docs: mark phase 86 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 86 covers writer and reader round-trip for common n-ary operator characters while preserving existing limits, body, and control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The JSON property remains `operator` and maps consistently to OMML `m:naryPr/m:chr`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary operator"` failed because `operator: "product"` was emitted as `m:chr m:val="∑"`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary operator"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary operator"` failed because parsed JSON converted `operator: "integral"` to `operator: "sum"`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary operator"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "nary operator"` passed 2 tests, 262 skipped.
- Full suite: `npm test` passed 264 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
