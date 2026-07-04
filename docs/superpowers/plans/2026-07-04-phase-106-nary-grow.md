# Phase 106 N-Ary Grow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:naryPr/m:grow` support for `nary` math nodes so large-operator growth settings survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `nary` node with optional `grow?: boolean`. Writer emits explicit `m:grow m:val="1"` or `m:grow m:val="0"` inside `m:naryPr` when the property is present; reader uses the existing optional math boolean parser so explicit false is preserved instead of dropped.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `grow?: boolean` on `nary`.
- Modify `src/docx-writer.ts`: emit `m:naryPr/m:grow` when `nary.grow` is present.
- Modify `src/docx-reader.ts`: parse `naryPr.grow` into `grow`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary grow` that builds a summation with `grow: false` and expects `m:grow m:val="0"` inside `m:naryPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary grow"
```

Expected: FAIL because writer does not emit `m:naryPr/m:grow`.

- [x] **Step 3: Extend schema**

Update the `nary` node:

```ts
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; limitLocation?: "underOver" | "subSup"; grow?: boolean; hideLowerLimit?: boolean; hideUpperLimit?: boolean; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add `grow` to n-ary properties:

```ts
const naryProperties = [
  naryOperatorXml(node.operator),
  naryLimitLocationXml(node.limitLocation),
  node.grow !== undefined ? `<m:grow m:val="${node.grow ? "1" : "0"}"/>` : "",
  node.hideLowerLimit ? '<m:subHide m:val="1"/>' : "",
  node.hideUpperLimit ? '<m:supHide m:val="1"/>' : "",
  properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary grow"
```

Expected: PASS with `m:grow m:val="0"` while preserving limits and body.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary grow` that builds, parses, and compares an n-ary node with `grow: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary grow"
```

Expected: FAIL because `parseMathNodes` ignores `naryPr.grow`.

- [x] **Step 3: Implement reader support**

Read `grow` beside the other n-ary properties:

```ts
...mathOptionalBooleanProperty(naryProperties.grow, "grow"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary grow"
```

Expected: PASS and parsed JSON equals the source JSON, including explicit `grow: false`.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-106-nary-grow.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary grow"
```

Expected: PASS for both Phase 106 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-106-nary-grow.md
git commit -m "feat: add phase 106 nary grow"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-106-nary-grow
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-106-nary-grow.md
git commit -m "docs: mark phase 106 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 106 covers writer and reader round-trip for n-ary operator growth while preserving explicit false values.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `grow` and maps to OMML `m:naryPr/m:grow`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary grow"` failed because `m:naryPr/m:grow` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary grow"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary grow"` failed because parsed JSON did not include explicit `grow: false`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary grow"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "nary grow"` passed 2 tests, 302 skipped.
- Full suite: `npm test` passed 304 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-106-nary-grow`
- Implementation commit: `ba30137 feat: add phase 106 nary grow`
- Remote: `origin/phase-106-nary-grow`
