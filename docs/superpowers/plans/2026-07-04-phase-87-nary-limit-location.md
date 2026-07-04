# Phase 87 N-Ary Limit Location Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:naryPr/m:limLoc` support for `nary` math nodes so large operator limit placement survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `nary` node with optional `limitLocation?: "underOver" | "subSup"`. Reuse the current `m:naryPr` container so `limLoc` can coexist with operator `chr` and `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `limitLocation?: "underOver" | "subSup"` on `nary`.
- Modify `src/docx-writer.ts`: emit `m:naryPr/m:limLoc` after `m:chr` when `limitLocation` is present.
- Modify `src/docx-reader.ts`: parse `naryPr.limLoc.val` into `limitLocation`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary limit location` that builds a sum with `limitLocation: "underOver"` and expects `m:limLoc m:val="undOvr"`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary limit location"
```

Expected: FAIL because writer does not emit `m:limLoc`.

- [x] **Step 3: Extend schema**

Update the `nary` node:

```ts
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; limitLocation?: "underOver" | "subSup"; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add helper mapping:

```ts
function naryLimitLocationXml(location: Extract<MathNode, { type: "nary" }>["limitLocation"]): string {
  if (location === undefined) {
    return "";
  }
  return `<m:limLoc m:val="${location === "underOver" ? "undOvr" : "subSup"}"/>`;
}
```

Then emit this helper inside `m:naryPr`.

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary limit location"
```

Expected: PASS with `m:limLoc m:val="undOvr"` while preserving operator, lower limit, upper limit, and body.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary limit location` that builds, parses, and compares an n-ary node with `limitLocation: "subSup"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary limit location"
```

Expected: FAIL because `parseMathNodes` ignores `naryPr.limLoc.val`.

- [x] **Step 3: Implement reader support**

Add helper mapping:

```ts
function naryLimitLocationValue(value: unknown): Extract<MathNode, { type: "nary" }>["limitLocation"] | undefined {
  if (value === "undOvr") {
    return "underOver";
  }
  return value === "subSup" ? "subSup" : undefined;
}
```

Then include:

```ts
          ...(limitLocation ? { limitLocation } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary limit location"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-87-nary-limit-location.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary limit location"
```

Expected: PASS for both Phase 87 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-87-nary-limit-location.md
git commit -m "feat: add phase 87 nary limit location"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-87-nary-limit-location
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-87-nary-limit-location.md
git commit -m "docs: mark phase 87 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 87 covers writer and reader round-trip for n-ary limit placement while preserving operator, limits, body, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The JSON property is consistently named `limitLocation` and maps to OMML `undOvr` and `subSup`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary limit location"` failed because `m:naryPr/m:limLoc` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary limit location"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary limit location"` failed because parsed JSON did not include `limitLocation`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary limit location"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "nary limit location"` passed 2 tests, 264 skipped.
- Full suite: `npm test` passed 266 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-87-nary-limit-location`
- Implementation commit: `239f0d3 feat: add phase 87 nary limit location`
- Remote: `origin/phase-87-nary-limit-location`
