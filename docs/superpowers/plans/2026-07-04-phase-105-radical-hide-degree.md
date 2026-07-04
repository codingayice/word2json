# Phase 105 Radical Hide Degree Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:radPr/m:degHide` support for `radical` math nodes so explicit radical degree visibility survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `radical` node with optional `hideDegree?: boolean`. Writer emits explicit `m:degHide m:val="1"` or `m:degHide m:val="0"` inside `m:radPr` when the property is present; reader uses the existing optional math boolean parser so explicit false is preserved instead of dropped.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `hideDegree?: boolean` on `radical`.
- Modify `src/docx-writer.ts`: emit `m:radPr/m:degHide` when `radical.hideDegree` is present.
- Modify `src/docx-reader.ts`: parse `radPr.degHide` into `hideDegree`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes radical hide degree` that builds a radical with `hideDegree: false` and expects `m:degHide m:val="0"` inside `m:radPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes radical hide degree"
```

Expected: FAIL because writer does not emit `m:radPr/m:degHide`.

- [x] **Step 3: Extend schema**

Update the `radical` node:

```ts
  | { type: "radical"; controlProperties?: MathControlProperties; hideDegree?: boolean; degree?: MathNode[]; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add `hideDegree` to radical properties:

```ts
const radicalProperties = [
  node.hideDegree !== undefined ? `<m:degHide m:val="${node.hideDegree ? "1" : "0"}"/>` : "",
  properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes radical hide degree"
```

Expected: PASS with `m:degHide m:val="0"` while preserving degree and content.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips radical hide degree` that builds, parses, and compares a radical with `hideDegree: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips radical hide degree"
```

Expected: FAIL because `parseMathNodes` ignores `radPr.degHide`.

- [x] **Step 3: Implement reader support**

Read `degHide` beside radical control properties:

```ts
...mathOptionalBooleanProperty(radicalProperties.degHide, "hideDegree"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips radical hide degree"
```

Expected: PASS and parsed JSON equals the source JSON, including explicit `hideDegree: false`.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-105-radical-hide-degree.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "radical hide degree"
```

Expected: PASS for both Phase 105 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-105-radical-hide-degree.md
git commit -m "feat: add phase 105 radical hide degree"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-105-radical-hide-degree
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-105-radical-hide-degree.md
git commit -m "docs: mark phase 105 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 105 covers writer and reader round-trip for radical degree visibility while preserving explicit false values.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `hideDegree` and maps to OMML `m:radPr/m:degHide`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes radical hide degree"` failed because `m:radPr/m:degHide` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes radical hide degree"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips radical hide degree"` failed because parsed JSON did not include explicit `hideDegree: false`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips radical hide degree"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "radical hide degree"` passed 2 tests, 300 skipped.
- Full suite: `npm test` passed 302 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
