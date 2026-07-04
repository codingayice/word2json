# Phase 112 Text Run Underline None Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `underline: false` on normal text runs so a run can deliberately disable inherited underline formatting.

**Architecture:** Keep the existing `underline?: boolean` field on `TextRun`. Writer emits `<w:u w:val="single"/>` for `true`, `<w:u w:val="none"/>` for `false`, and omits underline when undefined; reader maps `<w:u w:val="none"/>` and other explicit off values back to `underline: false`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: make `runPropertiesXml` write underline false as `w:u w:val="none"`.
- Modify `src/docx-reader.ts`: parse run underline values with optional boolean semantics instead of true-only presence.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for `TextRun.underline: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side test named `writes text run underline none` that builds a normal text run with `underline: false`, then expects `<w:u w:val="none"/>` in `word/document.xml`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run underline none"
```

Expected: FAIL because `runPropertiesXml` currently omits underline when the value is false.

- [x] **Step 3: Implement writer support**

Update `runPropertiesXml`:

```ts
run.underline !== undefined ? `<w:u w:val="${run.underline ? "single" : "none"}"/>` : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run underline none"
```

Expected: PASS with `<w:u w:val="none"/>` present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side test named `round-trips text run underline none` that builds, parses, and compares a text run with `underline: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run underline none"
```

Expected: FAIL because the writer omits false and the reader currently parses any `w:u` as `underline: true`.

- [x] **Step 3: Implement reader support**

Add a small helper:

```ts
function parseUnderline(value: unknown): Partial<Pick<TextRun, "underline">> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { underline: !(val === "none" || val === "0" || val === false) };
}
```

Use it in `parseRun`:

```ts
...parseUnderline(properties.u),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run underline none"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-112-text-run-underline-none.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run underline none"
```

Expected: PASS for both Phase 112 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-112-text-run-underline-none.md
git commit -m "feat: add phase 112 text run underline none"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-112-text-run-underline-none
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-112-text-run-underline-none.md
git commit -m "docs: mark phase 112 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 112 covers writer and reader round-trip preservation for explicit false underline on normal text runs.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing `underline?: boolean` remains the JSON API and maps to WordprocessingML `w:u`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run underline none"` failed because the writer omitted `<w:u>` for `underline: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run underline none"` failed because parsed JSON dropped `underline: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run underline none"` passed after writer emitted `<w:u w:val="none"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run underline none"` passed after reader parsed `w:u w:val="none"` as `underline: false`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run underline none"` passed 2 tests with 314 skipped.
- Full suite: `npm test` passed 316 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-112-text-run-underline-none`
- Implementation commit: `cd4318a feat: add phase 112 text run underline none`
- Remote: `origin/phase-112-text-run-underline-none`
