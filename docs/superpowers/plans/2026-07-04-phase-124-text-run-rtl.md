# Phase 124 Text Run RTL Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run right-to-left direction control, including explicit `rtl: false` to disable inherited RTL formatting.

**Architecture:** Add `rtl?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:rtl/>` for true, `<w:rtl w:val="0"/>` for false, and omits RTL when undefined; reader maps `w:rtl` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `rtl?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:rtl` for text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:rtl` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.rtl: true` and `TextRun.rtl: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run rtl on` and `writes text run rtl off`. The first builds a normal text run with `rtl: true` and expects `<w:rtl/>`; the second builds `rtl: false` and expects `<w:rtl w:val="0"/>`.

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run rtl"
```

Expected: FAIL because `TextRun` does not accept `rtl` and the writer does not emit `w:rtl`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
rtl?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.rtl !== undefined ? (run.rtl ? "<w:rtl/>" : '<w:rtl w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run rtl"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run rtl on` and `round-trips text run rtl off`. Each builds, parses, and compares a text run with the corresponding `rtl` boolean value.

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run rtl"
```

Expected: FAIL because the reader currently ignores `w:rtl`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.rtl, "rtl"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run rtl"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-124-text-run-rtl.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run rtl"
```

Expected: PASS for all Phase 124 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-124-text-run-rtl.md
git commit -m "feat: add phase 124 text run rtl"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-124-text-run-rtl
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-124-text-run-rtl.md
git commit -m "docs: mark phase 124 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 124 covers schema, writer XML, and reader round-trip preservation for normal text run RTL on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `rtl?: boolean` maps to WordprocessingML `w:rtl`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run rtl"` failed because the writer omitted `<w:rtl>` for `rtl: true` and `rtl: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run rtl"` failed because parsed JSON dropped both `rtl: true` and `rtl: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run rtl"` passed 2 tests after writer emitted `<w:rtl/>` and `<w:rtl w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run rtl"` passed 2 tests after reader parsed `w:rtl` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run rtl"` passed 4 tests with 346 skipped.
- Full suite: `npm test` passed 350 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-124-text-run-rtl`
- Implementation commit: `fff5df4 feat: add phase 124 text run rtl`
- Remote: `origin/phase-124-text-run-rtl`
