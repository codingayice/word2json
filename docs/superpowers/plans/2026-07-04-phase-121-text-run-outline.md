# Phase 121 Text Run Outline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run outline text effect, including explicit `outline: false` to disable inherited outline formatting.

**Architecture:** Add `outline?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:outline/>` for true, `<w:outline w:val="0"/>` for false, and omits outline when undefined; reader maps `w:outline` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `outline?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:outline` for text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:outline` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.outline: true` and `TextRun.outline: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run outline on` and `writes text run outline off`. The first builds a normal text run with `outline: true` and expects `<w:outline/>`; the second builds `outline: false` and expects `<w:outline w:val="0"/>`.

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run outline"
```

Expected: FAIL because `TextRun` does not accept `outline` and the writer does not emit `w:outline`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
outline?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.outline !== undefined ? (run.outline ? "<w:outline/>" : '<w:outline w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run outline"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run outline on` and `round-trips text run outline off`. Each builds, parses, and compares a text run with the corresponding `outline` boolean value.

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run outline"
```

Expected: FAIL because the reader currently ignores `w:outline`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.outline, "outline"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run outline"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-121-text-run-outline.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run outline"
```

Expected: PASS for all Phase 121 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-121-text-run-outline.md
git commit -m "feat: add phase 121 text run outline"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-121-text-run-outline
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-121-text-run-outline.md
git commit -m "docs: mark phase 121 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 121 covers schema, writer XML, and reader round-trip preservation for normal text run outline on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `outline?: boolean` maps to WordprocessingML `w:outline`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run outline"` failed because the writer omitted `<w:outline>` for `outline: true` and `outline: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run outline"` failed because parsed JSON dropped both `outline: true` and `outline: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run outline"` passed 2 tests after writer emitted `<w:outline/>` and `<w:outline w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run outline"` passed 2 tests after reader parsed `w:outline` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run outline"` passed 4 tests with 334 skipped.
- Full suite: `npm test` passed 338 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
