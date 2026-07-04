# Phase 122 Text Run Emboss Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run emboss text effect, including explicit `emboss: false` to disable inherited emboss formatting.

**Architecture:** Add `emboss?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:emboss/>` for true, `<w:emboss w:val="0"/>` for false, and omits emboss when undefined; reader maps `w:emboss` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `emboss?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:emboss` for text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:emboss` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.emboss: true` and `TextRun.emboss: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run emboss on` and `writes text run emboss off`. The first builds a normal text run with `emboss: true` and expects `<w:emboss/>`; the second builds `emboss: false` and expects `<w:emboss w:val="0"/>`.

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run emboss"
```

Expected: FAIL because `TextRun` does not accept `emboss` and the writer does not emit `w:emboss`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
emboss?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.emboss !== undefined ? (run.emboss ? "<w:emboss/>" : '<w:emboss w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run emboss"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run emboss on` and `round-trips text run emboss off`. Each builds, parses, and compares a text run with the corresponding `emboss` boolean value.

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run emboss"
```

Expected: FAIL because the reader currently ignores `w:emboss`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.emboss, "emboss"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run emboss"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-122-text-run-emboss.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run emboss"
```

Expected: PASS for all Phase 122 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-122-text-run-emboss.md
git commit -m "feat: add phase 122 text run emboss"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-122-text-run-emboss
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-122-text-run-emboss.md
git commit -m "docs: mark phase 122 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 122 covers schema, writer XML, and reader round-trip preservation for normal text run emboss on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `emboss?: boolean` maps to WordprocessingML `w:emboss`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run emboss"` failed because the writer omitted `<w:emboss>` for `emboss: true` and `emboss: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run emboss"` failed because parsed JSON dropped both `emboss: true` and `emboss: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run emboss"` passed 2 tests after writer emitted `<w:emboss/>` and `<w:emboss w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run emboss"` passed 2 tests after reader parsed `w:emboss` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run emboss"` passed 4 tests with 338 skipped.
- Full suite: `npm test` passed 342 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-122-text-run-emboss`
- Implementation commit: `ba517a3 feat: add phase 122 text run emboss`
- Remote: `origin/phase-122-text-run-emboss`
