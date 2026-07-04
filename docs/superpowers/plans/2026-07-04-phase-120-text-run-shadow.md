# Phase 120 Text Run Shadow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run shadow text effect, including explicit `shadow: false` to disable inherited shadow formatting.

**Architecture:** Add `shadow?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:shadow/>` for true, `<w:shadow w:val="0"/>` for false, and omits shadow when undefined; reader maps `w:shadow` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `shadow?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:shadow` for text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:shadow` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.shadow: true` and `TextRun.shadow: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run shadow on` and `writes text run shadow off`. The first builds a normal text run with `shadow: true` and expects `<w:shadow/>`; the second builds `shadow: false` and expects `<w:shadow w:val="0"/>`.

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run shadow"
```

Expected: FAIL because `TextRun` does not accept `shadow` and the writer does not emit `w:shadow`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
shadow?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.shadow !== undefined ? (run.shadow ? "<w:shadow/>" : '<w:shadow w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run shadow"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run shadow on` and `round-trips text run shadow off`. Each builds, parses, and compares a text run with the corresponding `shadow` boolean value.

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run shadow"
```

Expected: FAIL because the reader currently ignores `w:shadow`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.shadow, "shadow"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run shadow"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-120-text-run-shadow.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run shadow"
```

Expected: PASS for all Phase 120 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-120-text-run-shadow.md
git commit -m "feat: add phase 120 text run shadow"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-120-text-run-shadow
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-120-text-run-shadow.md
git commit -m "docs: mark phase 120 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 120 covers schema, writer XML, and reader round-trip preservation for normal text run shadow on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `shadow?: boolean` maps to WordprocessingML `w:shadow`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run shadow"` failed because the writer omitted `<w:shadow>` for `shadow: true` and `shadow: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run shadow"` failed because parsed JSON dropped both `shadow: true` and `shadow: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run shadow"` passed 2 tests after writer emitted `<w:shadow/>` and `<w:shadow w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run shadow"` passed 2 tests after reader parsed `w:shadow` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run shadow"` passed 4 tests with 330 skipped.
- Full suite: `npm test` passed 334 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
