# Phase 123 Text Run Imprint Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run imprint text effect, including explicit `imprint: false` to disable inherited imprint formatting.

**Architecture:** Add `imprint?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:imprint/>` for true, `<w:imprint w:val="0"/>` for false, and omits imprint when undefined; reader maps `w:imprint` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `imprint?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:imprint` for text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:imprint` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.imprint: true` and `TextRun.imprint: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run imprint on` and `writes text run imprint off`. The first builds a normal text run with `imprint: true` and expects `<w:imprint/>`; the second builds `imprint: false` and expects `<w:imprint w:val="0"/>`.

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run imprint"
```

Expected: FAIL because `TextRun` does not accept `imprint` and the writer does not emit `w:imprint`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
imprint?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.imprint !== undefined ? (run.imprint ? "<w:imprint/>" : '<w:imprint w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run imprint"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run imprint on` and `round-trips text run imprint off`. Each builds, parses, and compares a text run with the corresponding `imprint` boolean value.

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run imprint"
```

Expected: FAIL because the reader currently ignores `w:imprint`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.imprint, "imprint"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run imprint"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-123-text-run-imprint.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run imprint"
```

Expected: PASS for all Phase 123 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-123-text-run-imprint.md
git commit -m "feat: add phase 123 text run imprint"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-123-text-run-imprint
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-123-text-run-imprint.md
git commit -m "docs: mark phase 123 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 123 covers schema, writer XML, and reader round-trip preservation for normal text run imprint on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `imprint?: boolean` maps to WordprocessingML `w:imprint`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run imprint"` failed because the writer omitted `<w:imprint>` for `imprint: true` and `imprint: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run imprint"` failed because parsed JSON dropped both `imprint: true` and `imprint: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run imprint"` passed 2 tests after writer emitted `<w:imprint/>` and `<w:imprint w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run imprint"` passed 2 tests after reader parsed `w:imprint` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run imprint"` passed 4 tests with 342 skipped.
- Full suite: `npm test` passed 346 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
