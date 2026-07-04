# Phase 113 Style Run Underline None Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `run.underline: false` in style definitions so styles can deliberately disable inherited underline formatting.

**Architecture:** Reuse the existing `StyleRunProperties.underline?: boolean` field. Writer emits `<w:u w:val="single"/>` for true, `<w:u w:val="none"/>` for false, and omits underline when undefined; reader shares underline value parsing with normal text runs.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: make `styleRunPropertiesXml` write underline false as `w:u w:val="none"`.
- Modify `src/docx-reader.ts`: make `parseStyleRunProperties` parse `w:u w:val="none"` as `underline: false`.
- Modify `tests/docx-core.test.ts`: add one styles XML writer test and one round-trip test for character style `run.underline: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side test named `writes character style underline none` that builds a character style with `run: { underline: false }`, then expects `<w:rPr><w:u w:val="none"/></w:rPr>` in `word/styles.xml`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes character style underline none"
```

Expected: FAIL because `styleRunPropertiesXml` currently omits underline when the value is false.

- [x] **Step 3: Implement writer support**

Update `styleRunPropertiesXml`:

```ts
run.underline !== undefined ? `<w:u w:val="${run.underline ? "single" : "none"}"/>` : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes character style underline none"
```

Expected: PASS with `<w:u w:val="none"/>` present in `styles.xml`.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side test named `round-trips character style underline none` that builds, parses, and compares a character style with `run: { underline: false }`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips character style underline none"
```

Expected: FAIL because the writer omits false and the reader currently parses any style `w:u` as `underline: true`.

- [x] **Step 3: Implement reader support**

Use the existing `parseUnderline` helper in `parseStyleRunProperties`:

```ts
...parseUnderline(properties.u),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips character style underline none"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-113-style-run-underline-none.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "character style underline none"
```

Expected: PASS for both Phase 113 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-113-style-run-underline-none.md
git commit -m "feat: add phase 113 style run underline none"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-113-style-run-underline-none
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-113-style-run-underline-none.md
git commit -m "docs: mark phase 113 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 113 covers writer and reader round-trip preservation for explicit false underline in style run properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing `underline?: boolean` remains the JSON API and maps to WordprocessingML `w:u`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes character style underline none"` failed because style run properties omitted `<w:u>` for `underline: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips character style underline none"` failed because parsed JSON dropped `run.underline: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes character style underline none"` passed after style writer emitted `<w:u w:val="none"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips character style underline none"` passed after style reader parsed `w:u w:val="none"` as `underline: false`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "character style underline none"` passed 2 tests with 316 skipped.
- Full suite: `npm test` passed 318 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
