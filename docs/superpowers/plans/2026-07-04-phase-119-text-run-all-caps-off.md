# Phase 119 Text Run All Caps Off Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `allCaps: false` on normal text runs so a run can deliberately disable inherited all-caps formatting.

**Architecture:** Keep the existing `allCaps?: boolean` field on `TextRun`. Writer emits `<w:caps/>` for true, `<w:caps w:val="0"/>` for false, and omits caps when undefined; reader maps explicit off values back to `allCaps: false` using the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: make `runPropertiesXml` write all caps false as `w:caps w:val="0"`.
- Modify `src/docx-reader.ts`: parse normal run all caps values with optional boolean semantics instead of true-only presence.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for `TextRun.allCaps: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side test named `writes text run all caps off` that builds a normal text run with `allCaps: false`, then expects `<w:caps w:val="0"/>` in `word/document.xml`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run all caps off"
```

Expected: FAIL because `runPropertiesXml` currently omits all caps when the value is false.

- [x] **Step 3: Implement writer support**

Update `runPropertiesXml`:

```ts
run.allCaps !== undefined ? (run.allCaps ? "<w:caps/>" : '<w:caps w:val="0"/>') : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run all caps off"
```

Expected: PASS with `<w:caps w:val="0"/>` present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side test named `round-trips text run all caps off` that builds, parses, and compares a text run with `allCaps: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run all caps off"
```

Expected: FAIL because the writer omits false and the reader currently parses any `w:caps` as `allCaps: true`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.caps, "allCaps"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run all caps off"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-119-text-run-all-caps-off.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run all caps off"
```

Expected: PASS for both Phase 119 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-119-text-run-all-caps-off.md
git commit -m "feat: add phase 119 text run all caps off"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-119-text-run-all-caps-off
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-119-text-run-all-caps-off.md
git commit -m "docs: mark phase 119 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 119 covers writer and reader round-trip preservation for explicit false all caps on normal text runs.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing `allCaps?: boolean` remains the JSON API and maps to WordprocessingML `w:caps`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run all caps off"` failed because the writer omitted `<w:caps>` for `allCaps: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run all caps off"` failed because parsed JSON dropped `allCaps: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run all caps off"` passed after writer emitted `<w:caps w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run all caps off"` passed after reader parsed `w:caps w:val="0"` as `allCaps: false`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run all caps off"` passed 2 tests with 328 skipped.
- Full suite: `npm test` passed 330 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-119-text-run-all-caps-off`
- Implementation commit: `c791f04 feat: add phase 119 text run all caps off`
- Remote: `origin/phase-119-text-run-all-caps-off`
