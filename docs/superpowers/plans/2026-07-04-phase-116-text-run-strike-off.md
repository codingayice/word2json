# Phase 116 Text Run Strike Off Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `strike: false` on normal text runs so a run can deliberately disable inherited strikethrough formatting.

**Architecture:** Keep the existing `strike?: boolean` field on `TextRun`. Writer emits `<w:strike/>` for true, `<w:strike w:val="0"/>` for false, and omits strike when undefined; reader maps explicit off values back to `strike: false` using the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: make `runPropertiesXml` write strike false as `w:strike w:val="0"`.
- Modify `src/docx-reader.ts`: parse normal run strike values with optional boolean semantics instead of true-only presence.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for `TextRun.strike: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side test named `writes text run strike off` that builds a normal text run with `strike: false`, then expects `<w:strike w:val="0"/>` in `word/document.xml`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run strike off"
```

Expected: FAIL because `runPropertiesXml` currently omits strike when the value is false.

- [x] **Step 3: Implement writer support**

Update `runPropertiesXml`:

```ts
run.strike !== undefined ? (run.strike ? "<w:strike/>" : '<w:strike w:val="0"/>') : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run strike off"
```

Expected: PASS with `<w:strike w:val="0"/>` present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side test named `round-trips text run strike off` that builds, parses, and compares a text run with `strike: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run strike off"
```

Expected: FAIL because the writer omits false and the reader currently parses any `w:strike` as `strike: true`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.strike, "strike"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run strike off"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-116-text-run-strike-off.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run strike off"
```

Expected: PASS for both Phase 116 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-116-text-run-strike-off.md
git commit -m "feat: add phase 116 text run strike off"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-116-text-run-strike-off
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-116-text-run-strike-off.md
git commit -m "docs: mark phase 116 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 116 covers writer and reader round-trip preservation for explicit false strike on normal text runs.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing `strike?: boolean` remains the JSON API and maps to WordprocessingML `w:strike`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run strike off"` failed because the writer omitted `<w:strike>` for `strike: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run strike off"` failed because parsed JSON dropped `strike: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run strike off"` passed after writer emitted `<w:strike w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run strike off"` passed after reader parsed `w:strike w:val="0"` as `strike: false`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run strike off"` passed 2 tests with 322 skipped.
- Full suite: `npm test` passed 324 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
