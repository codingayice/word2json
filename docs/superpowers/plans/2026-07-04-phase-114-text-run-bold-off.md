# Phase 114 Text Run Bold Off Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `bold: false` on normal text runs so a run can deliberately disable inherited bold formatting.

**Architecture:** Keep the existing `bold?: boolean` field on `TextRun`. Writer emits `<w:b/>` for true, `<w:b w:val="0"/>` for false, and omits bold when undefined; reader maps explicit off values back to `bold: false`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: make `runPropertiesXml` write bold false as `w:b w:val="0"`.
- Modify `src/docx-reader.ts`: parse normal run bold values with optional boolean semantics instead of true-only presence.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for `TextRun.bold: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side test named `writes text run bold off` that builds a normal text run with `bold: false`, then expects `<w:b w:val="0"/>` in `word/document.xml`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run bold off"
```

Expected: FAIL because `runPropertiesXml` currently omits bold when the value is false.

- [x] **Step 3: Implement writer support**

Update `runPropertiesXml`:

```ts
run.bold !== undefined ? (run.bold ? "<w:b/>" : '<w:b w:val="0"/>') : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run bold off"
```

Expected: PASS with `<w:b w:val="0"/>` present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side test named `round-trips text run bold off` that builds, parses, and compares a text run with `bold: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run bold off"
```

Expected: FAIL because the writer omits false and the reader currently parses any `w:b` as `bold: true`.

- [x] **Step 3: Implement reader support**

Add a helper:

```ts
function parseOnOffRunProperty<K extends keyof TextRun>(value: unknown, key: K): Partial<Pick<TextRun, K>> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { [key]: !(val === "0" || val === false || val === "false") } as Partial<Pick<TextRun, K>>;
}
```

Use it in `parseRun`:

```ts
...parseOnOffRunProperty(properties.b, "bold"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run bold off"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-114-text-run-bold-off.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run bold off"
```

Expected: PASS for both Phase 114 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-114-text-run-bold-off.md
git commit -m "feat: add phase 114 text run bold off"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-114-text-run-bold-off
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-114-text-run-bold-off.md
git commit -m "docs: mark phase 114 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 114 covers writer and reader round-trip preservation for explicit false bold on normal text runs.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing `bold?: boolean` remains the JSON API and maps to WordprocessingML `w:b`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run bold off"` failed because the writer omitted `<w:b>` for `bold: false`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run bold off"` failed because parsed JSON dropped `bold: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run bold off"` passed after writer emitted `<w:b w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run bold off"` passed after reader parsed `w:b w:val="0"` as `bold: false`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run bold off"` passed 2 tests with 318 skipped.
- Full suite: `npm test` passed 320 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-114-text-run-bold-off`
- Implementation commit: `e391675 feat: add phase 114 text run bold off`
- Remote: `origin/phase-114-text-run-bold-off`
