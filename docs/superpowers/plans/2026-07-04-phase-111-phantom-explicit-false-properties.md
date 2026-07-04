# Phase 111 Phantom Explicit False Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `false` values for Office Math `phantom` zero and transparency flags so visible dimensions and opacity round-trip as deliberate author intent.

**Architecture:** Keep the existing `zeroWidth?: boolean`, `zeroAscent?: boolean`, `zeroDescent?: boolean`, and `transparent?: boolean` schema fields. Writer emits `m:val="1"` or `m:val="0"` whenever these phantom flags are defined; reader uses optional boolean parsing so present `m:val="0"` becomes `false` instead of being dropped. Existing `show?: boolean` behavior remains unchanged because it already supports explicit false.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: write `m:zeroWid`, `m:zeroAsc`, `m:zeroDesc`, and `m:transp` for both true and false phantom values.
- Modify `src/docx-reader.ts`: parse phantom zero and transparency flags with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for explicit false phantom flags.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes phantom explicit false properties` that builds a `phantom` node with `zeroWidth`, `zeroAscent`, `zeroDescent`, and `transparent` set to `false`, then expects `m:zeroWid`, `m:zeroAsc`, `m:zeroDesc`, and `m:transp` with `m:val="0"` inside `m:phantPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes phantom explicit false properties"
```

Expected: FAIL because the writer currently omits these phantom flags when their values are false.

- [x] **Step 3: Implement writer support**

Update the phantom property list:

```ts
node.zeroWidth !== undefined ? `<m:zeroWid m:val="${node.zeroWidth ? "1" : "0"}"/>` : "",
node.zeroAscent !== undefined ? `<m:zeroAsc m:val="${node.zeroAscent ? "1" : "0"}"/>` : "",
node.zeroDescent !== undefined ? `<m:zeroDesc m:val="${node.zeroDescent ? "1" : "0"}"/>` : "",
node.transparent !== undefined ? `<m:transp m:val="${node.transparent ? "1" : "0"}"/>` : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes phantom explicit false properties"
```

Expected: PASS with all four explicit false phantom flags present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips phantom explicit false properties` that builds, parses, and compares a `phantom` node with `zeroWidth`, `zeroAscent`, `zeroDescent`, and `transparent` set to `false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips phantom explicit false properties"
```

Expected: FAIL because `parseMathNodes` currently drops false phantom zero and transparency flags by using true-only boolean parsing.

- [x] **Step 3: Implement reader support**

Change phantom flag parsing to optional boolean parsing:

```ts
...mathOptionalBooleanProperty(phantomProperties.zeroWid, "zeroWidth"),
...mathOptionalBooleanProperty(phantomProperties.zeroAsc, "zeroAscent"),
...mathOptionalBooleanProperty(phantomProperties.zeroDesc, "zeroDescent"),
...mathOptionalBooleanProperty(phantomProperties.transp, "transparent"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips phantom explicit false properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-111-phantom-explicit-false-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "phantom explicit false properties"
```

Expected: PASS for both Phase 111 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-111-phantom-explicit-false-properties.md
git commit -m "feat: add phase 111 phantom explicit false properties"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-111-phantom-explicit-false-properties
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-111-phantom-explicit-false-properties.md
git commit -m "docs: mark phase 111 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 111 covers writer and reader round-trip preservation for explicit false `phantom` zero and transparency flags.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing schema fields remain `zeroWidth`, `zeroAscent`, `zeroDescent`, and `transparent`, mapping to OMML `m:zeroWid`, `m:zeroAsc`, `m:zeroDesc`, and `m:transp`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes phantom explicit false properties"` failed because the writer omitted `m:phantPr` and all false zero/transparency flags.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips phantom explicit false properties"` failed because parsed JSON dropped `zeroWidth: false`, `zeroAscent: false`, `zeroDescent: false`, and `transparent: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes phantom explicit false properties"` passed after writer emitted explicit `m:val="0"` flags.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips phantom explicit false properties"` passed after reader used optional boolean parsing for phantom zero/transparency flags.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "phantom explicit false properties"` passed 2 tests with 312 skipped.
- Full suite: `npm test` passed 314 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
