# Phase 109 Box Explicit Visible Borders Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `false` values for Office Math `box` border hide flags so visible borders round-trip as deliberate author intent.

**Architecture:** Keep the existing `hideTop?: boolean`, `hideBottom?: boolean`, `hideLeft?: boolean`, and `hideRight?: boolean` schema fields. Writer emits `m:val="1"` or `m:val="0"` whenever a box hide flag is defined; reader uses optional boolean parsing for `box` hide flags so present `m:val="0"` becomes `false` instead of being dropped.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: write `m:hideTop`, `m:hideBot`, `m:hideLeft`, and `m:hideRight` for both true and false box values.
- Modify `src/docx-reader.ts`: parse box hide flags with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for explicit visible box borders.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes box explicit visible borders` that builds a `box` node with all four hide flags set to `false`, then expects `m:hideTop`, `m:hideBot`, `m:hideLeft`, and `m:hideRight` with `m:val="0"` inside `m:boxPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes box explicit visible borders"
```

Expected: FAIL because the writer currently omits box hide flags when their values are false.

- [x] **Step 3: Implement writer support**

Update the box property list:

```ts
node.hideTop !== undefined ? `<m:hideTop m:val="${node.hideTop ? "1" : "0"}"/>` : "",
node.hideBottom !== undefined ? `<m:hideBot m:val="${node.hideBottom ? "1" : "0"}"/>` : "",
node.hideLeft !== undefined ? `<m:hideLeft m:val="${node.hideLeft ? "1" : "0"}"/>` : "",
node.hideRight !== undefined ? `<m:hideRight m:val="${node.hideRight ? "1" : "0"}"/>` : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes box explicit visible borders"
```

Expected: PASS with all four explicit visible border flags present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips box explicit visible borders` that builds, parses, and compares a `box` node with all four hide flags set to `false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips box explicit visible borders"
```

Expected: FAIL because `parseMathNodes` currently drops false box hide flags by using true-only boolean parsing.

- [x] **Step 3: Implement reader support**

Change box hide parsing to optional boolean parsing:

```ts
...mathOptionalBooleanProperty(boxProperties.hideTop, "hideTop"),
...mathOptionalBooleanProperty(boxProperties.hideBot, "hideBottom"),
...mathOptionalBooleanProperty(boxProperties.hideLeft, "hideLeft"),
...mathOptionalBooleanProperty(boxProperties.hideRight, "hideRight"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips box explicit visible borders"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-109-box-explicit-visible-borders.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "box explicit visible borders"
```

Expected: PASS for both Phase 109 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-109-box-explicit-visible-borders.md
git commit -m "feat: add phase 109 box explicit visible borders"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-109-box-explicit-visible-borders
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-109-box-explicit-visible-borders.md
git commit -m "docs: mark phase 109 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 109 covers writer and reader round-trip preservation for explicit false `box` border hide flags.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing schema fields remain `hideTop`, `hideBottom`, `hideLeft`, and `hideRight`, mapping to OMML `m:hideTop`, `m:hideBot`, `m:hideLeft`, and `m:hideRight`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes box explicit visible borders"` failed because the writer omitted `m:boxPr` and all false border hide flags.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips box explicit visible borders"` failed because parsed JSON dropped `hideTop: false`, `hideBottom: false`, `hideLeft: false`, and `hideRight: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes box explicit visible borders"` passed after writer emitted explicit `m:val="0"` flags.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips box explicit visible borders"` passed after reader used optional boolean parsing for box hide flags.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "box explicit visible borders"` passed 2 tests with 308 skipped.
- Full suite: `npm test` passed 310 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
