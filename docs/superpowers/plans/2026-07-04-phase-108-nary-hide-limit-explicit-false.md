# Phase 108 N-Ary Hide Limit Explicit False Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `false` values for n-ary lower and upper limit hide flags so OMML `m:subHide` and `m:supHide` round-trip without losing author intent.

**Architecture:** Keep the existing `hideLowerLimit?: boolean` and `hideUpperLimit?: boolean` schema fields. Writer emits `m:val="1"` or `m:val="0"` whenever a hide flag is defined; reader uses optional boolean parsing for n-ary hide flags so present `m:val="0"` becomes `false` instead of being dropped.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/docx-writer.ts`: write `m:subHide` and `m:supHide` for both true and false values.
- Modify `src/docx-reader.ts`: parse n-ary `subHide` and `supHide` with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one round-trip test for explicit false hide flags.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary explicit visible limits` that builds an n-ary node with `hideLowerLimit: false` and `hideUpperLimit: false`, then expects `m:subHide m:val="0"` and `m:supHide m:val="0"` inside `m:naryPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary explicit visible limits"
```

Expected: FAIL because the writer currently omits hide flags when their values are false.

- [x] **Step 3: Implement writer support**

Update the n-ary property list:

```ts
node.hideLowerLimit !== undefined ? `<m:subHide m:val="${node.hideLowerLimit ? "1" : "0"}"/>` : "",
node.hideUpperLimit !== undefined ? `<m:supHide m:val="${node.hideUpperLimit ? "1" : "0"}"/>` : "",
```

- [x] **Step 4: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary explicit visible limits"
```

Expected: PASS with both explicit false OMML flags present.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary explicit visible limits` that builds, parses, and compares an n-ary node with `hideLowerLimit: false` and `hideUpperLimit: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary explicit visible limits"
```

Expected: FAIL because `parseMathNodes` currently drops false hide flags by using true-only boolean parsing.

- [x] **Step 3: Implement reader support**

Change n-ary hide parsing to optional boolean parsing:

```ts
...mathOptionalBooleanProperty(naryProperties.subHide, "hideLowerLimit"),
...mathOptionalBooleanProperty(naryProperties.supHide, "hideUpperLimit"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary explicit visible limits"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-108-nary-hide-limit-explicit-false.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary explicit visible limits"
```

Expected: PASS for both Phase 108 tests.

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
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-108-nary-hide-limit-explicit-false.md
git commit -m "feat: add phase 108 nary explicit hide limits"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-108-nary-hide-limit-explicit-false
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-108-nary-hide-limit-explicit-false.md
git commit -m "docs: mark phase 108 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 108 covers writer and reader round-trip preservation for explicit false n-ary hide flags.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: Existing schema fields remain `hideLowerLimit` and `hideUpperLimit`, mapping to OMML `m:subHide` and `m:supHide`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary explicit visible limits"` failed because the writer omitted `m:subHide` and `m:supHide` for false values.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary explicit visible limits"` failed because parsed JSON dropped `hideLowerLimit: false` and `hideUpperLimit: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary explicit visible limits"` passed after writer emitted explicit `m:val="0"` flags.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary explicit visible limits"` passed after reader used optional boolean parsing for n-ary hide flags.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "nary explicit visible limits"` passed 2 tests with 306 skipped.
- Full suite: `npm test` passed 308 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
