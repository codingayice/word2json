# Phase 89 Delimiter Grow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:dPr/m:grow` support for `delimiter` math nodes so delimiter growth behavior survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `delimiter` node with optional `grow?: boolean`. Reuse the current `m:dPr` container so `grow` can coexist with `begin`, `end`, and `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `grow?: boolean` on `delimiter`.
- Modify `src/docx-writer.ts`: emit `m:dPr/m:grow` when `grow` is present.
- Modify `src/docx-reader.ts`: parse `dPr.grow` into `grow`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes delimiter grow` that builds a delimiter with `grow: false` and expects `m:grow m:val="0"`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter grow"
```

Expected: FAIL because writer does not emit `m:grow`.

- [x] **Step 3: Extend schema**

Update the `delimiter` node:

```ts
  | { type: "delimiter"; controlProperties?: MathControlProperties; begin?: string; end?: string; grow?: boolean; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Emit this property in `m:dPr` when present:

```ts
node.grow !== undefined ? `<m:grow m:val="${node.grow ? "1" : "0"}"/>` : ""
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter grow"
```

Expected: PASS with `m:grow m:val="0"` while preserving begin, end, and content.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips delimiter grow` that builds, parses, and compares a delimiter with `grow: false`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter grow"
```

Expected: FAIL because `parseMathNodes` ignores `dPr.grow`.

- [x] **Step 3: Implement reader support**

Reuse `mathOptionalBooleanProperty`:

```ts
          ...mathOptionalBooleanProperty(delimiterProperties.grow, "grow"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter grow"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-89-delimiter-grow.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "delimiter grow"
```

Expected: PASS for both Phase 89 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-89-delimiter-grow.md
git commit -m "feat: add phase 89 delimiter grow"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-89-delimiter-grow
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-89-delimiter-grow.md
git commit -m "docs: mark phase 89 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 89 covers writer and reader round-trip for delimiter growth while preserving existing begin, end, content, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `grow` and maps to OMML `m:dPr/m:grow`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter grow"` failed because `m:dPr/m:grow` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter grow"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter grow"` failed because parsed JSON did not include `grow`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter grow"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "delimiter grow"` passed 2 tests, 268 skipped.
- Full suite: `npm test` passed 270 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-89-delimiter-grow`
- Implementation commit: `4bb28f8 feat: add phase 89 delimiter grow`
- Remote: `origin/phase-89-delimiter-grow`
