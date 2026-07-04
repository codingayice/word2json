# Phase 90 Delimiter Separator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:dPr/m:sepChr` support for `delimiter` math nodes so separator characters survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `delimiter` node with optional `separator?: string`. Reuse the current `m:dPr` container beside `begin`, `end`, `grow`, and `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `separator?: string` on `delimiter`.
- Modify `src/docx-writer.ts`: emit `m:dPr/m:sepChr` when `separator` is present.
- Modify `src/docx-reader.ts`: parse `dPr.sepChr` into `separator`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes delimiter separator` that builds a delimiter with `separator: "|"`, `begin: "{"`, and `end: "}"`, then expects `m:sepChr m:val="|"`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter separator"
```

Expected: FAIL because writer does not emit `m:sepChr`.

- [x] **Step 3: Extend schema**

Update the `delimiter` node:

```ts
  | { type: "delimiter"; controlProperties?: MathControlProperties; begin?: string; end?: string; grow?: boolean; separator?: string; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Emit this property in `m:dPr` when present:

```ts
node.separator !== undefined ? `<m:sepChr m:val="${escapeAttribute(node.separator)}"/>` : ""
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter separator"
```

Expected: PASS with `m:sepChr m:val="|"` while preserving begin, end, and content.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips delimiter separator` that builds, parses, and compares a delimiter with `separator: "|"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter separator"
```

Expected: FAIL because `parseMathNodes` ignores `dPr.sepChr`.

- [x] **Step 3: Implement reader support**

Read `sepChr` beside `begChr` and `endChr`, then include it when it is a string:

```ts
const separator = asObject(delimiterProperties.sepChr).val;
```

```ts
...(typeof separator === "string" ? { separator } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter separator"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-90-delimiter-separator.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "delimiter separator"
```

Expected: PASS for both Phase 90 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-90-delimiter-separator.md
git commit -m "feat: add phase 90 delimiter separator"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-90-delimiter-separator
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-90-delimiter-separator.md
git commit -m "docs: mark phase 90 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 90 covers writer and reader round-trip for delimiter separators while preserving existing begin, end, grow, content, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `separator` and maps to OMML `m:dPr/m:sepChr`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter separator"` failed because `m:dPr/m:sepChr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter separator"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter separator"` failed because parsed JSON did not include `separator`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter separator"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "delimiter separator"` passed 2 tests, 270 skipped.
- Full suite: `npm test` passed 272 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
