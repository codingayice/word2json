# Phase 73 Delimiter Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:dPr/m:ctrlPr` support for `delimiter` math nodes so delimiter object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `delimiter` node with `controlProperties` and keep existing `begin`/`end` delimiter character serialization unchanged.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `delimiter`.
- Modify `src/docx-writer.ts`: emit `dPr/ctrlPr` when `delimiter.controlProperties` is present.
- Modify `src/docx-reader.ts`: parse `dPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes delimiter control properties` that builds a delimiter node with `begin`, `end`, `content`, and `controlProperties`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter control properties"
```

Expected: FAIL because `delimiter.controlProperties` is ignored and no `m:dPr/m:ctrlPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `delimiter` node:

```ts
  | { type: "delimiter"; controlProperties?: MathControlProperties; begin?: string; end?: string; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Build `m:dPr` from optional `m:begChr`, `m:endChr`, and `m:ctrlPr`:

```ts
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const delimiterProperties = [
      node.begin !== undefined ? `<m:begChr m:val="${escapeAttribute(node.begin)}"/>` : "",
      node.end !== undefined ? `<m:endChr m:val="${escapeAttribute(node.end)}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes delimiter control properties"
```

Expected: PASS with the expected `m:dPr/m:ctrlPr/w:rPr` XML.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips delimiter control properties` that builds, parses, and compares a delimiter node with `controlProperties`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter control properties"
```

Expected: FAIL because `parseMathNodes` ignores `dPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `delimiter` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(delimiterProperties.ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips delimiter control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-73-delimiter-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "delimiter control properties"
```

Expected: PASS for both Phase 73 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-73-delimiter-control-properties.md
git commit -m "feat: add phase 73 delimiter control properties"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-73-delimiter-control-properties
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-73-delimiter-control-properties.md
git commit -m "docs: mark phase 73 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 73 covers `m:dPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter control properties"` failed because `m:dPr/m:ctrlPr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes delimiter control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter control properties"` failed because parsed JSON did not include `controlProperties`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "delimiter control properties"` passed 2 tests, 236 skipped.
- Full suite: `npm test` passed 238 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
