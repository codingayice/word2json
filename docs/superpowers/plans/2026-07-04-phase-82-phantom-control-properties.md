# Phase 82 Phantom Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:phantPr/m:ctrlPr` support for `phantom` math nodes so phantom object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `phantom` node with `controlProperties`, preserving existing phantom flags in the same `m:phantPr` container.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `phantom`.
- Modify `src/docx-writer.ts`: emit `phantPr/ctrlPr` alongside existing phantom flags.
- Modify `src/docx-reader.ts`: parse `phantPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes phantom control properties` that builds a phantom with flags and `controlProperties`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes phantom control properties"
```

Expected: FAIL because `phantom.controlProperties` is ignored and no `m:phantPr/m:ctrlPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `phantom` node:

```ts
  | { type: "phantom"; controlProperties?: MathControlProperties; show?: boolean; zeroWidth?: boolean; zeroAscent?: boolean; zeroDescent?: boolean; transparent?: boolean; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add optional control properties to the existing phantom property list:

```ts
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
```

and append:

```ts
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes phantom control properties"
```

Expected: PASS with the expected `m:phantPr/m:ctrlPr/w:rPr` XML while preserving phantom flags.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips phantom control properties` that builds, parses, and compares a phantom node with flags and `controlProperties`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips phantom control properties"
```

Expected: FAIL because `parseMathNodes` ignores `phantPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `phantom` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(phantomProperties.ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips phantom control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-82-phantom-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "phantom control properties"
```

Expected: PASS for both Phase 82 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-82-phantom-control-properties.md
git commit -m "feat: add phase 82 phantom control properties"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-82-phantom-control-properties
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-82-phantom-control-properties.md
git commit -m "docs: mark phase 82 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 82 covers `m:phantPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties while preserving existing phantom flags.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes phantom control properties"` failed because `m:phantPr/m:ctrlPr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes phantom control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips phantom control properties"` failed because parsed JSON did not include `controlProperties`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips phantom control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "phantom control properties"` passed 2 tests, 254 skipped.
- Full suite: `npm test` passed 256 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
