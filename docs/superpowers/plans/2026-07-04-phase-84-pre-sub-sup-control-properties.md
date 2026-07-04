# Phase 84 Pre Sub Sup Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:preSubSupPr/m:ctrlPr` support for `preSubSup` math nodes so prescript object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `preSubSup` node with `controlProperties`, preserving existing base, subscript, and superscript children.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `preSubSup`.
- Modify `src/docx-writer.ts`: emit `preSubSupPr/ctrlPr` before the existing base, subscript, and superscript children.
- Modify `src/docx-reader.ts`: parse `preSubSupPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes pre sub sup control properties` that builds a `preSubSup` with `controlProperties`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes pre sub sup control properties"
```

Expected: FAIL because `preSubSup.controlProperties` is ignored and no `m:preSubSupPr/m:ctrlPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `preSubSup` node:

```ts
  | { type: "preSubSup"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Add optional control properties to the existing `preSubSup` writer branch:

```ts
    const properties = mathControlPropertiesXml(node.controlProperties);
```

and emit:

```ts
${properties ? `<m:preSubSupPr><m:ctrlPr>${properties}</m:ctrlPr></m:preSubSupPr>` : ""}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes pre sub sup control properties"
```

Expected: PASS with the expected `m:preSubSupPr/m:ctrlPr/w:rPr` XML while preserving base, subscript, and superscript.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips pre sub sup control properties` that builds, parses, and compares a `preSubSup` node with `controlProperties`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup control properties"
```

Expected: FAIL because `parseMathNodes` ignores `preSubSupPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `preSubSup` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(asObject(preSubSupNode.preSubSupPr).ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-84-pre-sub-sup-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "pre sub sup control properties"
```

Expected: PASS for both Phase 84 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-84-pre-sub-sup-control-properties.md
git commit -m "feat: add phase 84 pre sub sup control properties"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-84-pre-sub-sup-control-properties
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-84-pre-sub-sup-control-properties.md
git commit -m "docs: mark phase 84 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 84 covers `m:preSubSupPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties while preserving base, subscript, and superscript children.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes pre sub sup control properties"` failed because `m:preSubSupPr/m:ctrlPr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes pre sub sup control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup control properties"` failed because parsed JSON did not include `controlProperties`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "pre sub sup control properties"` passed 2 tests, 258 skipped.
- Full suite: `npm test` passed 260 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
