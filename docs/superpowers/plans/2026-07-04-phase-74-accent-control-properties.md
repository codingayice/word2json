# Phase 74 Accent Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:accPr/m:ctrlPr` support for `accent` math nodes so accent object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `accent` node with `controlProperties` and keep existing accent mark serialization unchanged.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `accent`.
- Modify `src/docx-writer.ts`: emit `accPr/ctrlPr` when `accent.controlProperties` is present.
- Modify `src/docx-reader.ts`: parse `accPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes accent control properties` that builds an accent node with `mark`, `content`, and `controlProperties`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes accent control properties"
```

Expected: FAIL because `accent.controlProperties` is ignored and no `m:accPr/m:ctrlPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `accent` node:

```ts
  | { type: "accent"; controlProperties?: MathControlProperties; mark: string; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Build `m:accPr` from existing `m:chr` plus optional `m:ctrlPr`:

```ts
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:acc><m:accPr><m:chr m:val="${escapeAttribute(node.mark)}"/>${properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : ""}</m:accPr><m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:acc>`;
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes accent control properties"
```

Expected: PASS with the expected `m:accPr/m:ctrlPr/w:rPr` XML.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips accent control properties` that builds, parses, and compares an accent node with `controlProperties`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips accent control properties"
```

Expected: FAIL because `parseMathNodes` ignores `accPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `accent` reader branch:

```ts
        const accentProperties = asObject(accentNode.accPr);
        const controlProperties = parseMathControlProperties(asObject(accentProperties.ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips accent control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-74-accent-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "accent control properties"
```

Expected: PASS for both Phase 74 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-74-accent-control-properties.md
git commit -m "feat: add phase 74 accent control properties"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-74-accent-control-properties
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-74-accent-control-properties.md
git commit -m "docs: mark phase 74 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 74 covers `m:accPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes accent control properties"` failed because `m:accPr/m:ctrlPr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes accent control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips accent control properties"` failed because parsed JSON did not include `controlProperties`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips accent control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "accent control properties"` passed 2 tests, 238 skipped.
- Full suite: `npm test` passed 240 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-74-accent-control-properties`
- Implementation commit: `4ace668 feat: add phase 74 accent control properties`
- Remote: `origin/phase-74-accent-control-properties`
