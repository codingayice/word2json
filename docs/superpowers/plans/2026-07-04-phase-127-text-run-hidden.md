# Phase 127 Text Run Hidden Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run hidden-text formatting, including explicit `hidden: false` to disable inherited hidden formatting.

**Architecture:** Add `hidden?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:vanish/>` for true, `<w:vanish w:val="0"/>` for false, and omits the element when undefined; reader maps `w:vanish` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `hidden?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:vanish` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:vanish` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.hidden: true` and `TextRun.hidden: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run hidden on` and `writes text run hidden off`.

```ts
  it("writes text run hidden on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", hidden: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t>");
  });

  it("writes text run hidden off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", hidden: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:vanish w:val="0"/></w:rPr><w:t>Visible</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run hidden"
```

Expected: FAIL because `TextRun` does not accept `hidden` and the writer does not emit `w:vanish`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
hidden?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.hidden !== undefined ? (run.hidden ? "<w:vanish/>" : '<w:vanish w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run hidden"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run hidden on` and `round-trips text run hidden off`.

```ts
  it("round-trips text run hidden on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", hidden: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run hidden off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", hidden: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run hidden"
```

Expected: FAIL because the reader currently ignores `w:vanish`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.vanish, "hidden"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run hidden"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-127-text-run-hidden.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run hidden"
```

Expected: PASS for all Phase 127 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-127-text-run-hidden.md
git commit -m "feat: add phase 127 text run hidden"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-127-text-run-hidden
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-127-text-run-hidden.md
git commit -m "docs: mark phase 127 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 127 covers schema, writer XML, and reader round-trip preservation for normal text run hidden text on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `hidden?: boolean` maps to WordprocessingML `w:vanish`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run hidden"` failed because generated XML omitted `<w:vanish/>` and `<w:vanish w:val="0"/>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run hidden"` failed because parsed JSON dropped both `hidden: true` and `hidden: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run hidden"` passed 2 tests after writer emitted `<w:vanish/>` and `<w:vanish w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run hidden"` passed 2 tests after reader parsed `w:vanish` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run hidden"` passed 4 tests with 358 skipped.
- Full suite: `npm test` passed 362 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-127-text-run-hidden`
- Implementation commit: `2fec269 feat: add phase 127 text run hidden`
- Remote: `origin/phase-127-text-run-hidden`
