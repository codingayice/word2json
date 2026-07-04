# Phase 131 Text Run Snap To Grid Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run snap-to-grid formatting, including explicit `snapToGrid: false` to disable inherited grid snapping.

**Architecture:** Add `snapToGrid?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:snapToGrid/>` for true, `<w:snapToGrid w:val="0"/>` for false, and omits it when undefined; reader maps `w:snapToGrid` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `snapToGrid?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:snapToGrid` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:snapToGrid` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.snapToGrid: true` and `TextRun.snapToGrid: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run snap to grid on` and `writes text run snap to grid off`.

```ts
  it("writes text run snap to grid on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Grid", snapToGrid: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:snapToGrid/></w:rPr><w:t>Grid</w:t>");
  });

  it("writes text run snap to grid off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Free", snapToGrid: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:snapToGrid w:val="0"/></w:rPr><w:t>Free</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run snap to grid"
```

Expected: FAIL because `TextRun` does not accept `snapToGrid` and the writer does not emit `w:snapToGrid`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
snapToGrid?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.snapToGrid !== undefined ? (run.snapToGrid ? "<w:snapToGrid/>" : '<w:snapToGrid w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run snap to grid"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run snap to grid on` and `round-trips text run snap to grid off`.

```ts
  it("round-trips text run snap to grid on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Grid", snapToGrid: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run snap to grid off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Free", snapToGrid: false }],
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
npm test -- tests/docx-core.test.ts -t "round-trips text run snap to grid"
```

Expected: FAIL because the reader currently ignores `w:snapToGrid`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.snapToGrid, "snapToGrid"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run snap to grid"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-131-text-run-snap-to-grid.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run snap to grid"
```

Expected: PASS for all Phase 131 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-131-text-run-snap-to-grid.md
git commit -m "feat: add phase 131 text run snap to grid"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-131-text-run-snap-to-grid
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-131-text-run-snap-to-grid.md
git commit -m "docs: mark phase 131 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 131 covers schema, writer XML, and reader round-trip preservation for normal text run snap-to-grid on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `snapToGrid?: boolean` maps to WordprocessingML `w:snapToGrid`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run snap to grid"` failed because generated XML omitted `<w:snapToGrid/>` and `<w:snapToGrid w:val="0"/>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run snap to grid"` failed because parsed JSON dropped both `snapToGrid: true` and `snapToGrid: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run snap to grid"` passed 2 tests after writer emitted `<w:snapToGrid/>` and `<w:snapToGrid w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run snap to grid"` passed 2 tests after reader parsed `w:snapToGrid`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run snap to grid"` passed 4 tests with 374 skipped.
- Full suite: `npm test` passed 378 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
