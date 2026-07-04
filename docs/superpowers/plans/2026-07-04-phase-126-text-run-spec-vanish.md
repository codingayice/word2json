# Phase 126 Text Run Spec Vanish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run special vanish formatting, including explicit `specVanish: false` to disable inherited special vanish formatting.

**Architecture:** Add `specVanish?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:specVanish/>` for true, `<w:specVanish w:val="0"/>` for false, and omits the element when undefined; reader maps `w:specVanish` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `specVanish?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:specVanish` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:specVanish` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.specVanish: true` and `TextRun.specVanish: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run spec vanish on` and `writes text run spec vanish off`.

```ts
  it("writes text run spec vanish on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", specVanish: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:specVanish/></w:rPr><w:t>Hidden</w:t>");
  });

  it("writes text run spec vanish off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", specVanish: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:specVanish w:val="0"/></w:rPr><w:t>Visible</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run spec vanish"
```

Expected: FAIL because `TextRun` does not accept `specVanish` and the writer does not emit `w:specVanish`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
specVanish?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.specVanish !== undefined ? (run.specVanish ? "<w:specVanish/>" : '<w:specVanish w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run spec vanish"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run spec vanish on` and `round-trips text run spec vanish off`.

```ts
  it("round-trips text run spec vanish on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", specVanish: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run spec vanish off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", specVanish: false }],
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
npm test -- tests/docx-core.test.ts -t "round-trips text run spec vanish"
```

Expected: FAIL because the reader currently ignores `w:specVanish`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.specVanish, "specVanish"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run spec vanish"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-126-text-run-spec-vanish.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run spec vanish"
```

Expected: PASS for all Phase 126 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-126-text-run-spec-vanish.md
git commit -m "feat: add phase 126 text run spec vanish"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-126-text-run-spec-vanish
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-126-text-run-spec-vanish.md
git commit -m "docs: mark phase 126 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 126 covers schema, writer XML, and reader round-trip preservation for normal text run special vanish on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `specVanish?: boolean` maps to WordprocessingML `w:specVanish`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run spec vanish"` failed because generated XML omitted `<w:specVanish/>` and `<w:specVanish w:val="0"/>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run spec vanish"` failed because parsed JSON dropped both `specVanish: true` and `specVanish: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run spec vanish"` passed 2 tests after writer emitted `<w:specVanish/>` and `<w:specVanish w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run spec vanish"` passed 2 tests after reader parsed `w:specVanish` on/off values.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run spec vanish"` passed 4 tests with 354 skipped.
- Full suite: `npm test` passed 358 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
