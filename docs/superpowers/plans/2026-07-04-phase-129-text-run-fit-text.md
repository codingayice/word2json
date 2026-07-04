# Phase 129 Text Run Fit Text Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run fit-text formatting, including required fit width and optional fit-text id.

**Architecture:** Add `fitText?: { width: number; id?: number }` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:fitText w:val="..." w:id="..."/>` when present, omitting `w:id` when absent; reader maps `w:fitText` attributes back into the same object shape.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add a reusable `RunFitText` type and `fitText?: RunFitText` to run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:fitText` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:fitText` values with number semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.fitText` with and without `id`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run fit text with id` and `writes text run fit text without id`.

```ts
  it("writes text run fit text with id", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 1440, id: 7 } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:fitText w:val="1440" w:id="7"/></w:rPr><w:t>Fit</w:t>');
  });

  it("writes text run fit text without id", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 720 } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:fitText w:val="720"/></w:rPr><w:t>Fit</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run fit text"
```

Expected: FAIL because `TextRun` does not accept `fitText` and the writer does not emit `w:fitText`.

- [x] **Step 3: Implement writer support**

Add to `src/schema.ts`:

```ts
export type RunFitText = {
  width: number;
  id?: number;
};
```

Add to `StyleRunProperties`:

```ts
fitText?: RunFitText;
```

Add to `runPropertiesXml`:

```ts
run.fitText ? `<w:fitText w:val="${run.fitText.width}"${run.fitText.id !== undefined ? ` w:id="${run.fitText.id}"` : ""}/>` : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run fit text"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run fit text with id` and `round-trips text run fit text without id`.

```ts
  it("round-trips text run fit text with id", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 1440, id: 7 } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run fit text without id", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 720 } }],
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
npm test -- tests/docx-core.test.ts -t "round-trips text run fit text"
```

Expected: FAIL because the reader currently ignores `w:fitText`.

- [x] **Step 3: Implement reader support**

Read `properties.fitText` inside `parseRunFont`:

```ts
const fitText = asObject(properties.fitText);
```

Add to the returned run properties:

```ts
...(fitText.val !== undefined
  ? { fitText: { width: parseNumber(fitText.val), ...(fitText.id !== undefined ? { id: parseNumber(fitText.id) } : {}) } }
  : {}),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run fit text"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-129-text-run-fit-text.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run fit text"
```

Expected: PASS for all Phase 129 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-129-text-run-fit-text.md
git commit -m "feat: add phase 129 text run fit text"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-129-text-run-fit-text
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-129-text-run-fit-text.md
git commit -m "docs: mark phase 129 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 129 covers schema, writer XML, and reader round-trip preservation for normal text run fit-text with and without id.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `fitText?: RunFitText` maps to WordprocessingML `w:fitText`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run fit text"` failed because generated XML omitted `<w:fitText>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run fit text"` failed because parsed JSON dropped both `fitText: { width: 1440, id: 7 }` and `fitText: { width: 720 }`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run fit text"` passed 2 tests after writer emitted `<w:fitText>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run fit text"` passed 2 tests after reader parsed `w:fitText` attributes.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run fit text"` passed 4 tests with 366 skipped.
- Full suite: `npm test` passed 370 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-129-text-run-fit-text`
- Implementation commit: `d60990f feat: add phase 129 text run fit text`
- Remote: `origin/phase-129-text-run-fit-text`
