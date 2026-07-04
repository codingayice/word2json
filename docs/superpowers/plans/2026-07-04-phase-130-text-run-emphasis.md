# Phase 130 Text Run Emphasis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run emphasis mark formatting, including explicit `emphasis: "none"` to disable inherited emphasis marks.

**Architecture:** Add a `RunEmphasis` union and `emphasis?: RunEmphasis` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:em w:val="..."/>` when present; reader maps `w:em/@w:val` back into the same string value.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `RunEmphasis` type and `emphasis?: RunEmphasis` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:em` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:em` values.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.emphasis`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run emphasis dot` and `writes text run emphasis none`.

```ts
  it("writes text run emphasis dot", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emphasis", emphasis: "dot" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:em w:val="dot"/></w:rPr><w:t>Emphasis</w:t>');
  });

  it("writes text run emphasis none", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emphasis: "none" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:em w:val="none"/></w:rPr><w:t>Plain</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run emphasis"
```

Expected: FAIL because `TextRun` does not accept `emphasis` and the writer does not emit `w:em`.

- [x] **Step 3: Implement writer support**

Add to `src/schema.ts`:

```ts
export type RunEmphasis = "dot" | "comma" | "circle" | "underDot" | "none";
```

Add to `StyleRunProperties`:

```ts
emphasis?: RunEmphasis;
```

Add to `runPropertiesXml`:

```ts
run.emphasis ? `<w:em w:val="${run.emphasis}"/>` : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run emphasis"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run emphasis dot` and `round-trips text run emphasis none`.

```ts
  it("round-trips text run emphasis dot", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emphasis", emphasis: "dot" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run emphasis none", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emphasis: "none" }],
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
npm test -- tests/docx-core.test.ts -t "round-trips text run emphasis"
```

Expected: FAIL because the reader currently ignores `w:em`.

- [x] **Step 3: Implement reader support**

Read `properties.em` inside `parseRunFont`:

```ts
const emphasis = asObject(properties.em);
```

Add to the returned run properties:

```ts
...(typeof emphasis.val === "string" ? { emphasis: emphasis.val as NonNullable<TextRun["emphasis"]> } : {}),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run emphasis"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-130-text-run-emphasis.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run emphasis"
```

Expected: PASS for all Phase 130 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-130-text-run-emphasis.md
git commit -m "feat: add phase 130 text run emphasis"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-130-text-run-emphasis
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-130-text-run-emphasis.md
git commit -m "docs: mark phase 130 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 130 covers schema, writer XML, and reader round-trip preservation for normal text run emphasis marks.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `emphasis?: RunEmphasis` maps to WordprocessingML `w:em`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run emphasis"` failed because generated XML omitted `<w:em>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run emphasis"` failed because parsed JSON dropped both `emphasis: "dot"` and `emphasis: "none"`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run emphasis"` passed 2 tests after writer emitted `<w:em>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run emphasis"` passed 2 tests after reader parsed `w:em`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run emphasis"` passed 4 tests with 370 skipped.
- Full suite: `npm test` passed 374 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
