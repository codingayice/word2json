# Phase 136 Text Run Kerning Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run kerning metadata serialized as `<w:kern>`.

**Architecture:** This phase adds `kerning?: number` to `StyleRunProperties`, preserving WordprocessingML's raw `w:val` integer so no unit conversion loses fidelity. The writer emits `<w:kern w:val="..."/>`, and the reader parses the same value back through the existing numeric parser.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Kerning

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run kerning small threshold", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Title", kerning: 24 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:kern w:val="24"/></w:rPr><w:t>Title</w:t>');
  });

  it("writes text run kerning large threshold", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Display", kerning: 48 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:kern w:val="48"/></w:rPr><w:t>Display</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run kerning small threshold", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Title", kerning: 24 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run kerning large threshold", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Display", kerning: 48 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run kerning"`

Expected: FAIL because `<w:kern>` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run kerning"`

Expected: FAIL because `kerning` is not preserved through parsing yet.

### Task 2: Implement Text Run Kerning

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the numeric field to `StyleRunProperties`:

```ts
  kerning?: number;
```

- [x] **Step 2: Emit writer XML**

Add the run property in `runPropertiesXml` near other numeric typography properties:

```ts
    run.kerning !== undefined ? `<w:kern w:val="${run.kerning}"/>` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
  const kerning = asObject(properties.kern);
```

```ts
    ...(kerning.val !== undefined ? { kerning: parseNumber(kerning.val) } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run kerning"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-136-text-run-kerning.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run kerning"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-136-text-run-kerning.md
git commit -m "feat: add phase 136 text run kerning"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-136-text-run-kerning
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-136-text-run-kerning.md
git commit -m "docs: mark phase 136 pushed"
git push
```

## Push Record

- Branch: `phase-136-text-run-kerning`
- Remote: `origin/phase-136-text-run-kerning`
- Implementation commit: `3c94bf3 feat: add phase 136 text run kerning`
- Pushed to: `https://github.com/codingayice/word2json.git`
