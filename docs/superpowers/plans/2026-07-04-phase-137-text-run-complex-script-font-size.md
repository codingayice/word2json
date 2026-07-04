# Phase 137 Text Run Complex Script Font Size Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run complex script font size serialized as `<w:szCs>`.

**Architecture:** This phase adds `complexScriptFontSize?: number` to `StyleRunProperties`, mirroring the existing `fontSize` point-based JSON API while serializing to WordprocessingML half-points. The writer emits `<w:szCs w:val="..."/>`, and the reader parses the same value back by dividing by two.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Complex Script Font Size

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run complex script font size", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontSize: 14 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:szCs w:val="28"/></w:rPr><w:t>مرحبا</w:t>');
  });

  it("writes text run latin and complex script font sizes", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontSize: 12, complexScriptFontSize: 16 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:sz w:val="24"/><w:szCs w:val="32"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run complex script font size", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontSize: 14 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run latin and complex script font sizes", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontSize: 12, complexScriptFontSize: 16 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run complex script font size"`

Expected: FAIL because `<w:szCs>` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run complex script font size"`

Expected: FAIL because `complexScriptFontSize` is not preserved through parsing yet.

### Task 2: Implement Text Run Complex Script Font Size

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the numeric field to `StyleRunProperties`:

```ts
  complexScriptFontSize?: number;
```

- [x] **Step 2: Emit writer XML**

Add the run property in `runPropertiesXml` next to `fontSize`:

```ts
    run.complexScriptFontSize ? `<w:szCs w:val="${run.complexScriptFontSize * 2}"/>` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
  const complexScriptSize = asObject(properties.szCs);
```

```ts
    ...(typeof complexScriptSize.val === "number" ? { complexScriptFontSize: complexScriptSize.val / 2 } : {}),
    ...(typeof complexScriptSize.val === "string" ? { complexScriptFontSize: Number.parseInt(complexScriptSize.val, 10) / 2 } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run complex script font size"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-137-text-run-complex-script-font-size.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run complex script font size"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-137-text-run-complex-script-font-size.md
git commit -m "feat: add phase 137 text run complex script font size"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-137-text-run-complex-script-font-size
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-137-text-run-complex-script-font-size.md
git commit -m "docs: mark phase 137 pushed"
git push
```

## Push Record

- Pending.
