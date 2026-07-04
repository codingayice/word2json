# Phase 135 Text Run Character Position Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run character position metadata serialized as `<w:position>`.

**Architecture:** This phase adds `characterPosition?: number` to `StyleRunProperties`, preserving the raw WordprocessingML `w:val` integer so no unit conversion loses fidelity. The writer emits `<w:position w:val="..."/>`, and the reader parses the same value back through the existing numeric parser.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Character Position

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run character position raised", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Raised", characterPosition: 4 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:position w:val="4"/></w:rPr><w:t>Raised</w:t>');
  });

  it("writes text run character position lowered", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Lowered", characterPosition: -4 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:position w:val="-4"/></w:rPr><w:t>Lowered</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run character position raised", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Raised", characterPosition: 4 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run character position lowered", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Lowered", characterPosition: -4 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run character position"`

Expected: FAIL because `<w:position>` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run character position"`

Expected: FAIL because `characterPosition` is not preserved through parsing yet.

### Task 2: Implement Text Run Character Position

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the numeric field to `StyleRunProperties`:

```ts
  characterPosition?: number;
```

- [x] **Step 2: Emit writer XML**

Add the run property in `runPropertiesXml` near `verticalAlign` and `characterSpacing`:

```ts
    run.characterPosition !== undefined ? `<w:position w:val="${run.characterPosition}"/>` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
  const characterPosition = asObject(properties.position);
```

```ts
    ...(characterPosition.val !== undefined ? { characterPosition: parseNumber(characterPosition.val) } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run character position"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-135-text-run-character-position.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run character position"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-135-text-run-character-position.md
git commit -m "feat: add phase 135 text run character position"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-135-text-run-character-position
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-135-text-run-character-position.md
git commit -m "docs: mark phase 135 pushed"
git push
```

## Push Record

- Pending.
