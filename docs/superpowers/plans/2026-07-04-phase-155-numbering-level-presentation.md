# Phase 155 Numbering Level Presentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve numbering level presentation metadata through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Extend the existing `NumberingLevelDefinition` type with fields that map to OOXML elements inside `<w:lvl>`. Reuse the current style run property serializer/parser for numbering symbol `<w:rPr>` instead of creating a duplicate run-format model.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Numbering Presentation

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a test near the numbering writer tests:

```ts
  it("writes numbering level style alignment and run properties", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 30,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1.",
                start: 1,
                styleId: "LegalClause",
                alignment: "right" as const,
                left: 720,
                hanging: 360,
                run: {
                  bold: true,
                  italic: false,
                  fontFamily: "Aptos",
                  fontSize: 11,
                  color: "C00000",
                },
              },
            ],
          },
        ],
        nums: [{ id: 30, abstractId: 30 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="1"/><w:pStyle w:val="LegalClause"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:b/><w:i w:val="0"/><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:color w:val="C00000"/></w:rPr></w:lvl>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a reader round-trip test:

```ts
  it("round-trips numbering level style alignment and run properties", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 30,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1.",
                start: 1,
                styleId: "LegalClause",
                alignment: "right" as const,
                left: 720,
                hanging: 360,
                run: {
                  bold: true,
                  italic: false,
                  fontFamily: "Aptos",
                  fontSize: 11,
                  color: "C00000",
                },
              },
            ],
          },
        ],
        nums: [{ id: 30, abstractId: 30 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering level style alignment and run properties"`

Expected: FAIL because these numbering level fields are not emitted or parsed yet.

### Task 2: Implement Numbering Presentation Fields

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend `NumberingLevelDefinition`**

Add:

```ts
  styleId?: string;
  alignment?: ParagraphAlignment;
  run?: StyleRunProperties;
```

- [x] **Step 2: Emit presentation fields**

In `numberingLevelXml`, create:

```ts
  const style = level.styleId ? `<w:pStyle w:val="${escapeAttribute(level.styleId)}"/>` : "";
  const alignment = level.alignment ? `<w:lvlJc w:val="${level.alignment}"/>` : "";
  const runProperties = styleRunPropertiesXml(level.run);
```

Then include `style` after `<w:start>`, `alignment` before paragraph properties, and `runProperties` after paragraph properties.

- [x] **Step 3: Parse presentation fields**

In `parseAbstractNumberingDefinition`, read:

```ts
      const style = asObject(level.pStyle);
      const alignment = asObject(level.lvlJc);
      const run = parseStyleRunProperties(level.rPr);
```

Then return:

```ts
        ...(typeof style.val === "string" ? { styleId: style.val } : {}),
        ...(typeof alignment.val === "string" ? { alignment: alignment.val as ParagraphAlignment } : {}),
        ...(run ? { run } : {}),
```

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-155-numbering-level-presentation.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering level style alignment and run properties"`

Expected: PASS.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: all commands exit 0. Existing LF/CRLF warnings from `git diff --check` are acceptable if the exit code is 0.

- [ ] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-155-numbering-level-presentation.md
git commit -m "feat: add phase 155 numbering level presentation"
git push -u origin phase-155-numbering-level-presentation
```

- [ ] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-155-numbering-level-presentation
```

- [ ] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-155-numbering-level-presentation.md
git commit -m "docs: mark phase 155 pushed"
git push
```

## Self-Review

- Spec coverage: Covers numbering level paragraph style binding, numbering alignment, numbering symbol run properties, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `styleId`, `alignment`, and `run` consistently across schema, writer, reader, and tests.
