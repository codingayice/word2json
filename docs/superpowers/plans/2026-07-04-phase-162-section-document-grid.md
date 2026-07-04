# Phase 162 Section Document Grid Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level document grid settings.

**Architecture:** Extend `SectionNode` with a `documentGrid` object, serialize it as `<w:docGrid>` inside section properties, and parse it back from `document.xml`. The implementation follows the existing section page-numbering, line-numbering, note-property, and columns patterns.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Document Grid Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section line numbering and note property writer tests:

```ts
it("writes section document grid settings", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        documentGrid: {
          type: "linesAndChars" as const,
          linePitch: 360,
          charSpace: 180,
        },
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Grid" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="180"/>');
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section line numbering and note property round-trip tests:

```ts
it("round-trips section document grid settings", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        documentGrid: {
          type: "linesAndChars" as const,
          linePitch: 360,
          charSpace: 180,
        },
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Grid" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section document grid settings"`

Expected: the writer and round-trip tests fail because section document grid settings are not serialized or parsed yet.

### Task 2: Implement Section Document Grid Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
documentGrid?: SectionDocumentGrid;
```

- [x] **Step 2: Add the grid type**

Add near the existing section types:

```ts
export type SectionDocumentGrid = {
  type?: "default" | "lines" | "linesAndChars" | "snapToChars";
  linePitch?: number;
  charSpace?: number;
};
```

### Task 3: Serialize Section Document Grid

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit document grid from section properties**

In `sectionPropertiesXml`, compute:

```ts
const documentGrid = documentGridXml(section.documentGrid);
```

Append it after columns so it is written within `<w:sectPr>`.

- [x] **Step 2: Add document grid XML helper**

Add:

```ts
function documentGridXml(documentGrid?: SectionNode["documentGrid"]): string {
  if (!documentGrid) {
    return "";
  }

  const attributes = [
    documentGrid.type ? ` w:type="${documentGrid.type}"` : "",
    documentGrid.linePitch !== undefined ? ` w:linePitch="${documentGrid.linePitch}"` : "",
    documentGrid.charSpace !== undefined ? ` w:charSpace="${documentGrid.charSpace}"` : "",
  ].join("");

  return attributes ? `<w:docGrid${attributes}/>` : "";
}
```

### Task 4: Parse Section Document Grid

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse document grid in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const documentGrid = parseDocumentGrid(sectPrValue);
```

Include it in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseLineNumbering` and `parseNoteProperties`:

```ts
function parseDocumentGrid(sectionPropertiesValue: unknown): SectionNode["documentGrid"] | undefined {
  const documentGrid = asObject(asObject(sectionPropertiesValue).docGrid);
  const parsed = {
    ...(typeof documentGrid.type === "string" ? { type: documentGrid.type as NonNullable<SectionNode["documentGrid"]>["type"] } : {}),
    ...(documentGrid.linePitch !== undefined ? { linePitch: parseNumber(documentGrid.linePitch) } : {}),
    ...(documentGrid.charSpace !== undefined ? { charSpace: parseNumber(documentGrid.charSpace) } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-162-section-document-grid.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section document grid settings"`

Expected: both tests pass.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: tests and build exit 0; `git diff --check` exits 0. CRLF warnings are acceptable only when the command still exits 0.

- [x] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-162-section-document-grid.md
git commit -m "feat: add phase 162 section document grid"
```

- [x] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-162-section-document-grid
```

- [x] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-162-section-document-grid
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-162-section-document-grid.md
git commit -m "docs: mark phase 162 pushed"
git push
```

## Push Record

- Branch: `phase-162-section-document-grid`
- Feature commit: `fa53cde feat: add phase 162 section document grid`
- Verification:
  - `npm test -- tests/docx-core.test.ts -t "section document grid settings"`: 2 passed
  - `npm test`: 466 passed
  - `npm run build`: exit 0
  - `git diff --check`: exit 0 with LF/CRLF warnings
- PR: https://github.com/codingayice/word2json/pull/new/phase-162-section-document-grid
