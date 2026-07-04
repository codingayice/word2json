# Phase 161 Section Note Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level footnote and endnote properties.

**Architecture:** Extend `SectionNode` with note property objects, serialize them into `<w:footnotePr>` and `<w:endnotePr>` inside section properties, and parse those elements back from `document.xml`. The implementation follows the existing section page-numbering and line-numbering patterns.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Note Property Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section page-numbering and line-numbering writer tests:

```ts
it("writes section footnote and endnote properties", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        footnoteProperties: {
          position: "beneathText" as const,
          numbering: { format: "lowerRoman" as const, start: 2, restart: "eachSect" as const },
        },
        endnoteProperties: {
          position: "sectEnd" as const,
          numbering: { format: "upperRoman" as const, start: 4, restart: "continuous" as const },
        },
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Notes" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachSect"/></w:footnotePr>');
  expect(xml).toContain('<w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="upperRoman"/><w:numStart w:val="4"/><w:numRestart w:val="continuous"/></w:endnotePr>');
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section page-numbering and line-numbering round-trip tests:

```ts
it("round-trips section footnote and endnote properties", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        footnoteProperties: {
          position: "beneathText" as const,
          numbering: { format: "lowerRoman" as const, start: 2, restart: "eachSect" as const },
        },
        endnoteProperties: {
          position: "sectEnd" as const,
          numbering: { format: "upperRoman" as const, start: 4, restart: "continuous" as const },
        },
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Notes" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section footnote and endnote properties"`

Expected: the writer and round-trip tests fail because section note properties are not serialized or parsed yet.

### Task 2: Implement Section Note Property Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add section note property fields**

Add to `SectionNode`:

```ts
footnoteProperties?: SectionNoteProperties;
endnoteProperties?: SectionNoteProperties;
```

- [x] **Step 2: Add the shared type**

Add near the existing section types:

```ts
export type SectionNoteProperties = {
  position?: "pageBottom" | "beneathText" | "sectEnd" | "docEnd";
  numbering?: {
    format?: NumberingFormat;
    start?: number;
    restart?: "continuous" | "eachSect" | "eachPage";
  };
};
```

### Task 3: Serialize Section Note Properties

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit note properties from section properties**

In `sectionPropertiesXml`, compute:

```ts
const footnoteProperties = notePropertiesXml("footnotePr", section.footnoteProperties);
const endnoteProperties = notePropertiesXml("endnotePr", section.endnoteProperties);
```

Append them after line numbering and before columns.

- [x] **Step 2: Add note property XML helper**

Add:

```ts
function notePropertiesXml(root: "footnotePr" | "endnotePr", properties?: SectionNode["footnoteProperties"]): string {
  if (!properties) {
    return "";
  }

  const children = [
    properties.position ? `<w:pos w:val="${properties.position}"/>` : "",
    properties.numbering?.format ? `<w:numFmt w:val="${properties.numbering.format}"/>` : "",
    properties.numbering?.start !== undefined ? `<w:numStart w:val="${properties.numbering.start}"/>` : "",
    properties.numbering?.restart ? `<w:numRestart w:val="${properties.numbering.restart}"/>` : "",
  ].join("");

  return children ? `<w:${root}>${children}</w:${root}>` : "";
}
```

### Task 4: Parse Section Note Properties

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse note properties in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const footnoteProperties = parseNoteProperties(sectPrValue, "footnotePr");
const endnoteProperties = parseNoteProperties(sectPrValue, "endnotePr");
```

Include them in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parsePageNumbering` and `parseLineNumbering`:

```ts
function parseNoteProperties(sectionPropertiesValue: unknown, key: "footnotePr" | "endnotePr"): SectionNode["footnoteProperties"] | undefined {
  const properties = asObject(asObject(sectionPropertiesValue)[key]);
  const position = asObject(properties.pos);
  const format = asObject(properties.numFmt);
  const start = asObject(properties.numStart);
  const restart = asObject(properties.numRestart);
  const numbering = {
    ...(typeof format.val === "string" ? { format: format.val as NonNullable<NonNullable<SectionNode["footnoteProperties"]>["numbering"]>["format"] } : {}),
    ...(start.val !== undefined ? { start: parseNumber(start.val) } : {}),
    ...(typeof restart.val === "string" ? { restart: restart.val as NonNullable<NonNullable<SectionNode["footnoteProperties"]>["numbering"]>["restart"] } : {}),
  };
  const parsed = {
    ...(typeof position.val === "string" ? { position: position.val as NonNullable<SectionNode["footnoteProperties"]>["position"] } : {}),
    ...(Object.keys(numbering).length > 0 ? { numbering } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-161-section-note-properties.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section footnote and endnote properties"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-161-section-note-properties.md
git commit -m "feat: add phase 161 section note properties"
```

- [x] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-161-section-note-properties
```

- [x] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-161-section-note-properties
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-161-section-note-properties.md
git commit -m "docs: mark phase 161 pushed"
git push
```

## Push Record

- Branch: `phase-161-section-note-properties`
- Feature commit: `9d18f8f feat: add phase 161 section note properties`
- Verification:
  - `npm test -- tests/docx-core.test.ts -t "section footnote and endnote properties"`: 2 passed
  - `npm test`: 464 passed
  - `npm run build`: exit 0
  - `git diff --check`: exit 0 with LF/CRLF warnings
- PR: https://github.com/codingayice/word2json/pull/new/phase-161-section-note-properties
