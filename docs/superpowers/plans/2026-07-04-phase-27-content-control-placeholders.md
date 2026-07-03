# Phase 27 Content Control Placeholders Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word content control placeholder docPart metadata during JSON to DOCX writing and DOCX to JSON round-tripping.

**Architecture:** Extend the existing `ContentControl` schema with an optional `placeholder` object that stores the Word `w:docPart` value. The writer emits `w:placeholder/w:docPart` inside `w:sdtPr`; the reader parses the same node from block and run SDTs through the shared `parseContentControl()` path.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `ContentControlPlaceholder` and `ContentControl.placeholder`.
- `src/docx-writer.ts`: emit `w:placeholder` metadata inside content control properties.
- `src/docx-reader.ts`: parse `w:placeholder/w:docPart` into JSON.
- `tests/docx-core.test.ts`: add writer and round-trip tests for placeholder metadata.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes content control placeholders`:

```ts
it("writes content control placeholders", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Click or tap here",
          contentControl: {
            alias: "Recipient",
            tag: "recipient",
            placeholder: { docPart: "DefaultPlaceholder_22610170" },
          },
        },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:placeholder><w:docPart w:val="DefaultPlaceholder_22610170"/></w:placeholder>');
});
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control placeholders"`

Expected: FAIL because `placeholder` is not part of the schema and writer yet.

- [x] **Step 3: Implement minimal writer support**

Add these schema fields:

```ts
export type ContentControl = {
  alias?: string;
  tag?: string;
  lock?: "sdtLocked" | "contentLocked" | "sdtContentLocked" | "unlocked";
  placeholder?: ContentControlPlaceholder;
  checkbox?: CheckboxContentControl;
  dropdown?: DropdownContentControl;
  date?: DateContentControl;
};

export type ContentControlPlaceholder = {
  docPart: string;
};
```

Add a helper in `src/docx-writer.ts`:

```ts
function placeholderContentControlXml(placeholder: NonNullable<NonNullable<ParagraphNode["contentControl"]>["placeholder"]>): string {
  return `<w:placeholder><w:docPart w:val="${escapeAttribute(placeholder.docPart)}"/></w:placeholder>`;
}
```

Include it in the `properties` array before form-specific metadata.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control placeholders"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips content control placeholders`:

```ts
it("round-trips content control placeholders", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Click or tap here",
          contentControl: {
            alias: "Recipient",
            tag: "recipient",
            placeholder: { docPart: "DefaultPlaceholder_22610170" },
          },
        },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "content control placeholders"`

Expected: FAIL because the reader drops placeholder metadata.

- [x] **Step 3: Implement minimal reader support**

Parse `properties.placeholder.docPart.val` and return:

```ts
placeholder: { docPart: docPart.val }
```

when the value is a string.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "content control placeholders"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 27 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-27-content-control-placeholders.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 27 content control placeholders"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-27-content-control-placeholders
```

Expected: branch `phase-27-content-control-placeholders` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers the placeholder docPart metadata Word stores on content controls.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `placeholder.docPart` is the same property name in schema, writer, reader, and tests.
