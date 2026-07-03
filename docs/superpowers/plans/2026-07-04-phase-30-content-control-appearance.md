# Phase 30 Content Control Appearance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word content control visual and placeholder-state metadata for higher fidelity template round-tripping.

**Architecture:** Extend the shared `ContentControl` model used by block and run SDTs. The writer emits `w:appearance`, `w:color`, and `w:showingPlcHdr` inside `w:sdtPr`; the reader parses the same nodes through `parseContentControl()` so both block and run content controls round-trip.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `appearance`, `color`, and `showingPlaceholder` fields to `ContentControl`.
- `src/docx-writer.ts`: emit content control visual metadata inside `w:sdtPr`.
- `src/docx-reader.ts`: parse content control visual metadata into JSON.
- `tests/docx-core.test.ts`: add writer and round-trip tests for appearance metadata.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes content control appearance metadata`:

```ts
it("writes content control appearance metadata", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Click or tap here",
          contentControl: {
            alias: "Prompt",
            tag: "prompt",
            appearance: "tags",
            color: "2F5496",
            showingPlaceholder: true,
          },
        },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:appearance w:val="tags"/>');
  expect(xml).toContain('<w:color w:val="2F5496"/>');
  expect(xml).toContain('<w:showingPlcHdr/>');
});
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control appearance metadata"`

Expected: FAIL because these content control metadata fields are not supported yet.

- [x] **Step 3: Implement minimal writer support**

Add fields to `ContentControl`:

```ts
appearance?: "boundingBox" | "tags" | "hidden";
color?: string;
showingPlaceholder?: boolean;
```

Emit:

```xml
<w:appearance w:val="tags"/>
<w:color w:val="2F5496"/>
<w:showingPlcHdr/>
```

only when the corresponding JSON values are present.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control appearance metadata"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips content control appearance metadata`:

```ts
it("round-trips content control appearance metadata", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Click or tap here",
          contentControl: {
            alias: "Prompt",
            tag: "prompt",
            appearance: "tags",
            color: "2F5496",
            showingPlaceholder: true,
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

Run: `npm test -- tests/docx-core.test.ts -t "content control appearance metadata"`

Expected: FAIL because the reader drops appearance metadata.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
appearance: properties.appearance.val
color: properties.color.val
showingPlaceholder: true
```

when the relevant Word nodes are present and values are strings.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "content control appearance metadata"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 30 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-30-content-control-appearance.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 30 content control appearance"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-30-content-control-appearance
```

Expected: branch `phase-30-content-control-appearance` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers content control appearance, color, and placeholder showing state.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `showingPlaceholder` while writer/reader map it to Word XML `w:showingPlcHdr`.
