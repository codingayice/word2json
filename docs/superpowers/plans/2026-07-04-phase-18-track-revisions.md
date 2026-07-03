# Phase 18 Track Revisions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common tracked-change metadata and run-level insert/delete revisions in reviewed Word documents.

**Architecture:** Add document-level track revision setting and run-level revision metadata. The writer wraps affected runs in `w:ins` or `w:del`; the reader unwraps those revision containers into normal `TextRun` objects carrying revision metadata.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `trackRevisions` to `DocumentSettings` and `revision` to `TextRun`.
- `src/docx-writer.ts`: write `w:trackRevisions`, `w:ins`, `w:del`, and `w:delText`.
- `src/docx-reader.ts`: parse `w:trackRevisions`, `w:ins`, `w:del`, and deleted text.
- `tests/docx-core.test.ts`: add XML writer and round-trip tests.

## Task 1: Track Revisions Setting

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes track revisions setting", async () => {
  const document = {
    version: "1.0" as const,
    settings: { trackRevisions: true },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Tracked" }] }] }],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const settings = await zip.file("word/settings.xml")!.async("string");

  expect(settings).toContain("<w:trackRevisions/>");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips track revisions setting", async () => {
  const source = {
    version: "1.0" as const,
    settings: { trackRevisions: true },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Tracked" }] }] }],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "track revisions setting"`

Expected: FAIL because `trackRevisions` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `trackRevisions?: boolean` to `DocumentSettings`, emit `<w:trackRevisions/>`, and parse it from `word/settings.xml`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "track revisions setting"`

Expected: PASS.

## Task 2: Inserted Run Revisions

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes inserted run revisions", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Added", revision: { type: "insert", id: 1, author: "Ada", date: "2026-07-04T00:00:00.000Z" } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:ins w:id="1" w:author="Ada" w:date="2026-07-04T00:00:00.000Z">');
  expect(xml).toContain("<w:t>Added</w:t>");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips inserted run revisions", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Added", revision: { type: "insert", id: 1, author: "Ada", date: "2026-07-04T00:00:00.000Z" } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "inserted run revisions"`

Expected: FAIL because inserted revisions are not supported yet.

- [x] **Step 4: Implement minimal support**

Add `RunRevision` with insert metadata to `TextRun`. Wrap normal run XML in `<w:ins>`, and parse `w:ins/w:r` back into a run with revision metadata.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "inserted run revisions"`

Expected: PASS.

## Task 3: Deleted Run Revisions

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes deleted run revisions", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Removed", revision: { type: "delete", id: 2, author: "Lin", date: "2026-07-04T01:00:00.000Z" } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:del w:id="2" w:author="Lin" w:date="2026-07-04T01:00:00.000Z">');
  expect(xml).toContain("<w:delText>Removed</w:delText>");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips deleted run revisions", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Removed", revision: { type: "delete", id: 2, author: "Lin", date: "2026-07-04T01:00:00.000Z" } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "deleted run revisions"`

Expected: FAIL because deleted revisions and `w:delText` are not supported yet.

- [x] **Step 4: Implement minimal support**

Wrap deleted runs in `<w:del>` and write text as `w:delText`. Parse `w:del/w:r/w:delText` back into `TextRun.text` with delete revision metadata.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "deleted run revisions"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 18 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 18 track revisions"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-18-track-revisions
```

Expected: branch `phase-18-track-revisions` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers document revision tracking setting and run-level insert/delete revisions.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `trackRevisions` lives on settings; run revision metadata lives on `TextRun.revision`.
