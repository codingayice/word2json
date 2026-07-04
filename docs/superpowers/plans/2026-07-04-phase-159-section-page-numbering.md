# Phase 159 Section Page Numbering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve section page numbering settings through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Add `pageNumbering` to `SectionNode` and map it directly to `<w:pgNumType>`. Keep page size/margins parsing separate from page numbering so existing page layout behavior remains unchanged.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Section Page Numbering

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a writer test near section tests:

```ts
  it("writes section page numbering settings", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          pageNumbering: {
            start: 3,
            format: "lowerRoman" as const,
            chapterStyle: 1,
            chapterSeparator: "hyphen" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Preface" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pgNumType w:start="3" w:fmt="lowerRoman" w:chapStyle="1" w:chapSep="hyphen"/>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a reader round-trip test:

```ts
  it("round-trips section page numbering settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          pageNumbering: {
            start: 3,
            format: "lowerRoman" as const,
            chapterStyle: 1,
            chapterSeparator: "hyphen" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Preface" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "section page numbering settings"`

Expected: FAIL because `<w:pgNumType>` is not emitted or parsed yet.

### Task 2: Implement Section Page Numbering

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add to `SectionNode`:

```ts
  pageNumbering?: SectionPageNumbering;
```

Add:

```ts
export type SectionPageNumbering = {
  start?: number;
  format?: NumberingFormat;
  chapterStyle?: number;
  chapterSeparator?: "colon" | "emDash" | "enDash" | "hyphen" | "period";
};
```

- [x] **Step 2: Emit `<w:pgNumType>`**

Add a helper:

```ts
function pageNumberingXml(pageNumbering?: SectionNode["pageNumbering"]): string {
  if (!pageNumbering) return "";
  const attributes = [
    pageNumbering.start !== undefined ? ` w:start="${pageNumbering.start}"` : "",
    pageNumbering.format ? ` w:fmt="${pageNumbering.format}"` : "",
    pageNumbering.chapterStyle !== undefined ? ` w:chapStyle="${pageNumbering.chapterStyle}"` : "",
    pageNumbering.chapterSeparator ? ` w:chapSep="${pageNumbering.chapterSeparator}"` : "",
  ].join("");
  return attributes ? `<w:pgNumType${attributes}/>` : "";
}
```

Include it in `sectionPropertiesXml`.

- [x] **Step 3: Parse `<w:pgNumType>`**

Add `parsePageNumbering(sectionProperties)` and include its result in parsed sections.

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-159-section-page-numbering.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "section page numbering settings"`

Expected: PASS.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: all commands exit 0. Existing LF/CRLF warnings from `git diff --check` are acceptable if the exit code is 0.

- [x] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-159-section-page-numbering.md
git commit -m "feat: add phase 159 section page numbering"
git push -u origin phase-159-section-page-numbering
```

- [x] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-159-section-page-numbering
```

- [x] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-159-section-page-numbering.md
git commit -m "docs: mark phase 159 pushed"
git push
```

## Self-Review

- Spec coverage: Covers page numbering start, format, chapter style, chapter separator, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `pageNumbering`, `start`, `format`, `chapterStyle`, and `chapterSeparator` consistently across schema, writer, reader, and tests.

## Push Record

- Branch: `phase-159-section-page-numbering`
- Feature commit: `8f24b84 feat: add phase 159 section page numbering`
- PR URL: `https://github.com/codingayice/word2json/pull/new/phase-159-section-page-numbering`
- RED verification: `npm test -- tests/docx-core.test.ts -t "section page numbering settings"` failed with 2 expected failures before implementation.
- GREEN targeted verification: `npm test -- tests/docx-core.test.ts -t "section page numbering settings"` passed with 2 tests.
- Full verification: `npm test` passed with 460 tests.
- Build verification: `npm run build` exited 0.
- Whitespace verification: `git diff --check` exited 0 with existing LF/CRLF warnings.
