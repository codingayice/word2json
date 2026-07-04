# Phase 158 Section Header Footer Types Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve first-page and even-page section header/footer references through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Extend section header/footer content from default-only to default/first/even variants, and add a section-level `titlePage` flag for `<w:titlePg/>`. Reuse the existing header/footer part creation and parsing path, parameterized by reference type.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Header/Footer Reference Types

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a writer test near the existing header/footer tests:

```ts
  it("writes first and even page headers and footers", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          titlePage: true,
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default header" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First header" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even header" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default footer" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First footer" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even footer" }] }],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const firstHeader = await zip.file("word/header2.xml")!.async("string");
    const evenFooter = await zip.file("word/footer3.xml")!.async("string");

    expect(xml).toContain('<w:headerReference w:type="default" r:id="rIdHeader1"/>');
    expect(xml).toContain('<w:headerReference w:type="first" r:id="rIdHeader2"/>');
    expect(xml).toContain('<w:headerReference w:type="even" r:id="rIdHeader3"/>');
    expect(xml).toContain('<w:footerReference w:type="default" r:id="rIdFooter1"/>');
    expect(xml).toContain('<w:footerReference w:type="first" r:id="rIdFooter2"/>');
    expect(xml).toContain('<w:footerReference w:type="even" r:id="rIdFooter3"/>');
    expect(xml).toContain("<w:titlePg/>");
    expect(firstHeader).toContain("<w:t>First header</w:t>");
    expect(evenFooter).toContain("<w:t>Even footer</w:t>");
    expect(rels).toContain('Target="header2.xml"');
    expect(rels).toContain('Target="footer3.xml"');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a reader round-trip test:

```ts
  it("round-trips first and even page headers and footers", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          titlePage: true,
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default header" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First header" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even header" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default footer" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First footer" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even footer" }] }],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "first and even page headers and footers"`

Expected: FAIL because only default header/footer references are emitted and parsed today.

### Task 2: Implement Header/Footer Reference Types

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add `titlePage?: boolean` to `SectionNode`.

Change `HeaderFooterContent` to:

```ts
export type HeaderFooterContent = {
  default?: ParagraphNode[];
  first?: ParagraphNode[];
  even?: ParagraphNode[];
};
```

- [x] **Step 2: Emit header/footer references**

Make `createHeaderReference` and `createFooterReference` accept a type:

```ts
function createHeaderReference(type: keyof HeaderFooterContent, blocks: ParagraphNode[], context: WriterContext): string
```

In `sectionPropertiesXml`, emit references for `default`, `first`, and `even`, plus `<w:titlePg/>` when `section.titlePage` is true.

- [x] **Step 3: Parse header/footer references**

Parse all `headerReference` and `footerReference` nodes with `asArray`, selecting by `type` with default fallback to `"default"`.

Parse `titlePage` from `sectionProperties.titlePg`.

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-158-section-header-footer-types.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "first and even page headers and footers"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-158-section-header-footer-types.md
git commit -m "feat: add phase 158 section header footer types"
git push -u origin phase-158-section-header-footer-types
```

- [x] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-158-section-header-footer-types
```

- [x] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-158-section-header-footer-types.md
git commit -m "docs: mark phase 158 pushed"
git push
```

## Self-Review

- Spec coverage: Covers first/even/default header and footer references, title page marker, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `titlePage`, `default`, `first`, and `even` consistently across schema, writer, reader, and tests.

## Push Record

- Branch: `phase-158-section-header-footer-types`
- Feature commit: `310ea55 feat: add phase 158 section header footer types`
- PR URL: `https://github.com/codingayice/word2json/pull/new/phase-158-section-header-footer-types`
- RED verification: `npm test -- tests/docx-core.test.ts -t "first and even page headers and footers"` failed with 2 expected failures before implementation.
- GREEN targeted verification: `npm test -- tests/docx-core.test.ts -t "first and even page headers and footers"` passed with 2 tests.
- Full verification: `npm test` passed with 458 tests.
- Build verification: `npm run build` exited 0.
- Whitespace verification: `git diff --check` exited 0 with existing LF/CRLF warnings.
