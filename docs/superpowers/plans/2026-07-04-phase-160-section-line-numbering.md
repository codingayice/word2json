# Phase 160 Section Line Numbering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve section line numbering settings through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Add `lineNumbering` to `SectionNode` and map it directly to `<w:lnNumType>`. Follow the Phase 159 `pageNumbering` pattern so section layout metadata remains localized in section properties read/write code.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Section Line Numbering

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a writer test near section layout tests:

```ts
  it("writes section line numbering settings", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          lineNumbering: {
            start: 5,
            countBy: 2,
            distance: 360,
            restart: "newPage" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Draft" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:lnNumType w:start="5" w:countBy="2" w:distance="360" w:restart="newPage"/>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a reader round-trip test:

```ts
  it("round-trips section line numbering settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          lineNumbering: {
            start: 5,
            countBy: 2,
            distance: 360,
            restart: "newPage" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Draft" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "section line numbering settings"`

Expected: FAIL because `<w:lnNumType>` is not emitted or parsed yet.

### Task 2: Implement Section Line Numbering

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add to `SectionNode`:

```ts
  lineNumbering?: SectionLineNumbering;
```

Add:

```ts
export type SectionLineNumbering = {
  start?: number;
  countBy?: number;
  distance?: number;
  restart?: "continuous" | "newPage" | "newSection";
};
```

- [x] **Step 2: Emit `<w:lnNumType>`**

Add `lineNumberingXml(section.lineNumbering)` to `sectionPropertiesXml`, near `pageNumberingXml`.

- [x] **Step 3: Parse `<w:lnNumType>`**

Add `parseLineNumbering(sectionProperties)` and include its result in parsed sections.

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-160-section-line-numbering.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "section line numbering settings"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-160-section-line-numbering.md
git commit -m "feat: add phase 160 section line numbering"
git push -u origin phase-160-section-line-numbering
```

- [x] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-160-section-line-numbering
```

- [x] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-160-section-line-numbering.md
git commit -m "docs: mark phase 160 pushed"
git push
```

## Self-Review

- Spec coverage: Covers line numbering start, count interval, text distance, restart mode, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `lineNumbering`, `start`, `countBy`, `distance`, and `restart` consistently across schema, writer, reader, and tests.

## Push Record

- Branch: `phase-160-section-line-numbering`
- Feature commit: `a04b128 feat: add phase 160 section line numbering`
- PR URL: `https://github.com/codingayice/word2json/pull/new/phase-160-section-line-numbering`
- RED verification: `npm test -- tests/docx-core.test.ts -t "section line numbering settings"` failed with 2 expected failures before implementation.
- GREEN targeted verification: `npm test -- tests/docx-core.test.ts -t "section line numbering settings"` passed with 2 tests.
- Full verification: `npm test` passed with 462 tests.
- Build verification: `npm run build` exited 0.
- Whitespace verification: `git diff --check` exited 0 with existing LF/CRLF warnings.
