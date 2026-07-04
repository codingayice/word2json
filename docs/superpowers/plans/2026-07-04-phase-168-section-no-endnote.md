# Phase 168 Section No Endnote Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level endnote suppression.

**Architecture:** Extend `SectionNode` with a boolean `noEndnote` flag, serialize it as `<w:noEndnote/>` inside section properties, and parse it back from `document.xml`. This follows the existing section boolean layout-property pattern used by `bidi`, `rtlGutter`, and `mirrorMargins`.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section No Endnote Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section bidi, RTL gutter, and note-property writer tests:

```ts
it("writes section no endnote", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        noEndnote: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "No endnote here" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:noEndnote/>");
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section bidi, RTL gutter, and note-property round-trip tests:

```ts
it("round-trips section no endnote", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        noEndnote: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "No endnote here" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section no endnote"`

Expected: the writer and round-trip tests fail because section no-endnote is not serialized or parsed yet.

### Task 2: Implement Section No Endnote Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
noEndnote?: boolean;
```

### Task 3: Serialize Section No Endnote

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit no-endnote from section properties**

In `sectionPropertiesXml`, compute:

```ts
const noEndnote = section.noEndnote ? "<w:noEndnote/>" : "";
```

Append it inside `<w:sectPr>` near the other section note/layout boolean properties.

### Task 4: Parse Section No Endnote

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse no-endnote in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const noEndnote = parseNoEndnote(sectPrValue);
```

Include `{ noEndnote: true }` in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseSectionBidi`:

```ts
function parseNoEndnote(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).noEndnote !== undefined ? true : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-168-section-no-endnote.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section no endnote"`

Expected: both tests pass.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: tests and build exit 0; `git diff --check` exits 0. CRLF warnings are acceptable only when the command still exits 0.

- [ ] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-168-section-no-endnote.md
git commit -m "feat: add phase 168 section no endnote"
```

- [ ] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-168-section-no-endnote
```

- [ ] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-168-section-no-endnote
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-168-section-no-endnote.md
git commit -m "docs: mark phase 168 pushed"
git push
```
