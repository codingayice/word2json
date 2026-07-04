# Phase 164 Section Mirror Margins Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level mirrored page margins.

**Architecture:** Extend `SectionNode` with a boolean `mirrorMargins` flag, serialize it as `<w:mirrorMargins/>` inside section properties, and parse it back from `document.xml`. This follows the existing section layout-property pattern used by title page, document grid, and vertical alignment.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Mirror Margins Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section document grid and vertical alignment writer tests:

```ts
it("writes section mirror margins", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        mirrorMargins: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Booklet" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:mirrorMargins/>");
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section document grid and vertical alignment round-trip tests:

```ts
it("round-trips section mirror margins", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        mirrorMargins: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Booklet" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section mirror margins"`

Expected: the writer and round-trip tests fail because section mirror margins are not serialized or parsed yet.

### Task 2: Implement Section Mirror Margins Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
mirrorMargins?: boolean;
```

### Task 3: Serialize Section Mirror Margins

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit mirror margins from section properties**

In `sectionPropertiesXml`, compute:

```ts
const mirrorMargins = section.mirrorMargins ? "<w:mirrorMargins/>" : "";
```

Append it inside `<w:sectPr>` near the other section layout properties.

### Task 4: Parse Section Mirror Margins

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse mirror margins in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const mirrorMargins = parseMirrorMargins(sectPrValue);
```

Include `{ mirrorMargins: true }` in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseSectionVerticalAlignment`:

```ts
function parseMirrorMargins(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).mirrorMargins !== undefined ? true : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-164-section-mirror-margins.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section mirror margins"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-164-section-mirror-margins.md
git commit -m "feat: add phase 164 section mirror margins"
```

- [ ] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-164-section-mirror-margins
```

- [ ] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-164-section-mirror-margins
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-164-section-mirror-margins.md
git commit -m "docs: mark phase 164 pushed"
git push
```
