# Phase 167 Section Bidi Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level bidirectional layout.

**Architecture:** Extend `SectionNode` with a boolean `bidi` flag, serialize it as `<w:bidi/>` inside section properties, and parse it back from `document.xml`. This follows the existing section boolean layout-property pattern used by `mirrorMargins` and `rtlGutter`.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Bidi Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section text direction, RTL gutter, and mirror margins writer tests:

```ts
it("writes section bidi layout", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        bidi: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Bidi section" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:bidi/>");
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section text direction, RTL gutter, and mirror margins round-trip tests:

```ts
it("round-trips section bidi layout", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        bidi: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Bidi section" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section bidi layout"`

Expected: the writer and round-trip tests fail because section bidi layout is not serialized or parsed yet.

### Task 2: Implement Section Bidi Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
bidi?: boolean;
```

### Task 3: Serialize Section Bidi

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit bidi from section properties**

In `sectionPropertiesXml`, compute:

```ts
const bidi = section.bidi ? "<w:bidi/>" : "";
```

Append it inside `<w:sectPr>` near `rtlGutter` and `mirrorMargins`.

### Task 4: Parse Section Bidi

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse bidi in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const bidi = parseSectionBidi(sectPrValue);
```

Include `{ bidi: true }` in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseRtlGutter`:

```ts
function parseSectionBidi(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).bidi !== undefined ? true : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-167-section-bidi.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section bidi layout"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-167-section-bidi.md
git commit -m "feat: add phase 167 section bidi"
```

- [x] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-167-section-bidi
```

- [x] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-167-section-bidi
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-167-section-bidi.md
git commit -m "docs: mark phase 167 pushed"
git push
```

## Push Record

- Branch: `phase-167-section-bidi`
- Feature commit: `1f41e00 feat: add phase 167 section bidi`
- Verification:
  - `npm test -- tests/docx-core.test.ts -t "section bidi layout"`: 2 passed
  - `npm test`: 476 passed
  - `npm run build`: exit 0
  - `git diff --check`: exit 0 with LF/CRLF warnings
- PR: https://github.com/codingayice/word2json/pull/new/phase-167-section-bidi
