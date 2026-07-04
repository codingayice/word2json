# Phase 166 Section RTL Gutter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level right-to-left gutter layout.

**Architecture:** Extend `SectionNode` with a boolean `rtlGutter` flag, serialize it as `<w:rtlGutter/>` inside section properties, and parse it back from `document.xml`. This mirrors the existing section boolean layout-property pattern used by `mirrorMargins`.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section RTL Gutter Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section text direction and mirror margins writer tests:

```ts
it("writes section rtl gutter", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        rtlGutter: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "RTL gutter" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:rtlGutter/>");
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section text direction and mirror margins round-trip tests:

```ts
it("round-trips section rtl gutter", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        rtlGutter: true,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "RTL gutter" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section rtl gutter"`

Expected: the writer and round-trip tests fail because section RTL gutter is not serialized or parsed yet.

### Task 2: Implement Section RTL Gutter Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
rtlGutter?: boolean;
```

### Task 3: Serialize Section RTL Gutter

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit RTL gutter from section properties**

In `sectionPropertiesXml`, compute:

```ts
const rtlGutter = section.rtlGutter ? "<w:rtlGutter/>" : "";
```

Append it inside `<w:sectPr>` near the other section layout boolean properties.

### Task 4: Parse Section RTL Gutter

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse RTL gutter in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const rtlGutter = parseRtlGutter(sectPrValue);
```

Include `{ rtlGutter: true }` in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseMirrorMargins`:

```ts
function parseRtlGutter(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).rtlGutter !== undefined ? true : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-166-section-rtl-gutter.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section rtl gutter"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-166-section-rtl-gutter.md
git commit -m "feat: add phase 166 section rtl gutter"
```

- [x] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-166-section-rtl-gutter
```

- [x] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-166-section-rtl-gutter
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-166-section-rtl-gutter.md
git commit -m "docs: mark phase 166 pushed"
git push
```

## Push Record

- Branch: `phase-166-section-rtl-gutter`
- Feature commit: `f15047b feat: add phase 166 section rtl gutter`
- Verification:
  - `npm test -- tests/docx-core.test.ts -t "section rtl gutter"`: 2 passed
  - `npm test`: 474 passed
  - `npm run build`: exit 0
  - `git diff --check`: exit 0 with LF/CRLF warnings
- PR: https://github.com/codingayice/word2json/pull/new/phase-166-section-rtl-gutter
