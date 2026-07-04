# Phase 165 Section Text Direction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level text direction.

**Architecture:** Extend `SectionNode` with `textDirection`, serialize it as `<w:textDirection>` inside section properties, and parse it back from `document.xml`. The value domain reuses the existing paragraph and table-cell text direction values already supported by the codebase.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Text Direction Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section vertical alignment and mirror margins writer tests:

```ts
it("writes section text direction", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        textDirection: "tbRl" as const,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Vertical" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section vertical alignment and mirror margins round-trip tests:

```ts
it("round-trips section text direction", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        textDirection: "btLr" as const,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Rotated" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section text direction"`

Expected: the writer and round-trip tests fail because section text direction is not serialized or parsed yet.

### Task 2: Implement Section Text Direction Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
textDirection?: SectionTextDirection;
```

- [x] **Step 2: Add the shared section type**

Add near the existing section types:

```ts
export type SectionTextDirection = "lrTb" | "tbRl" | "btLr";
```

### Task 3: Serialize Section Text Direction

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit text direction from section properties**

In `sectionPropertiesXml`, compute:

```ts
const textDirection = section.textDirection
  ? `<w:textDirection w:val="${section.textDirection}"/>`
  : "";
```

Append it inside `<w:sectPr>` near the other section layout properties.

### Task 4: Parse Section Text Direction

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse text direction in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const textDirection = parseSectionTextDirection(sectPrValue);
```

Include it in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseSectionVerticalAlignment`:

```ts
function parseSectionTextDirection(sectionPropertiesValue: unknown): SectionNode["textDirection"] | undefined {
  const textDirection = asObject(asObject(sectionPropertiesValue).textDirection);

  return typeof textDirection.val === "string"
    ? textDirection.val as NonNullable<SectionNode["textDirection"]>
    : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-165-section-text-direction.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section text direction"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-165-section-text-direction.md
git commit -m "feat: add phase 165 section text direction"
```

- [x] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-165-section-text-direction
```

- [x] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-165-section-text-direction
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-165-section-text-direction.md
git commit -m "docs: mark phase 165 pushed"
git push
```

## Push Record

- Branch: `phase-165-section-text-direction`
- Feature commit: `62bea81 feat: add phase 165 section text direction`
- Verification:
  - `npm test -- tests/docx-core.test.ts -t "section text direction"`: 2 passed
  - `npm test`: 472 passed
  - `npm run build`: exit 0
  - `git diff --check`: exit 0 with LF/CRLF warnings
- PR: https://github.com/codingayice/word2json/pull/new/phase-165-section-text-direction
