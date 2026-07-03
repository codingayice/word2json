# Phase 48 Office Math Runs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve simple inline Office Math content in paragraph runs through JSON -> DOCX and DOCX -> JSON.

**Architecture:** Add an optional `math` payload to `TextRun` and serialize it as an inline `m:oMath` node in `word/document.xml`. Parse `m:oMath` back from the paragraph XML token stream so formula position among normal text runs is preserved.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML WordprocessingML and OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after the first paragraph formatting writer tests in `tests/docx-core.test.ts`:

```ts
  it("writes inline office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Equation: " },
          { text: "", math: { text: "x+1=2" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(xml).toContain("<m:oMath>");
    expect(xml).toContain("<m:r><m:t>x+1=2</m:t></m:r>");
    expect(xml).toContain("</m:oMath>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes inline office math runs"`

Expected: FAIL because `TextRun` does not yet support `math`, and writer emits only normal text runs.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, add:

```ts
export type MathRun = {
  text: string;
};
```

and add `math?: MathRun;` to `TextRun`.

In `src/docx-writer.ts`, add the math namespace to the `documentXml` root and add this branch near the top of `runXml` after note references:

```ts
  if (run.math) {
    return `<m:oMath><m:r><m:t>${escapeXml(run.math.text)}</m:t></m:r></m:oMath>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes inline office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near other reader round-trip tests in `tests/docx-core.test.ts`:

```ts
  it("round-trips inline office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Equation: " },
          { text: "", math: { text: "x+1=2" } },
          { text: " solved" },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips inline office math runs"`

Expected: FAIL because parser does not yet include `m:oMath` as a run.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, update `parseParagraphRuns` so that when paragraph XML is available it tokenizes `<w:r>`, `<w:hyperlink>`, and `<m:oMath>` in order. Parse `<m:oMath>` by running the existing XML parser on the token, reading `oMath.r.t`, and returning:

```ts
{ text: "", math: { text: parsedText } }
```

Keep the existing AST fallback for paragraphs without XML.

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips inline office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-48-office-math-runs.md`

- [x] **Step 1: Run full verification**

Run:

```powershell
npm test
npm run build
git diff --check
```

Expected: all tests pass, build exits 0, diff check exits 0 with only Windows line-ending warnings if any.

- [ ] **Step 2: Commit implementation**

Run:

```powershell
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-48-office-math-runs.md
git commit -m "feat: add phase 48 office math runs"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-48-office-math-runs`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-48-office-math-runs.md
git commit -m "docs: mark phase 48 pushed"
git push
```

---

**Pushed:** `phase-48-office-math-runs`

**Implementation Commit:** `9351560 feat: add phase 48 office math runs`
