# Phase 49 Structured Office Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve structured OMML math constructs for fractions, superscripts, and subscripts through JSON -> DOCX and DOCX -> JSON.

**Architecture:** Extend `MathRun` from a text-only payload to a recursive `MathNode[]` expression tree while keeping `math.text` backward compatible. Serialize each `MathNode` into the matching OMML element (`m:f`, `m:sSup`, `m:sSub`, `m:r`) and parse those same elements back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes inline office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes structured office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "superscript" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "subscript" as const,
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>");
    expect(xml).toContain("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>");
    expect(xml).toContain("<m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes structured office math runs"`

Expected: FAIL because `MathRun` does not yet accept `nodes` and writer only emits `math.text`.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, replace `MathRun` with:

```ts
export type MathRun = {
  text?: string;
  nodes?: MathNode[];
};

export type MathNode =
  | { type: "text"; text: string }
  | { type: "fraction"; numerator: MathNode[]; denominator: MathNode[] }
  | { type: "superscript"; base: MathNode[]; superscript: MathNode[] }
  | { type: "subscript"; base: MathNode[]; subscript: MathNode[] };
```

In `src/docx-writer.ts`, replace the direct `math.text` emission with a helper:

```ts
function mathRunXml(math: NonNullable<TextRun["math"]>): string {
  const nodes = math.nodes ?? (math.text !== undefined ? [{ type: "text" as const, text: math.text }] : []);
  return `<m:oMath>${nodes.map(mathNodeXml).join("")}</m:oMath>`;
}
```

and implement `mathNodeXml` for `text`, `fraction`, `superscript`, and `subscript`.

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes structured office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips inline office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips structured office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "superscript" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "subscript" as const,
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips structured office math runs"`

Expected: FAIL because parser collapses math into linear text or omits structured nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, update `parseMathRunXml` to return `{ text: "", math: { nodes } }` when `m:oMath` contains `f`, `sSup`, or `sSub`; keep returning `{ text: "", math: { text } }` for the Phase 48 simple-text shape.

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips structured office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-49-structured-office-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-49-structured-office-math.md
git commit -m "feat: add phase 49 structured office math"
```

- [ ] **Step 3: Push branch**

Run: `git push -u origin phase-49-structured-office-math`

- [ ] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-49-structured-office-math.md
git commit -m "docs: mark phase 49 pushed"
git push
```
