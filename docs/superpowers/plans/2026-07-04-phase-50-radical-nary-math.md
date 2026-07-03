# Phase 50 Radical Nary Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve radical and n-ary math structures, covering common root and summation equations in Word formulas.

**Architecture:** Extend the recursive `MathNode` union with `radical` and `nary` nodes. Serialize them into OMML `m:rad` and `m:nary` elements, then parse those same elements back from `m:oMath` while preserving the existing text/fraction/superscript/subscript behavior.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes structured office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes radical and n-ary office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
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

    expect(xml).toContain("<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>");
    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes radical and n-ary office math runs"`

Expected: FAIL because `MathNode` does not yet include `radical` or `nary`, and writer has no output branch.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "radical"; degree?: MathNode[]; content: MathNode[] }
  | { type: "nary"; operator: "sum"; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] };
```

In `src/docx-writer.ts`, add `mathNodeXml` branches:

```ts
  if (node.type === "radical") {
    return `<m:rad>${node.degree ? `<m:deg>${node.degree.map((child) => mathNodeXml(child)).join("")}</m:deg>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:rad>`;
  }

  if (node.type === "nary") {
    return `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr>` +
      (node.lowerLimit ? `<m:sub>${node.lowerLimit.map((child) => mathNodeXml(child)).join("")}</m:sub>` : "") +
      (node.upperLimit ? `<m:sup>${node.upperLimit.map((child) => mathNodeXml(child)).join("")}</m:sup>` : "") +
      `<m:e>${node.body.map((child) => mathNodeXml(child)).join("")}</m:e></m:nary>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes radical and n-ary office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips structured office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips radical and n-ary office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips radical and n-ary office math runs"`

Expected: FAIL because parser omits `rad` and `nary` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` to parse:

```ts
...asArray(container.rad).map(...)
...asArray(container.nary).map(...)
```

For `nary`, read `naryPr.chr.val` and map `∑` to `operator: "sum"`.

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips radical and n-ary office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-50-radical-nary-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-50-radical-nary-math.md
git commit -m "feat: add phase 50 radical nary math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-50-radical-nary-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-50-radical-nary-math.md
git commit -m "docs: mark phase 50 pushed"
git push
```

---

**Pushed Branch:** `phase-50-radical-nary-math`

**Implementation Commit:** `13b329c feat: add phase 50 radical nary math`
