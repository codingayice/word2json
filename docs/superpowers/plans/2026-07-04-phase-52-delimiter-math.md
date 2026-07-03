# Phase 52 Delimiter Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math delimiter structures so bracketed formulas, absolute values, and parenthesized matrix expressions round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `delimiter` node containing optional beginning and ending delimiters plus nested content nodes. Serialize the node to OMML `m:d` with `m:dPr`, `m:begChr`, `m:endChr`, and `m:e`, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes matrix office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes delimiter office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "(",
                  end: ")",
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
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

    expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:e></m:d>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes delimiter office math runs"`

Expected: FAIL because `MathNode` has no `delimiter` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "delimiter"; begin?: string; end?: string; content: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "delimiter") {
    const delimiterProperties = node.begin !== undefined || node.end !== undefined
      ? `<m:dPr>${node.begin !== undefined ? `<m:begChr m:val="${escapeAttribute(node.begin)}"/>` : ""}${node.end !== undefined ? `<m:endChr m:val="${escapeAttribute(node.end)}"/>` : ""}</m:dPr>`
      : "";
    return `<m:d>${delimiterProperties}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:d>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes delimiter office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips matrix office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips delimiter office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "|",
                  end: "|",
                  content: [
                    {
                      type: "superscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      superscript: [{ type: "text" as const, text: "2" }],
                    },
                  ],
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter office math runs"`

Expected: FAIL because parser omits `m:d` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.d` branch:

```ts
    ...asArray(container.d)
      .map((delimiter) => {
        const delimiterNode = asObject(delimiter);
        const delimiterProperties = asObject(delimiterNode.dPr);
        const begin = asObject(delimiterProperties.begChr).val;
        const end = asObject(delimiterProperties.endChr).val;
        return {
          type: "delimiter" as const,
          ...(typeof begin === "string" ? { begin } : {}),
          ...(typeof end === "string" ? { end } : {}),
          content: parseMathNodes(asObject(delimiterNode.e)),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips delimiter office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-52-delimiter-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-52-delimiter-math.md
git commit -m "feat: add phase 52 delimiter math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-52-delimiter-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-52-delimiter-math.md
git commit -m "docs: mark phase 52 pushed"
git push
```

---

**Pushed Branch:** `phase-52-delimiter-math`

**Implementation Commit:** `4816cce feat: add phase 52 delimiter math`
