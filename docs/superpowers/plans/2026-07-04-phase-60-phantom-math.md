# Phase 60 Phantom Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math phantom structures so hidden placeholders and spacing-only formula layout round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `phantom` node containing nested content and optional OMML phantom flags. Serialize the node to OMML `m:phant` with optional `m:phantPr` flags, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes border box office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes phantom office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroWidth: true,
                  content: [{ type: "text" as const, text: "x+y" }],
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

    expect(xml).toContain('<m:phant><m:phantPr><m:show m:val="0"/><m:zeroWid m:val="1"/></m:phantPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:phant>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes phantom office math runs"`

Expected: FAIL because `MathNode` has no `phantom` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "phantom"; show?: boolean; zeroWidth?: boolean; zeroAscent?: boolean; zeroDescent?: boolean; transparent?: boolean; content: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "phantom") {
    const properties = [
      node.show !== undefined ? `<m:show m:val="${node.show ? "1" : "0"}"/>` : "",
      node.zeroWidth ? '<m:zeroWid m:val="1"/>' : "",
      node.zeroAscent ? '<m:zeroAsc m:val="1"/>' : "",
      node.zeroDescent ? '<m:zeroDesc m:val="1"/>' : "",
      node.transparent ? '<m:transp m:val="1"/>' : "",
    ].join("");
    return `<m:phant>${properties ? `<m:phantPr>${properties}</m:phantPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:phant>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes phantom office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips border box office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips phantom office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroAscent: true,
                  zeroDescent: true,
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

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips phantom office math runs"`

Expected: FAIL because parser omits `m:phant` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.phant` branch:

```ts
    ...asArray(container.phant)
      .map((phantom) => {
        const phantomNode = asObject(phantom);
        const phantomProperties = asObject(phantomNode.phantPr);
        return {
          type: "phantom" as const,
          ...mathBooleanProperty(phantomProperties.show, "show"),
          ...mathBooleanProperty(phantomProperties.zeroWid, "zeroWidth"),
          ...mathBooleanProperty(phantomProperties.zeroAsc, "zeroAscent"),
          ...mathBooleanProperty(phantomProperties.zeroDesc, "zeroDescent"),
          ...mathBooleanProperty(phantomProperties.transp, "transparent"),
          content: parseMathNodes(asObject(phantomNode.e)),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips phantom office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-60-phantom-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-60-phantom-math.md
git commit -m "feat: add phase 60 phantom math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-60-phantom-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-60-phantom-math.md
git commit -m "docs: mark phase 60 pushed"
git push
```

---

**Pushed Branch:** `phase-60-phantom-math`

**Implementation Commit:** `9d0fe96 feat: add phase 60 phantom math`
