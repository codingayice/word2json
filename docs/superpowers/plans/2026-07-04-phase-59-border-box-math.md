# Phase 59 Border Box Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math border box structures so visually boxed formulas and hidden-border settings round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `borderBox` node containing nested content and optional border visibility flags. Serialize the node to OMML `m:borderBox` with optional `m:borderBoxPr` flags, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes box office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes border box office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: true,
                  hideBottom: true,
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

    expect(xml).toContain('<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/></m:borderBoxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:borderBox>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes border box office math runs"`

Expected: FAIL because `MathNode` has no `borderBox` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "borderBox"; hideTop?: boolean; hideBottom?: boolean; hideLeft?: boolean; hideRight?: boolean; content: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "borderBox") {
    const properties = [
      node.hideTop ? '<m:hideTop m:val="1"/>' : "",
      node.hideBottom ? '<m:hideBot m:val="1"/>' : "",
      node.hideLeft ? '<m:hideLeft m:val="1"/>' : "",
      node.hideRight ? '<m:hideRight m:val="1"/>' : "",
    ].join("");
    return `<m:borderBox>${properties ? `<m:borderBoxPr>${properties}</m:borderBoxPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:borderBox>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes border box office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips box office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips border box office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideLeft: true,
                  hideRight: true,
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips border box office math runs"`

Expected: FAIL because parser omits `m:borderBox` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.borderBox` branch using existing math boolean helpers:

```ts
    ...asArray(container.borderBox)
      .map((borderBox) => {
        const borderBoxNode = asObject(borderBox);
        const borderBoxProperties = asObject(borderBoxNode.borderBoxPr);
        return {
          type: "borderBox" as const,
          ...mathBooleanProperty(borderBoxProperties.hideTop, "hideTop"),
          ...mathBooleanProperty(borderBoxProperties.hideBot, "hideBottom"),
          ...mathBooleanProperty(borderBoxProperties.hideLeft, "hideLeft"),
          ...mathBooleanProperty(borderBoxProperties.hideRight, "hideRight"),
          content: parseMathNodes(asObject(borderBoxNode.e)),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips border box office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-59-border-box-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-59-border-box-math.md
git commit -m "feat: add phase 59 border box math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-59-border-box-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-59-border-box-math.md
git commit -m "docs: mark phase 59 pushed"
git push
```

---

**Pushed Branch:** `phase-59-border-box-math`

**Implementation Commit:** `2c5899a feat: add phase 59 border box math`
