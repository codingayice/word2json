# Phase 58 Box Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math box structures so boxed formulas and border-control layout details round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `box` node containing nested content and optional border visibility flags. Serialize the node to OMML `m:box` with optional `m:boxPr` flags, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes equation array office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes box office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
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

    expect(xml).toContain('<m:box><m:boxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/></m:boxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:box>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes box office math runs"`

Expected: FAIL because `MathNode` has no `box` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "box"; hideTop?: boolean; hideBottom?: boolean; hideLeft?: boolean; hideRight?: boolean; content: MathNode[] };
```

In `src/docx-writer.ts`, add helpers and a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "box") {
    const properties = [
      node.hideTop ? '<m:hideTop m:val="1"/>' : "",
      node.hideBottom ? '<m:hideBot m:val="1"/>' : "",
      node.hideLeft ? '<m:hideLeft m:val="1"/>' : "",
      node.hideRight ? '<m:hideRight m:val="1"/>' : "",
    ].join("");
    return `<m:box>${properties ? `<m:boxPr>${properties}</m:boxPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:box>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes box office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips equation array office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips box office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideLeft: true,
                  hideRight: true,
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips box office math runs"`

Expected: FAIL because parser omits `m:box` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.box` branch and helpers:

```ts
    ...asArray(container.box)
      .map((box) => {
        const boxNode = asObject(box);
        const boxProperties = asObject(boxNode.boxPr);
        return {
          type: "box" as const,
          ...mathBooleanProperty(boxProperties.hideTop, "hideTop"),
          ...mathBooleanProperty(boxProperties.hideBot, "hideBottom"),
          ...mathBooleanProperty(boxProperties.hideLeft, "hideLeft"),
          ...mathBooleanProperty(boxProperties.hideRight, "hideRight"),
          content: parseMathNodes(asObject(boxNode.e)),
        };
      }),
```

```ts
function mathBooleanProperty<K extends string>(node: unknown, key: K): Partial<Record<K, true>> {
  if (mathBooleanValue(node)) {
    return { [key]: true } as Partial<Record<K, true>>;
  }
  return {};
}

function mathBooleanValue(node: unknown): boolean {
  if (node === undefined) {
    return false;
  }
  const value = asObject(node).val;
  return value === undefined || value === "1" || value === true;
}
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips box office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-58-box-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-58-box-math.md
git commit -m "feat: add phase 58 box math"
```

- [ ] **Step 3: Push branch**

Run: `git push -u origin phase-58-box-math`

- [ ] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-58-box-math.md
git commit -m "docs: mark phase 58 pushed"
git push
```
