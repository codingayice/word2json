# Phase 54 Bar Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math bar structures so overline and underline formulas round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `bar` node containing a bar position and nested content nodes. Serialize the node to OMML `m:bar` with `m:barPr`, `m:pos`, and `m:e`, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes accent office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes bar office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "top",
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

    expect(xml).toContain('<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:bar>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes bar office math runs"`

Expected: FAIL because `MathNode` has no `bar` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "bar"; position: "top" | "bottom"; content: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "bar") {
    return `<m:bar><m:barPr><m:pos m:val="${node.position}"/></m:barPr><m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:bar>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes bar office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips accent office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips bar office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "bottom" as const,
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips bar office math runs"`

Expected: FAIL because parser omits `m:bar` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.bar` branch and helper:

```ts
    ...asArray(container.bar)
      .map((bar) => {
        const barNode = asObject(bar);
        return {
          type: "bar" as const,
          position: barPositionValue(asObject(asObject(barNode.barPr).pos).val),
          content: parseMathNodes(asObject(barNode.e)),
        };
      }),
```

```ts
function barPositionValue(value: unknown): "top" | "bottom" {
  return value === "bottom" ? "bottom" : "top";
}
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips bar office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-54-bar-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-54-bar-math.md
git commit -m "feat: add phase 54 bar math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-54-bar-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-54-bar-math.md
git commit -m "docs: mark phase 54 pushed"
git push
```

---

**Pushed Branch:** `phase-54-bar-math`

**Implementation Commit:** `8cf5e82 feat: add phase 54 bar math`
