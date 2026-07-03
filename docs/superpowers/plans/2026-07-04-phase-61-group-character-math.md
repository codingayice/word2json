# Phase 61 Group Character Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math group-character structures so overbraces, underbraces, and similar grouped annotations round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `groupCharacter` node containing the grouping character, position, vertical justification, and nested content. Serialize the node to OMML `m:groupChr` with `m:groupChrPr`, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes phantom office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes group character office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏞",
                  position: "top",
                  verticalJustification: "bottom",
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

    expect(xml).toContain('<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bottom"/></m:groupChrPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:groupChr>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes group character office math runs"`

Expected: FAIL because `MathNode` has no `groupCharacter` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "groupCharacter"; character?: string; position?: "top" | "bottom"; verticalJustification?: "top" | "bottom"; content: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "groupCharacter") {
    const properties = [
      node.character !== undefined ? `<m:chr m:val="${escapeAttribute(node.character)}"/>` : "",
      node.position !== undefined ? `<m:pos m:val="${node.position}"/>` : "",
      node.verticalJustification !== undefined ? `<m:vertJc m:val="${node.verticalJustification}"/>` : "",
    ].join("");
    return `<m:groupChr>${properties ? `<m:groupChrPr>${properties}</m:groupChrPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:groupChr>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes group character office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips phantom office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips group character office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏟",
                  position: "bottom" as const,
                  verticalJustification: "top" as const,
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips group character office math runs"`

Expected: FAIL because parser omits `m:groupChr` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.groupChr` branch and helpers:

```ts
    ...asArray(container.groupChr)
      .map((groupCharacter) => {
        const groupCharacterNode = asObject(groupCharacter);
        const groupCharacterProperties = asObject(groupCharacterNode.groupChrPr);
        const character = asObject(groupCharacterProperties.chr).val;
        const position = groupCharacterPositionValue(asObject(groupCharacterProperties.pos).val);
        const verticalJustification = groupCharacterVerticalJustificationValue(asObject(groupCharacterProperties.vertJc).val);
        return {
          type: "groupCharacter" as const,
          ...(typeof character === "string" ? { character } : {}),
          ...(position ? { position } : {}),
          ...(verticalJustification ? { verticalJustification } : {}),
          content: parseMathNodes(asObject(groupCharacterNode.e)),
        };
      }),
```

```ts
function groupCharacterPositionValue(value: unknown): "top" | "bottom" | undefined {
  return value === "top" || value === "bottom" ? value : undefined;
}

function groupCharacterVerticalJustificationValue(value: unknown): "top" | "bottom" | undefined {
  return value === "top" || value === "bottom" ? value : undefined;
}
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips group character office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-61-group-character-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-61-group-character-math.md
git commit -m "feat: add phase 61 group character math"
```

- [ ] **Step 3: Push branch**

Run: `git push -u origin phase-61-group-character-math`

- [ ] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-61-group-character-math.md
git commit -m "docs: mark phase 61 pushed"
git push
```
