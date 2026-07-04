# Phase 157 Numbering Instance Overrides Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve numbering instance level overrides through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Extend `NumberingInstance` with optional `overrides` that map to `<w:lvlOverride>`. Reuse the existing numbering level serializer/parser for embedded override levels, and support `startOverride` independently for common restart-at-N list instances.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Numbering Instance Overrides

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a test near the numbering writer tests:

```ts
  it("writes numbering instance level overrides", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 50,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1 },
            ],
          },
        ],
        nums: [
          {
            id: 50,
            abstractId: 50,
            overrides: [
              { level: 0, start: 7 },
              {
                level: 1,
                definition: {
                  level: 1,
                  format: "upperRoman" as const,
                  text: "%2.",
                  start: 3,
                  left: 1440,
                  hanging: 360,
                },
              },
            ],
          },
        ],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:num w:numId="50"><w:abstractNumId w:val="50"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:lvlOverride></w:num>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a reader round-trip test:

```ts
  it("round-trips numbering instance level overrides", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 50,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1 },
            ],
          },
        ],
        nums: [
          {
            id: 50,
            abstractId: 50,
            overrides: [
              { level: 0, start: 7 },
              {
                level: 1,
                definition: {
                  level: 1,
                  format: "upperRoman" as const,
                  text: "%2.",
                  start: 3,
                  left: 1440,
                  hanging: 360,
                },
              },
            ],
          },
        ],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering instance level overrides"`

Expected: FAIL because `lvlOverride` and `startOverride` are not emitted or parsed yet.

### Task 2: Implement Numbering Instance Overrides

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend `NumberingInstance`**

Add:

```ts
  overrides?: NumberingLevelOverride[];
```

Create:

```ts
export type NumberingLevelOverride = {
  level: number;
  start?: number;
  definition?: NumberingLevelDefinition;
};
```

- [x] **Step 2: Emit overrides in numbering XML**

Add `numberingInstanceXml` and `numberingLevelOverrideXml` helpers, then use them in `numberingXml`:

```ts
function numberingInstanceXml(num: NonNullable<DocumentJson["numbering"]>["nums"][number]): string {
  return `<w:num w:numId="${num.id}"><w:abstractNumId w:val="${num.abstractId}"/>` +
    (num.overrides ?? []).map((override) => numberingLevelOverrideXml(override)).join("") +
    `</w:num>`;
}
```

- [x] **Step 3: Parse overrides from numbering XML**

In `parseNumbering`, parse each `<w:lvlOverride>` with:

```ts
const overrides = asArray(num.lvlOverride).map((overrideValue) => parseNumberingLevelOverride(overrideValue));
```

Add parsed overrides to the returned `NumberingInstance` only when the array is non-empty.

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-157-numbering-instance-overrides.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering instance level overrides"`

Expected: PASS.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: all commands exit 0. Existing LF/CRLF warnings from `git diff --check` are acceptable if the exit code is 0.

- [ ] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-157-numbering-instance-overrides.md
git commit -m "feat: add phase 157 numbering instance overrides"
git push -u origin phase-157-numbering-instance-overrides
```

- [ ] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-157-numbering-instance-overrides
```

- [ ] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-157-numbering-instance-overrides.md
git commit -m "docs: mark phase 157 pushed"
git push
```

## Self-Review

- Spec coverage: Covers numbering instance level overrides, start overrides, embedded override levels, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `overrides`, `level`, `start`, and `definition` consistently across schema, writer, reader, and tests.
