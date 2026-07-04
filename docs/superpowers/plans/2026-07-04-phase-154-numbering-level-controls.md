# Phase 154 Numbering Level Controls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve additional Word numbering level controls through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Extend the existing `NumberingLevelDefinition` model and keep the conversion logic inside the current numbering reader/writer functions. The writer emits OOXML level controls in stable order, and the reader maps the same elements back into JSON without affecting built-in numbering behavior.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Numbering Level Controls

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a test near the existing custom numbering writer coverage:

```ts
  it("writes numbering level suffix restart and legal controls", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 20,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1)",
                start: 3,
                suffix: "space" as const,
                restart: 2,
                legal: true,
                left: 720,
                hanging: 360,
              },
              {
                level: 1,
                format: "lowerLetter" as const,
                text: "%2.",
                legal: false,
              },
            ],
          },
        ],
        nums: [{ id: 20, abstractId: 20 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/><w:suff w:val="space"/><w:lvlRestart w:val="2"/><w:isLgl/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>');
    expect(numbering).toContain('<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:isLgl w:val="0"/></w:lvl>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a test near the existing custom numbering round-trip coverage:

```ts
  it("round-trips numbering level suffix restart and legal controls", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 20,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1)",
                start: 3,
                suffix: "space" as const,
                restart: 2,
                legal: true,
                left: 720,
                hanging: 360,
              },
              {
                level: 1,
                format: "lowerLetter" as const,
                text: "%2.",
                legal: false,
              },
            ],
          },
        ],
        nums: [{ id: 20, abstractId: 20 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering level suffix restart and legal controls"`

Expected: FAIL because `suffix`, `restart`, and `legal` are not yet emitted or parsed.

### Task 2: Implement Numbering Level Control Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend the JSON schema type**

Add optional numbering level fields:

```ts
  suffix?: "nothing" | "space" | "tab";
  restart?: number;
  legal?: boolean;
```

- [x] **Step 2: Emit the fields in numbering XML**

In `numberingLevelXml`, emit:

```ts
  const suffix = level.suffix ? `<w:suff w:val="${level.suffix}"/>` : "";
  const restart = level.restart !== undefined ? `<w:lvlRestart w:val="${level.restart}"/>` : "";
  const legal = level.legal !== undefined ? (level.legal ? "<w:isLgl/>" : '<w:isLgl w:val="0"/>') : "";
```

Place them after `<w:lvlText>` and before paragraph properties.

- [x] **Step 3: Parse the fields from numbering XML**

In `parseAbstractNumberingDefinition`, read `suff`, `lvlRestart`, and `isLgl`.

Add a small helper that treats a present empty on/off element as `true`, while preserving explicit false values:

```ts
function parseOptionalOnOff(value: unknown): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }

  const rawValue = asObject(value).val;
  return rawValue === "0" || rawValue === false || rawValue === "false" || rawValue === "off" ? false : true;
}
```

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-154-numbering-level-controls.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "numbering level suffix restart and legal controls"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-154-numbering-level-controls.md
git commit -m "feat: add phase 154 numbering level controls"
git push -u origin phase-154-numbering-level-controls
```

- [ ] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-154-numbering-level-controls
```

- [ ] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-154-numbering-level-controls.md
git commit -m "docs: mark phase 154 pushed"
git push
```

## Self-Review

- Spec coverage: Covers schema, writer, reader, targeted tests, full verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `suffix`, `restart`, and `legal` consistently across schema, writer, reader, and tests.
