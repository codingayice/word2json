# Phase 156 Numbering Abstract Identity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word abstract numbering identity and template-link metadata through DOCX to JSON and JSON to DOCX conversion.

**Architecture:** Extend `AbstractNumberingDefinition` with optional metadata fields that map directly to child elements of `<w:abstractNum>`. Keep level parsing and number instance parsing unchanged so this phase is limited to abstract numbering definitions.

**Tech Stack:** TypeScript, JSZip DOCX package generation, fast-xml-parser, Vitest.

---

### Task 1: Add RED Coverage For Abstract Numbering Identity

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Add writer coverage**

Add a test near the custom numbering writer tests:

```ts
  it("writes abstract numbering identity and style links", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 40,
            nsid: "5E2A1C9B",
            multiLevelType: "hybridMultilevel" as const,
            templateCode: "03A54D6C",
            styleLink: "LegalList",
            numberingStyleLink: "LegalListNumbering",
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
            ],
          },
        ],
        nums: [{ id: 40, abstractId: 40 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:abstractNum w:abstractNumId="40"><w:nsid w:val="5E2A1C9B"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="03A54D6C"/><w:styleLink w:val="LegalList"/><w:numStyleLink w:val="LegalListNumbering"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>');
  });
```

- [x] **Step 2: Add round-trip coverage**

Add a test near the custom numbering reader tests:

```ts
  it("round-trips abstract numbering identity and style links", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 40,
            nsid: "5E2A1C9B",
            multiLevelType: "hybridMultilevel" as const,
            templateCode: "03A54D6C",
            styleLink: "LegalList",
            numberingStyleLink: "LegalListNumbering",
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
            ],
          },
        ],
        nums: [{ id: 40, abstractId: 40 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run RED test**

Run: `npm test -- tests/docx-core.test.ts -t "abstract numbering identity and style links"`

Expected: FAIL because abstract numbering identity metadata is not emitted or parsed yet.

### Task 2: Implement Abstract Numbering Identity Fields

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend `AbstractNumberingDefinition`**

Add:

```ts
  nsid?: string;
  multiLevelType?: "singleLevel" | "multilevel" | "hybridMultilevel";
  templateCode?: string;
  styleLink?: string;
  numberingStyleLink?: string;
```

- [x] **Step 2: Emit abstract numbering metadata**

Add an `abstractNumberingXml` helper:

```ts
function abstractNumberingXml(abstractNum: NonNullable<DocumentJson["numbering"]>["abstractNums"][number]): string {
  const metadata = [
    abstractNum.nsid ? `<w:nsid w:val="${escapeAttribute(abstractNum.nsid)}"/>` : "",
    abstractNum.multiLevelType ? `<w:multiLevelType w:val="${abstractNum.multiLevelType}"/>` : "",
    abstractNum.templateCode ? `<w:tmpl w:val="${escapeAttribute(abstractNum.templateCode)}"/>` : "",
    abstractNum.styleLink ? `<w:styleLink w:val="${escapeAttribute(abstractNum.styleLink)}"/>` : "",
    abstractNum.numberingStyleLink ? `<w:numStyleLink w:val="${escapeAttribute(abstractNum.numberingStyleLink)}"/>` : "",
  ].join("");

  return `<w:abstractNum w:abstractNumId="${abstractNum.id}">` +
    metadata +
    abstractNum.levels.map((level) => numberingLevelXml(level)).join("") +
    `</w:abstractNum>`;
}
```

Then use this helper in `numberingXml`.

- [x] **Step 3: Parse abstract numbering metadata**

In `parseAbstractNumberingDefinition`, read metadata from the `abstractNum` node and include it in the returned object:

```ts
  const nsid = asObject(value.nsid);
  const multiLevelType = asObject(value.multiLevelType);
  const template = asObject(value.tmpl);
  const styleLink = asObject(value.styleLink);
  const numberingStyleLink = asObject(value.numStyleLink);
```

Return optional properties before `levels`.

### Task 3: Verify, Commit, Push, And Record

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-156-numbering-abstract-identity.md`

- [x] **Step 1: Run targeted GREEN test**

Run: `npm test -- tests/docx-core.test.ts -t "abstract numbering identity and style links"`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-156-numbering-abstract-identity.md
git commit -m "feat: add phase 156 numbering abstract identity"
git push -u origin phase-156-numbering-abstract-identity
```

- [ ] **Step 4: Record push**

Append a Push Record with the branch, commit hash, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-156-numbering-abstract-identity
```

- [ ] **Step 5: Commit push record**

Run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-156-numbering-abstract-identity.md
git commit -m "docs: mark phase 156 pushed"
git push
```

## Self-Review

- Spec coverage: Covers abstract numbering identity fields, template code, style links, writer, reader, tests, verification, commit, push, and push record.
- Placeholder scan: No placeholders or deferred implementation notes remain.
- Type consistency: Uses `nsid`, `multiLevelType`, `templateCode`, `styleLink`, and `numberingStyleLink` consistently across schema, writer, reader, and tests.
