# Phase 146 Style Run Explicit Off Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit `false` on/off run properties inside Word style definitions.

**Architecture:** Inline text runs already distinguish omitted properties from explicit off values such as `<w:b w:val="0"/>`. This phase brings style run writing and parsing to the same fidelity for boolean run properties, while leaving underline's existing `w:val="none"` behavior unchanged.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Style Run Explicit Off

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing style run tests:

```ts
  it("writes style run explicit off boolean properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "ExplicitOffText",
          name: "Explicit Off Text",
          run: {
            bold: false,
            italic: false,
            strike: false,
            doubleStrike: false,
            smallCaps: false,
            allCaps: false,
            shadow: false,
            outline: false,
            emboss: false,
            imprint: false,
            rtl: false,
            complexScript: false,
            specVanish: false,
            hidden: false,
            webHidden: false,
            snapToGrid: false,
            noProof: false,
            officeMath: false,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Off", styleId: "ExplicitOffText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:b w:val="0"/><w:i w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:smallCaps w:val="0"/><w:caps w:val="0"/><w:shadow w:val="0"/><w:outline w:val="0"/><w:emboss w:val="0"/><w:imprint w:val="0"/><w:rtl w:val="0"/><w:cs w:val="0"/><w:specVanish w:val="0"/><w:vanish w:val="0"/><w:webHidden w:val="0"/><w:snapToGrid w:val="0"/><w:noProof w:val="0"/><w:oMath w:val="0"/></w:rPr>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing style run round-trip tests:

```ts
  it("round-trips style run explicit off boolean properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "ExplicitOffText",
          name: "Explicit Off Text",
          run: {
            bold: false,
            italic: false,
            strike: false,
            doubleStrike: false,
            smallCaps: false,
            allCaps: false,
            shadow: false,
            outline: false,
            emboss: false,
            imprint: false,
            rtl: false,
            complexScript: false,
            specVanish: false,
            hidden: false,
            webHidden: false,
            snapToGrid: false,
            noProof: false,
            officeMath: false,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Off", styleId: "ExplicitOffText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes style run explicit off"`

Expected: FAIL because style `<w:rPr>` currently omits false on/off properties.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips style run explicit off"`

Expected: FAIL because style parsing currently treats present on/off properties as true and cannot preserve explicit false from the writer.

### Task 2: Implement Style Run Explicit Off

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Emit style run on/off XML with explicit false**

Add a writer helper:

```ts
function styleOnOffXml(tag: string, value: boolean | undefined): string {
  return value !== undefined ? (value ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`) : "";
}
```

Replace the style run true-only boolean entries with calls such as:

```ts
    styleOnOffXml("b", run.bold),
    styleOnOffXml("i", run.italic),
    styleOnOffXml("strike", run.strike),
```

- [x] **Step 2: Parse style run on/off XML with explicit false**

Add a reader helper:

```ts
function parseStyleOnOffRunProperty<K extends keyof StyleRunProperties>(value: unknown, key: K): Partial<Pick<StyleRunProperties, K>> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { [key]: !(val === "0" || val === false || val === "false") } as Partial<Pick<StyleRunProperties, K>>;
}
```

Replace the style run existence-only boolean parsing with calls such as:

```ts
    ...parseStyleOnOffRunProperty(properties.b, "bold"),
    ...parseStyleOnOffRunProperty(properties.i, "italic"),
    ...parseStyleOnOffRunProperty(properties.strike, "strike"),
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run explicit off"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-146-style-run-explicit-off.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "style run explicit off"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-146-style-run-explicit-off.md
git commit -m "feat: add phase 146 style run explicit off"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-146-style-run-explicit-off
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-146-style-run-explicit-off.md
git commit -m "docs: mark phase 146 pushed"
git push
```

## Push Record

- Branch: `phase-146-style-run-explicit-off`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `c5afdefcf48b74f630a9e5d4f520e924a2a037ac`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-146-style-run-explicit-off`
