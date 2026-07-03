# Phase 47 Math Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve a focused subset of Office Math document settings in `word/settings.xml` through JSON -> DOCX and DOCX -> JSON.

**Architecture:** Add a `settings.math` object to the public schema, serialize it as `m:mathPr` children inside `word/settings.xml`, and parse the same children back from namespace-stripped XML. Keep this phase scoped to document-level math defaults, not math body content.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML settings part.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes write protection settings` in `tests/docx-core.test.ts`:

```ts
  it("writes math settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        math: {
          mathFont: "Cambria Math",
          breakBinary: "before" as const,
          smallFraction: true,
          displayDefaults: true,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Math settings" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(settings).toContain("<m:mathPr>");
    expect(settings).toContain('<m:mathFont m:val="Cambria Math"/>');
    expect(settings).toContain('<m:brkBin m:val="before"/>');
    expect(settings).toContain('<m:smallFrac m:val="1"/>');
    expect(settings).toContain("<m:dispDef/>");
    expect(settings).toContain("</m:mathPr>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes math settings"`

Expected: FAIL because TypeScript does not know `settings.math` or generated XML has no `m:mathPr`.

- [x] **Step 3: Add the minimal writer implementation**

Add `math?: DocumentMathSettings;` to `DocumentSettings` in `src/schema.ts` and define:

```ts
export type DocumentMathSettings = {
  mathFont?: string;
  breakBinary?: "before" | "after" | "repeat";
  smallFraction?: boolean;
  displayDefaults?: boolean;
};
```

In `src/docx-writer.ts`, make `settingsXml` add the math namespace only when `settings.math` exists, append `mathSettingsXml(settings.math)`, and implement:

```ts
function mathSettingsXml(math: NonNullable<DocumentJson["settings"]>["math"]): string {
  if (!math) {
    return "";
  }

  const children = [
    math.mathFont ? `<m:mathFont m:val="${escapeAttribute(math.mathFont)}"/>` : "",
    math.breakBinary ? `<m:brkBin m:val="${escapeAttribute(math.breakBinary)}"/>` : "",
    math.smallFraction !== undefined ? `<m:smallFrac m:val="${math.smallFraction ? "1" : "0"}"/>` : "",
    math.displayDefaults ? "<m:dispDef/>" : "",
  ].join("");

  return children ? `<m:mathPr>${children}</m:mathPr>` : "";
}
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes math settings"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips write protection settings` in `tests/docx-core.test.ts`:

```ts
  it("round-trips math settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        math: {
          mathFont: "Cambria Math",
          breakBinary: "before" as const,
          smallFraction: true,
          displayDefaults: true,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Math settings" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips math settings"`

Expected: FAIL because parser omits `settings.math`.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, call `parseMathSettings(settings)` inside `parseSettings`, include `...(math ? { math } : {})`, and implement:

```ts
function parseMathSettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["math"] | undefined {
  const mathPr = asObject(settings.mathPr);
  const mathFont = asObject(mathPr.mathFont);
  const breakBinary = asObject(mathPr.brkBin);
  const smallFraction = asObject(mathPr.smallFrac);
  const result = {
    ...(typeof mathFont.val === "string" ? { mathFont: mathFont.val } : {}),
    ...(breakBinaryValue(breakBinary.val) ? { breakBinary: breakBinaryValue(breakBinary.val) } : {}),
    ...(smallFraction.val !== undefined ? { smallFraction: parseOnOff(smallFraction.val) } : {}),
    ...(mathPr.dispDef !== undefined ? { displayDefaults: true } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function breakBinaryValue(value: unknown): NonNullable<NonNullable<DocumentJson["settings"]>["math"]>["breakBinary"] | undefined {
  return value === "before" || value === "after" || value === "repeat" ? value : undefined;
}
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips math settings"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-47-math-settings.md`

- [x] **Step 1: Run full verification**

Run:

```powershell
npm test
npm run build
git diff --check
```

Expected: all tests pass, build exits 0, diff check exits 0 with only existing Windows line-ending warnings if any.

- [ ] **Step 2: Commit implementation**

Run:

```powershell
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-47-math-settings.md
git commit -m "feat: add phase 47 math settings"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-47-math-settings`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-47-math-settings.md
git commit -m "docs: mark phase 47 pushed"
git push
```

---

**Pushed:** `phase-47-math-settings`

**Implementation Commit:** `874bf90 feat: add phase 47 math settings`
