# Phase 44 Document Protection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve document protection settings so editing restrictions and protection metadata round-trip through JSON.

**Architecture:** Add optional `DocumentSettings.protection` for `w:documentProtection` attributes. The writer emits the supported attributes into `word/settings.xml`; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentProtectionSettings` and reference it from `DocumentSettings.protection`.
- `src/docx-writer.ts`: emit `w:documentProtection` inside `settingsXml`.
- `src/docx-reader.ts`: parse `w:documentProtection` from `word/settings.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for document protection.

## Task 1: Writer Document Protection

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes document protection settings` with:

```ts
settings: {
  protection: {
    edit: "trackedChanges",
    enforcement: true,
    cryptProviderType: "rsaFull",
    cryptAlgorithmClass: "hash",
    cryptAlgorithmType: "typeAny",
    cryptAlgorithmSid: 4,
    cryptSpinCount: 100000,
    hash: "ABCDEF0123456789",
    salt: "0123456789ABCDEF",
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4" w:cryptSpinCount="100000" w:hash="ABCDEF0123456789" w:salt="0123456789ABCDEF"/>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "document protection settings"`

Expected: FAIL because `w:documentProtection` is not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentProtectionSettings = {
  edit?: "none" | "readOnly" | "comments" | "trackedChanges" | "forms";
  enforcement?: boolean;
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
};
```

and `protection?: DocumentProtectionSettings` to `DocumentSettings`. Emit `w:documentProtection` when any field is present, using `w:enforcement="1"` for `true` and `"0"` for `false`.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "document protection settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips document protection settings` using the same `settings.protection` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "document protection settings"`

Expected: FAIL because the reader drops `w:documentProtection`.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
w:documentProtection w:edit -> protection.edit
w:documentProtection w:enforcement -> protection.enforcement
w:documentProtection w:cryptProviderType -> protection.cryptProviderType
w:documentProtection w:cryptAlgorithmClass -> protection.cryptAlgorithmClass
w:documentProtection w:cryptAlgorithmType -> protection.cryptAlgorithmType
w:documentProtection w:cryptAlgorithmSid -> protection.cryptAlgorithmSid
w:documentProtection w:cryptSpinCount -> protection.cryptSpinCount
w:documentProtection w:hash -> protection.hash
w:documentProtection w:salt -> protection.salt
```

Only preserve known edit modes.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "document protection settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 44 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-44-document-protection.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 44 document protection"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-44-document-protection
```

Expected: branch `phase-44-document-protection` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of common `w:documentProtection` attributes.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `settings.protection`; OOXML uses `w:documentProtection`.
