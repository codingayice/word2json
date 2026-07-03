# Phase 46 Write Protection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve write protection settings so read-only recommendation and write-reservation metadata round-trip through JSON.

**Architecture:** Add optional `DocumentSettings.writeProtection` for `w:writeProtection` attributes. The writer emits the supported attributes into `word/settings.xml`; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentWriteProtectionSettings` and reference it from `DocumentSettings.writeProtection`.
- `src/docx-writer.ts`: emit `w:writeProtection` inside `settingsXml`.
- `src/docx-reader.ts`: parse `w:writeProtection` from `word/settings.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for write protection.

## Task 1: Writer Write Protection

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes write protection settings` with:

```ts
settings: {
  writeProtection: {
    recommended: true,
    cryptProviderType: "rsaFull",
    cryptAlgorithmClass: "hash",
    cryptAlgorithmType: "typeAny",
    cryptAlgorithmSid: 4,
    cryptSpinCount: 100000,
    hash: "FEDCBA9876543210",
    salt: "0011223344556677",
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:writeProtection w:recommended="1" w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4" w:cryptSpinCount="100000" w:hash="FEDCBA9876543210" w:salt="0011223344556677"/>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "write protection settings"`

Expected: FAIL because `w:writeProtection` is not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentWriteProtectionSettings = {
  recommended?: boolean;
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
};
```

and `writeProtection?: DocumentWriteProtectionSettings` to `DocumentSettings`. Emit `w:writeProtection` when any field is present, using `w:recommended="1"` for `true` and `"0"` for `false`.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "write protection settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips write protection settings` using the same `settings.writeProtection` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "write protection settings"`

Expected: FAIL because the reader drops `w:writeProtection`.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
w:writeProtection w:recommended -> writeProtection.recommended
w:writeProtection w:cryptProviderType -> writeProtection.cryptProviderType
w:writeProtection w:cryptAlgorithmClass -> writeProtection.cryptAlgorithmClass
w:writeProtection w:cryptAlgorithmType -> writeProtection.cryptAlgorithmType
w:writeProtection w:cryptAlgorithmSid -> writeProtection.cryptAlgorithmSid
w:writeProtection w:cryptSpinCount -> writeProtection.cryptSpinCount
w:writeProtection w:hash -> writeProtection.hash
w:writeProtection w:salt -> writeProtection.salt
```

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "write protection settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 46 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-46-write-protection.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 46 write protection"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-46-write-protection
```

Expected: branch `phase-46-write-protection` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of common `w:writeProtection` attributes.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `settings.writeProtection`; OOXML uses `w:writeProtection`.
