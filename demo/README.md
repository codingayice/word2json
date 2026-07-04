# Roundtrip Demo

Run the default sample:

```bash
node demo/roundtrip-demo.mjs
```

Run with your own DOCX:

```bash
node demo/roundtrip-demo.mjs path/to/input.docx demo/out/my-doc
```

The script builds the TypeScript project, imports DOCX to JSON, then exports that JSON back to DOCX.
Generated files are written under `demo/out/`:

- `demo-source.docx` when no input DOCX is provided
- `imported.json`
- `exported.docx`
