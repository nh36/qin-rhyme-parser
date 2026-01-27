Regression testing and canaries

Workflow

1. Run the parser to generate deterministic outputs and a manifest:

   python3 main.py --mode both --outdir outputs --pdf hudie2023_qin_rhymes.pdf

   This creates timestamped files in outputs/, including:
   - rhyme_annotations.<ts>.csv
   - poem_lines.<ts>.csv
   - rhyme_output.<ts>.txt
   - warnings.<ts>.json
   - summary.<ts>.json
   - manifest.<ts>.json and latest_manifest.json (lists exact filenames emitted)

2. Run the manifest-driven regression canary:

   python3 scripts/regression_canaries.py --manifest outputs/manifest.<ts>.json

   If --manifest is omitted the script will use outputs/latest_manifest.json.

What the canary checks

- Exact (TableID, RhymeSegment) matching: the canary compares the row identified by TableID+RhymeSegment against the expected token sequence (including duplicates and order).
- Strict token sequence equality: tokens must match exactly (no deduplication allowed).
- Token-count vs tone-count invariant: len(tokens) must equal len(tone_tokens) for the checked rows.
- If a specified key is missing the canary fails and provides candidate rows for that RhymeSegment (RowID, TableID, page, token preview) to make updating expected_canaries.json straightforward.

Files

- scripts/expected_canaries.json — the expected canary definitions (TableID, RhymeSegment, expected_tokens, expect_tone_count, comment).
- scripts/regression_canaries.py — manifest-driven canary runner.

Notes

- The canary is intentionally strict: any deviation is surfaced as a failing test so regressions are caught early.
- The images manifest and image separation reduces false positives in image integrity checks; glyph substitution images are listed in extracted_images/images_manifest.json.
