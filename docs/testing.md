Regression testing and canaries

Workflow

1. Run the parser to generate deterministic outputs and a manifest:

   python3 main.py --mode both --outdir outputs --pdf hudie2023_qin_rhymes.pdf

   This creates a timestamped run directory under `outputs/`, plus a root `latest_manifest.json`:
   - `outputs/run_<ts>/rhyme_annotations.<ts>.tsv`
   - `outputs/run_<ts>/poem_lines.<ts>.csv`
   - `outputs/run_<ts>/rhyme_output.<ts>.tsv`
   - `outputs/run_<ts>/warnings.<ts>.json`
   - `outputs/run_<ts>/summary.<ts>.json`
   - `outputs/run_<ts>/manifest.<ts>.json`
   - `outputs/latest_manifest.json` (records the fresh run id and emitted filenames)

2. Run the manifest-driven regression canary:

   python3 scripts/regression_canaries.py --manifest outputs/run_<ts>/manifest.<ts>.json

   If `--manifest` is omitted the script will use `outputs/latest_manifest.json` and resolve the annotation file from that manifest's `run_id`.

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

Canonical CI/test entrypoint

A canonical, deterministic test run is provided to avoid ambiguity about which outputs file is checked.

- Makefile target (recommended for CI):

  make test

  This runs the parser to produce fresh outputs and then runs the manifest-driven canary against the generated manifest:

    python3 main.py --mode both --outdir outputs --pdf hudie2023_qin_rhymes.pdf && python3 scripts/regression_canaries.py --manifest outputs/latest_manifest.json

- Manual invocation:

    python3 main.py --mode both --outdir outputs --pdf hudie2023_qin_rhymes.pdf
    python3 scripts/regression_canaries.py --manifest outputs/latest_manifest.json

Please use the Makefile target or the above commands to ensure the canary runs against the fresh manifested outputs rather than an arbitrary "newest" file.
