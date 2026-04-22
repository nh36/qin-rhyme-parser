#!/usr/bin/env python3
import argparse
import csv
import json
import sys
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description='Deterministic regression canary using manifest')
    parser.add_argument('--manifest', help='Path to manifest JSON (optional)')
    return parser.parse_args()


def resolve_manifest_path(cli_value):
    if cli_value:
        return Path(cli_value)

    manifest_candidates = sorted(Path('outputs').glob('manifest.*.json'))
    if manifest_candidates:
        return manifest_candidates[-1]

    latest_manifest = Path('outputs/latest_manifest.json')
    if latest_manifest.exists():
        return latest_manifest

    return None


def load_manifest(manifest_path):
    try:
        with manifest_path.open('r', encoding='utf-8') as mf:
            return json.load(mf)
    except Exception as exc:
        print(f'Failed to read manifest {manifest_path}: {exc}')
        sys.exit(2)


def resolve_manifest_file(manifest_path, manifest, prefix, suffixes):
    run_id = manifest.get('run_id') or ''
    run_dir = Path('outputs') / f'run_{run_id}' if run_id else None
    search_roots = []

    if manifest_path.parent != Path('.'):
        search_roots.append(manifest_path.parent)
    if run_dir:
        search_roots.append(run_dir)
    search_roots.append(Path('outputs'))

    seen = set()
    for fn in manifest.get('files', []):
        if not fn.startswith(prefix) or not any(fn.endswith(suffix) for suffix in suffixes):
            continue
        for root in search_roots:
            candidate = root / fn
            if candidate in seen:
                continue
            seen.add(candidate)
            if candidate.exists():
                return candidate

    for root in search_roots:
        if not root.exists():
            continue
        for suffix in suffixes:
            matches = sorted(root.glob(f'{prefix}*{suffix}'))
            if matches:
                return matches[-1]

    return None


def read_delimited_rows(path):
    delimiter = '\t' if path.suffix.lower() == '.tsv' else ','
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        return list(reader)


def load_expected_canaries():
    expected_path = Path('scripts/expected_canaries.json')
    if not expected_path.exists():
        print(f'Expected canaries file not found: {expected_path}')
        sys.exit(2)

    try:
        with expected_path.open('r', encoding='utf-8') as ef:
            expected_doc = json.load(ef)
    except Exception as exc:
        print(f'Failed to read expected canaries: {exc}')
        sys.exit(2)

    return expected_doc.get('canaries', [])


def build_rows_index(rows):
    rows_index = {}
    for row in rows:
        try:
            seg = int(row.get('RhymeSegment') or 0)
        except Exception:
            seg = None
        key = (row.get('TableID'), seg)
        rows_index.setdefault(key, []).append(row)
    return rows_index


def segment_candidates(rows_index, seg):
    candidates = []
    for (table_id, row_seg), rows in rows_index.items():
        if row_seg != seg:
            continue
        for row in rows:
            toks = [t.strip() for t in (row.get('rhyme_tokens') or '').split('|') if t.strip()]
            candidates.append({
                'RowID': row.get('RowID'),
                'TableID': table_id,
                'page': row.get('page'),
                'tokens_preview': toks[:10],
            })
    return candidates


def main():
    args = parse_args()
    manifest_path = resolve_manifest_path(args.manifest)
    if not manifest_path or not manifest_path.exists():
        print('No manifest found in outputs/; run parser first')
        sys.exit(2)

    manifest = load_manifest(manifest_path)
    ann_path = resolve_manifest_file(manifest_path, manifest, 'rhyme_annotations.', ('.tsv', '.csv'))
    if not ann_path:
        print('No annotation file found from manifest or outputs/; run parser first')
        sys.exit(2)

    expected_list = load_expected_canaries()
    rows = read_delimited_rows(ann_path)
    rows_index = build_rows_index(rows)
    errors = []

    for item in expected_list:
        table = item.get('TableID')
        seg = item.get('RhymeSegment')
        expected_tokens = item.get('expected_tokens', [])
        expect_tone_count = item.get('expect_tone_count', None)
        key = (table, seg)
        matches = rows_index.get(key, [])

        if not matches:
            cands = segment_candidates(rows_index, seg)
            errors.append(
                f'Expected (TableID={table},RhymeSegment={seg}) not found in {ann_path}; '
                f'candidates for segment {seg}: {cands}'
            )
            continue

        if len(matches) > 1:
            ids = [f"{row.get('RowID')} (page={row.get('page')})" for row in matches]
            errors.append(f'Expected key (TableID={table},RhymeSegment={seg}) is ambiguous; matches: {ids}')
            continue

        row = matches[0]
        toks = [t.strip() for t in (row.get('rhyme_tokens') or '').split('|') if t.strip()]
        tones = [t.strip() for t in (row.get('tone_tokens') or '').split('|') if t.strip()]

        if toks != expected_tokens:
            errors.append(
                f'For (TableID={table},RhymeSegment={seg}) tokens mismatch: '
                f'got {toks}, expected {expected_tokens} (RowID={row.get("RowID")})'
            )
        if expect_tone_count is not None and len(toks) != expect_tone_count:
            errors.append(
                f'For (TableID={table},RhymeSegment={seg}) token count {len(toks)} != '
                f'expected tone count {expect_tone_count} (RowID={row.get("RowID")})'
            )
        if len(toks) != len(tones):
            errors.append(
                f'For (TableID={table},RhymeSegment={seg}) token/tone count mismatch: '
                f'{len(toks)} vs {len(tones)} (RowID={row.get("RowID")})'
            )

    if errors:
        print('REGRESSION TEST FAILED')
        for error in errors:
            print('-', error)
        sys.exit(2)

    print('REGRESSION TEST PASSED')
    sys.exit(0)


if __name__ == '__main__':
    main()
