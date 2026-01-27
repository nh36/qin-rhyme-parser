#!/usr/bin/env python3
import csv, glob, sys, os, json, argparse

parser = argparse.ArgumentParser(description='Deterministic regression canary using manifest')
parser.add_argument('--manifest', help='Path to manifest JSON (optional)')
args = parser.parse_args()

manifest_path = args.manifest if 'args' in globals() else None
if not manifest_path:
    candidates = sorted(glob.glob('outputs/manifest.*.json'))
    if candidates:
        manifest_path = candidates[-1]
    elif os.path.exists('outputs/latest_manifest.json'):
        manifest_path = 'outputs/latest_manifest.json'
    else:
        print('No manifest found in outputs/; run parser first')
        sys.exit(2)

try:
    with open(manifest_path, 'r', encoding='utf-8') as mf:
        manifest = json.load(mf)
except Exception as e:
    print(f'Failed to read manifest {manifest_path}: {e}')
    sys.exit(2)

ann = None
for fn in manifest.get('files', []):
    if fn.startswith('rhyme_annotations.') and fn.endswith('.csv'):
        ann = os.path.join('outputs', fn)
        break
if not ann:
    files = sorted(glob.glob('outputs/rhyme_annotations.*.csv'))
    if not files:
        print('No annotation CSV found; run parser first')
        sys.exit(2)
    ann = files[-1]

# load explicit expected canaries file
expected_path = os.path.join('scripts', 'expected_canaries.json')
if not os.path.exists(expected_path):
    print(f'Expected canaries file not found: {expected_path}')
    sys.exit(2)
try:
    with open(expected_path, 'r', encoding='utf-8') as ef:
        expected_doc = json.load(ef)
except Exception as e:
    print(f'Failed to read expected canaries: {e}')
    sys.exit(2)

expected_list = expected_doc.get('canaries', [])
errors = []

# index rows by (TableID, RhymeSegment)
rows_index = {}
with open(ann, 'r', encoding='utf-8') as f:
    rd = csv.DictReader(f, delimiter='\t')
    for row in rd:
        try:
            seg = int(row.get('RhymeSegment') or 0)
        except:
            seg = None
        table = row.get('TableID')
        key = (table, seg)
        rows_index.setdefault(key, []).append(row)

# helper to present candidates for a segment
def segment_candidates(seg):
    candidates = []
    for (t, s), rows in rows_index.items():
        if s == seg:
            for r in rows:
                toks = [t_.strip() for t_ in (r.get('rhyme_tokens') or '').split('|') if t_.strip()]
                candidates.append({'RowID': r.get('RowID'), 'TableID': t, 'page': r.get('page'), 'tokens_preview': toks[:10]})
    return candidates

for item in expected_list:
    table = item.get('TableID')
    seg = item.get('RhymeSegment')
    expected_tokens = item.get('expected_tokens', [])
    expect_tone_count = item.get('expect_tone_count', None)
    key = (table, seg)
    rows = rows_index.get(key, [])
    if not rows:
        # fail with helpful candidate list
        cands = segment_candidates(seg)
        errors.append(f'Expected (TableID={table},RhymeSegment={seg}) not found in {ann}; candidates for segment {seg}: {cands}')
        continue
    if len(rows) > 1:
        ids = [r.get('RowID') + f" (page={r.get('page')})" for r in rows]
        errors.append(f'Expected key (TableID={table},RhymeSegment={seg}) is ambiguous; matches: {ids}')
        continue
    row = rows[0]
    toks = [t.strip() for t in (row.get('rhyme_tokens') or '').split('|') if t.strip()]
    tones = [t.strip() for t in (row.get('tone_tokens') or '').split('|') if t.strip()]
    if toks != expected_tokens:
        errors.append(f'For (TableID={table},RhymeSegment={seg}) tokens mismatch: got {toks}, expected {expected_tokens} (RowID={row.get("RowID")})')
    if expect_tone_count is not None and len(toks) != expect_tone_count:
        errors.append(f'For (TableID={table},RhymeSegment={seg}) token count {len(toks)} != expected tone count {expect_tone_count} (RowID={row.get("RowID")})')
    if len(toks) != len(tones):
        errors.append(f'For (TableID={table},RhymeSegment={seg}) token/tone count mismatch: {len(toks)} vs {len(tones)} (RowID={row.get("RowID")})')

if errors:
    print('REGRESSION TEST FAILED')
    for e in errors:
        print('-', e)
    sys.exit(2)

print('REGRESSION TEST PASSED')
sys.exit(0)
