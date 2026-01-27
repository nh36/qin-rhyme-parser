#!/usr/bin/env python3
import pdfplumber
import csv
import re
import json
from pathlib import Path
import argparse
from datetime import datetime
from collections import defaultdict
import unicodedata
import sys

# Conservative Qin rhyme parser with table-structure enforcement,
# duplicate-preserving tokenization, header suppression, forward-fill
# of RhymeSegment and Source, and image-placeholder insertion.

# ---- helpers --------------------------------------------------------------

def normalize_text(s):
    if s is None:
        return ''
    return ' '.join(str(s).split())


def split_chars(s):
    """Split a cell containing a list of rhyme characters into tokens.
    Preserve duplicates and order.
    """
    if s is None:
        return []
    parts = re.split(r'[、,，;；\s\n]+', s)
    return [p.strip().strip('"\'') for p in parts if p.strip()]


def split_tones(s):
    if not s:
        return []
    parts = re.split(r'[、,，;；\s\n]+|[—–\-]+', s)
    return [p.strip() for p in parts if p.strip()]


def append_preserve(lst, items):
    """Append items to list preserving duplicates and order."""
    if not items:
        return
    if isinstance(items, (list, tuple)):
        for it in items:
            if it is not None and it != '':
                lst.append(it)
    else:
        if items is not None and items != '':
            lst.append(items)


HEADER_TOKENS = ['韻字', '上古韻部', '中古聲調', '篇名', '韻段']


def is_header_row(cells):
    # If multiple header tokens co-occur, treat as a header row
    text = ' '.join([c for c in cells if c])
    hits = sum(1 for t in HEADER_TOKENS if t in text)
    if hits >= 2:
        return True
    for c in cells:
        if c and any(t == c.strip() for t in HEADER_TOKENS):
            return True
    return False


def is_probable_rhyz_cell(text):
    # Heuristic: contains CJK, not Latin/digits, not long prose
    if not text or not re.search(r'[\u4E00-\u9FFF]', text):
        return False
    if re.search(r'[A-Za-z0-9]', text):
        return False
    if len(re.findall(r'\s+', text)) >= 3:
        return False
    return True


def row_needs_images(cells):
    if not cells:
        return False
    allowed_re = re.compile(r"[^\u3400-\u4DBF\u4E00-\u9FFF0-9A-Za-z\u3000-\u303F\uFF00-\uFFEF\-\—\–\_\\.\\,，。、：；？！《》（）\(\)\s]")
    for c in cells:
        if not c:
            continue
        if '\ufffd' in c:
            return True
        for ch in c:
            if unicodedata.category(ch) in ('Co', 'Cn'):
                return True
        if allowed_re.search(c):
            return True
    return False


def replace_images_in_cells(cells, image_refs):
    """Replace non-allowed glyphs in cells with placeholders [IMG:ID].
    Returns (new_cells, cell_image_map) where cell_image_map is list-of-lists of image ids used in each cell.
    This consumes image_refs in order for a page; if exhausted, last id is re-used.
    """
    if not cells:
        return cells, [[] for _ in cells]
    new_cells = list(cells)
    cell_image_map = [[] for _ in cells]
    if not image_refs:
        return new_cells, cell_image_map
    img_iter = iter(image_refs)
    for idx, c in enumerate(cells):
        if not c:
            continue
        if row_needs_images([c]):
            out = []
            used = []
            for ch in c:
                if re.match(r"[\u3400-\u4DBF\u4E00-\u9FFF0-9A-Za-z\u3000-\u303F\uFF00-\uFFEF\-\—\–\_\\.\\,，。、：；？！《》（）\(\)\s]", ch):
                    out.append(ch)
                else:
                    try:
                        img_id = next(img_iter)
                    except StopIteration:
                        img_id = image_refs[-1]
                    placeholder = f"[IMG:{img_id}]"
                    out.append(placeholder)
                    used.append(img_id)
            new_cells[idx] = ''.join(out)
            cell_image_map[idx] = used
    return new_cells, cell_image_map


def build_rhyme_tokens_and_token_image_map(rhyme_cells_after):
    '''Return rhyme_tokens and a token-level image map (list of image ids per token).
    token_image_map entries are lists of image ids or empty list.
    '''
    tokens = []
    token_image_map = []
    for cell in rhyme_cells_after:
        cell = cell or ''
        # split cell into tokens preserving duplicates
        cell_tokens = split_chars(cell)
        # but cell_tokens might include placeholders like [IMG:ID] inside tokens; handle that
        for t in cell_tokens:
            tokens.append(t)
            imgs = re.findall(r'\[IMG:(.+?)\]', t)
            token_image_map.append(imgs if imgs else [])
    return tokens, token_image_map


# ---- parsing --------------------------------------------------------------

def parse_qin_rhymes(pdf_path):
    rhyme_groups = []
    row_records = []

    with pdfplumber.open(pdf_path) as pdf:
        current = None
        current_source = ''
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            page_images = getattr(page, 'images', None) or []
            img_files = [f"page_{page_num:03d}_img_{j+1:02d}.png" for j in range(len(page_images))]
            img_ids = [f"IMG_p{page_num:03d}_{j+1:03d}" for j in range(len(page_images))]

            for table_idx, table in enumerate(tables, start=1):
                table_id = f'P{page_num:03d}_T{table_idx:03d}'
                for row_index, row in enumerate(table, start=1):
                    if not row:
                        continue
                    cells = [c if isinstance(c, str) else '' for c in row]
                    if all(not normalize_text(c) for c in cells):
                        continue
                    # Drop repeated header rows
                    if is_header_row(cells):
                        continue

                    id_cell = cells[0].strip() if cells[0] else ''
                    group_id = None
                    if id_cell:
                        m = re.search(r'(\d+)', id_cell)
                        if m:
                            try:
                                group_id = int(m.group(1))
                            except Exception:
                                group_id = None

                    # detect inline source (篇名) and forward-fill
                    sources = []
                    for c in cells:
                        if isinstance(c, str) and '《' in c:
                            sources.extend(re.findall(r'《[^》]+》', c))
                    if sources:
                        current_source = ' '.join(sources)

                    effective_source = current_source

                    # forward-fill group id from current group when missing
                    effective_group = group_id if group_id is not None else (current['RhymeSegment'] if current else None)

                    row_text = ' '.join([normalize_text(c) for c in cells[1:] if normalize_text(c)])

                    group_forward_filled = False
                    source_forward_filled = False
                    if group_id is None and current:
                        group_forward_filled = True
                    if not sources and current_source:
                        source_forward_filled = True
                    row_records.append({
                        'page': page_num,
                        'table_id': table_id,
                        'table_index': table_idx,
                        'row_index': row_index,
                        'cells': cells,
                        'group_id': effective_group,
                        'row_text': row_text,
                        'has_images': bool(page_images),
                        'image_refs': img_ids,
                        'image_files': img_files,
                        'source': effective_source,
                        'group_forward_filled': group_forward_filled,
                        'source_forward_filled': source_forward_filled
                    })

                    # structural reconstruction: start or continue rhyme segment groups
                    if group_id is not None:
                        # flush previous
                        if current:
                            rhyme_groups.append({
                                'RhymeSegment': current['RhymeSegment'],
                                'Characters': '、'.join(current['Characters']),
                                'OC_RhymeGroup': '、'.join(current['OC_RhymeGroup']),
                                'Tones': ' '.join(current['Tones']),
                                'Source': current['Source'],
                                'Page': current['Page'],
                                'TableID': current.get('TableID','')
                            })
                        # start new
                        current = {
                            'RhymeSegment': group_id,
                            'Characters': [],
                            'OC_RhymeGroup': [],
                            'Tones': [],
                            'Source': effective_source or '',
                            'Page': page_num,
                            'TableID': table_id
                        }
                        # parse cells into group (preserve duplicates, preserve ordering)
                        for c in cells[1:]:
                            if not c:
                                continue
                            text = normalize_text(c)
                            if not text:
                                continue
                            if '《' in text:
                                # append titles
                                titles = re.findall(r'《[^》]+》', text)
                                if titles:
                                    if current['Source']:
                                        current['Source'] += ' ' + ' '.join(titles)
                                    else:
                                        current['Source'] = ' '.join(titles)
                                continue
                            if '部' in text:
                                parts = [p.strip() for p in re.split(r'[、,，;；\s]+', text) if p.strip()]
                                append_preserve(current['OC_RhymeGroup'], parts)
                                continue
                            if re.search(r'[—–\-]|平|上|去|入', text):
                                tparts = split_tones(text)
                                append_preserve(current['Tones'], tparts)
                                continue
                            # otherwise characters
                            append_preserve(current['Characters'], split_chars(text))
                    else:
                        # continuation row: attach to current group deterministically if present
                        if current:
                            for c in cells[1:]:
                                if not c:
                                    continue
                                text = normalize_text(c)
                                if not text:
                                    continue
                                if '《' in text:
                                    titles = re.findall(r'《[^》]+》', text)
                                    if titles:
                                        if current['Source']:
                                            current['Source'] += ' ' + ' '.join(titles)
                                        else:
                                            current['Source'] = ' '.join(titles)
                                    continue
                                if '部' in text:
                                    parts = [p.strip() for p in re.split(r'[、,，;；\s]+', text) if p.strip()]
                                    append_preserve(current['OC_RhymeGroup'], parts)
                                    continue
                                if re.search(r'[—–\-]|平|上|去|入', text):
                                    tparts = split_tones(text)
                                    append_preserve(current['Tones'], tparts)
                                    continue
                                append_preserve(current['Characters'], split_chars(text))
                        else:
                            # orphan continuation row -> keep record but do not merge
                            pass
        # after pages: flush current group
        if current:
            rhyme_groups.append({
                'RhymeSegment': current['RhymeSegment'],
                'Characters': '、'.join(current['Characters']),
                'OC_RhymeGroup': '、'.join(current['OC_RhymeGroup']),
                'Tones': ' '.join(current['Tones']),
                'Source': current['Source'],
                'Page': current['Page'],
                'TableID': current.get('TableID','')
            })

    return rhyme_groups, row_records


# ---- annotation / line exports -------------------------------------------

def make_annotation_rows(rows):
    ann = []
    per_page_counts = defaultdict(int)
    warnings = []
    summary = {
        'merged_rows': [],
        'merged_count': 0,
        'prose_dropped_rows': [],
        'prose_dropped_count': 0,
        'alignment_warnings_rows': [],
        'alignment_warnings_count': 0,
        'rows_with_images': [],
        'rows_with_images_count': 0,
        'forward_filled_rows': [],
        'forward_filled_count': 0,
        'source_forward_filled_rows': [],
        'source_forward_filled_count': 0,
        'unresolved_rows': []
    }
    for i, r in enumerate(rows):
        # prepare detailed prose-category diagnostics for any prose drop
        def classify_prose_row(r, raw_text, cells):
            reasons = {}
            # header tokens in cells
            if any(any(ht in (c or '') for ht in HEADER_TOKENS) for c in cells):
                reasons['header_row'] = True
            # latin/digits
            if any(re.search(r'[A-Za-z]', (c or '')) for c in cells):
                reasons['contains_latin'] = True
            if any(re.search(r'\d', (c or '')) for c in cells):
                reasons['contains_digits'] = True
            # numeric-dominant rows
            cjk_count = len(re.findall(r'[\u4E00-\u9FFF]', raw_text or ''))
            digit_count = len(re.findall(r'\d', raw_text or ''))
            if digit_count and digit_count > cjk_count:
                reasons['numeric_row'] = True
            # outside-table heuristics
            if re.search(r'[。．！？…：；，,;]', raw_text or '') or (cjk_count >= 6):
                reasons['outside_table'] = True
            # failed table signature: no rhyme-like cells
            candidate = ''
            for c in cells[1:]:
                if c and is_probable_rhyz_cell(c):
                    candidate = c
                    break
            if not candidate:
                reasons['failed_table_signature'] = True
            return reasons

        page = r.get('page')
        per_page_counts[page] += 1
        row_in_page = per_page_counts[page]
        table_id = r.get('table_id','')
        rowid = f'ROW_{table_id}_p{page:03d}_r{r.get("row_index",0):04d}'
        raw_text = r.get('row_text','')
        cells = r.get('cells', [])

        # determine if this row is prose (leaked narrative) or table row
        candidate = ''
        for c in cells[1:]:
            if c and is_probable_rhyz_cell(c):
                candidate = c
                break
        is_prose = not bool(candidate)

        # classify cells and remember indices for rhyme cells
        rhyme_cells = []
        rhyme_cell_indices = []
        oc_groups = []
        tone_cells = []
        sources = []
        for idx_cell, c in enumerate(cells[1:], start=1):
            if not c:
                continue
            if '《' in c:
                sources.extend(re.findall(r'《[^》]+》', c))
                continue
            if '部' in c:
                parts = [p.strip() for p in re.split(r'[、,，;；\s]+', c) if p.strip()]
                oc_groups.extend(parts)
                continue
            if re.search(r'[—–\-]|平|上|去|入', c):
                tone_cells.append(c)
                continue
            if is_probable_rhyz_cell(c):
                rhyme_cells.append(c)
                rhyme_cell_indices.append(idx_cell)
            else:
                # ambiguous token: if no rhyme_cells yet, consider it as rhyme cell and record index
                if not rhyme_cells:
                    rhyme_cells.append(c)
                    rhyme_cell_indices.append(idx_cell)

        rhyme_raw = ' '.join(rhyme_cells).strip()
        tones_raw = ' '.join(tone_cells).strip()
        oc_rhyme = '、'.join(oc_groups)
        source = ' '.join(sources) if sources else r.get('source','')

        # insert image placeholders and keep mapping
        image_refs = r.get('image_refs', []) if r.get('has_images', False) else []
        cells_with_imgs, cell_image_map = replace_images_in_cells(cells, image_refs)

        # build rhyme cells list from the originally-classified rhyme cell indices (avoid re-classifying)
        rhyme_cells_after = []
        for idx_cell in rhyme_cell_indices:
            if idx_cell < len(cells_with_imgs):
                v = cells_with_imgs[idx_cell]
                if v and v.strip():
                    rhyme_cells_after.append(v.strip())

        # build tokens and token-level image map
        rhyme_tokens, token_image_map = build_rhyme_tokens_and_token_image_map(rhyme_cells_after)
        rhyme_raw_after = ' '.join(rhyme_cells_after).strip()
        tone_tokens = split_tones(tones_raw) if tones_raw else []

        alignment_issue = False
        # header hard-failure: if header tokens appear in rhyme_raw or tokens, mark unresolved and skip
        header_in_raw = any(ht in rhyme_raw_after for ht in HEADER_TOKENS)
        header_in_tokens = any(any(ht in t for ht in HEADER_TOKENS) for t in rhyme_tokens)
        if header_in_raw or header_in_tokens:
            warnings.append({'rowid': rowid, 'reason': 'header_token_in_rhyme', 'raw': rhyme_raw_after, 'rhyme_tokens': rhyme_tokens})
            summary['unresolved_rows'].append({'RowID': rowid, 'TableID': table_id, 'page': page, 'rhyme_raw': rhyme_raw_after, 'rhyme_tokens': rhyme_tokens})
            # do not emit this row into annotation; treat as unresolved
            continue

        # token-tone count validation and simple merge-repair attempt
        if rhyme_tokens and tone_tokens and len(rhyme_tokens) != len(tone_tokens):
            repaired = False
            # try merging with adjacent rows (next then previous) if they belong to same table and same rhyme segment
            for offset in (1, -1):
                j = i + offset
                if j < 0 or j >= len(rows):
                    continue
                r2 = rows[j]
                if r2.get('table_id') != table_id:
                    continue
                if r2.get('group_id') != r.get('group_id'):
                    continue
                # prepare r2 cells with image placeholders
                r2_image_refs = r2.get('image_refs', []) if r2.get('has_images', False) else []
                cells2_with_imgs, cell_image_map2 = replace_images_in_cells(r2.get('cells', []), r2_image_refs)
                # collect additional candidate rhyme text from r2 (non-部, non-tone, non-source cells)
                add_parts = []
                for c2 in cells2_with_imgs[1:]:
                    if not c2:
                        continue
                    if '《' in c2 or '部' in c2 or re.search(r'[—–\\-]|平|上|去|入', c2):
                        continue
                    add_parts.append(c2.strip())
                if not add_parts:
                    continue
                combined_cells = rhyme_cells_after + add_parts
                comb_tokens, comb_token_image_map = build_rhyme_tokens_and_token_image_map(combined_cells)
                if len(comb_tokens) == len(tone_tokens):
                    rhyme_cells_after = combined_cells
                    rhyme_raw_after = ' '.join(rhyme_cells_after).strip()
                    rhyme_tokens = comb_tokens
                    token_image_map = comb_token_image_map
                    repaired = True
                    warnings.append({'rowid': rowid, 'reason': 'repaired_by_merge', 'merged_with_row': j, 'new_rhyme_tokens': rhyme_tokens})
                    summary['merged_rows'].append({'RowID': rowid, 'merged_with_row': j, 'new_rhyme_tokens': rhyme_tokens})
                    summary['merged_count'] += 1
                    break
            if not repaired:
                alignment_issue = True
                warnings.append({'rowid': rowid, 'reason': 'token_tone_mismatch', 'rhyme_tokens': rhyme_tokens, 'tone_tokens': tone_tokens, 'row': r})

        notes_obj = {
            'cells': r.get('cells', []),
            'cells_with_img_placeholders': cells_with_imgs,
            'cell_image_map': cell_image_map,
            'token_image_map': token_image_map,
            'has_images': bool(image_refs)
        }

        # record any per-token or per-cell image anchors
        has_token_images = any(bool(x) for x in token_image_map)
        has_cell_images = any(bool(x) for x in cell_image_map)
        if has_token_images or has_cell_images:
            summary['rows_with_images'].append({'RowID': rowid, 'token_image_map': token_image_map, 'cell_image_map': cell_image_map})
            summary['rows_with_images_count'] += 1

        # if this row looked like prose, skip emitting into table-level annotation (route to prose/warnings)
        if is_prose:
            # classify reason codes
            reasons = classify_prose_row(r, raw_text, cells)
            warnings.append({'rowid': rowid, 'reason': 'classified_as_prose', 'raw_text': raw_text, 'reasons': reasons})
            summary['prose_dropped_rows'].append({'RowID': rowid, 'TableID': table_id, 'page': page, 'row_in_source': row_in_page, 'reason_codes': reasons, 'snippet': (raw_text or '')[:200]})
            summary['prose_dropped_count'] += 1
        else:
            ann.append({
                'RowID': rowid,
                'TableID': table_id,
                'RhymeSegment': r.get('group_id',''),
                'page': page,
                'row_in_source': row_in_page,
                'raw_text': raw_text,
                'rhyme_raw': rhyme_raw_after,
                'OC_RhymeGroup': oc_rhyme,
                'tones_raw': tones_raw,
                'rhyme_tokens': ' | '.join(rhyme_tokens),
                'tone_tokens': ' | '.join(tone_tokens),
                'rhyme_numbers': ' '.join(re.findall(r'\b(\d+)\b', raw_text)),
                'source': source,
                'alignment_issue': '1' if alignment_issue else '',
                'notes': json.dumps(notes_obj, ensure_ascii=False),
                'image_refs': ','.join(image_refs)
            })
    # after row loop, produce sample snippets for prose drops
    prose_samples = summary.get('prose_dropped_rows', [])[:10]
    summary['prose_samples'] = prose_samples
    return ann, warnings, summary


def make_line_rows(rows):
    line_rows = []
    per_page_counts = defaultdict(int)
    for r in rows:
        page = r.get('page')
        per_page_counts[page] += 1
        row_in_page = per_page_counts[page]
        table_id = r.get('table_id','')
        rowid = f'ROW_{table_id}_p{page:03d}_r{r.get("row_index",0):04d}'
        text = r.get('row_text','')
        if not text:
            continue
        total_cjk = len(re.findall(r'[\u4E00-\u9FFF]', text))
        if not (re.search(r'[。．！？…：；，,;]', text) or total_cjk >= 6):
            continue
        row_type = 'text'
        if len(re.findall(r'《[^》]+》', text)) >= 2:
            row_type = 'heading'
        sources = []
        for c in r.get('cells', []):
            if isinstance(c, str) and '《' in c:
                sources.extend(re.findall(r'《[^》]+》', c))
        source = ' '.join(sources) if sources else r.get('source','')
        rhymeids = ' '.join(re.findall(r'\b(\d+)\b', text))
        needs_images = row_needs_images(r.get('cells', [])) and r.get('has_images', False)
        image_refs = r.get('image_refs', []) if needs_images else []
        notes_obj = {'cells': r.get('cells', []), 'has_images': needs_images, 'image_refs': image_refs}
        line_rows.append({
            'RowID': rowid,
            'TableID': table_id,
            'RhymeSegment': r.get('group_id',''),
            'page': page,
            'row_in_source': row_in_page,
            'text': text,
            'row_type': row_type,
            'rhymeids': rhymeids,
            'source': source,
            'notes': json.dumps(notes_obj, ensure_ascii=False),
            'image_refs': ','.join(image_refs)
        })
    return line_rows


# ---- CLI and IO ----------------------------------------------------------

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Parse Qin rhymes and export CLDF-friendly outputs')
    parser.add_argument('--mode', choices=['both','lines-only','annotation-only'], default='both')
    parser.add_argument('--outdir', default='outputs')
    parser.add_argument('--pdf', default='hudie2023_qin_rhymes.pdf')
    args = parser.parse_args()

    ts = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    rhymes, rows = parse_qin_rhymes(args.pdf)

    # build image id -> filename map based on expected extracted_images naming
    image_map = {}
    for r in rows:
        ids = r.get('image_refs', []) or []
        files = r.get('image_files', []) or []
        for i, f in zip(ids, files):
            image_map[i] = str(Path('extracted_images') / f)
    if image_map:
        Path('extracted_images').mkdir(parents=True, exist_ok=True)
        with open(Path('extracted_images') / 'images_index.json', 'w', encoding='utf-8') as mf:
            json.dump([{'id': k, 'file': v} for k, v in image_map.items()], mf, ensure_ascii=False, indent=2)

    # write rhyme segments (stable schema)
    rhyme_fname = outdir / f'rhyme_output.{ts}.txt'
    with open(rhyme_fname, 'w', encoding='utf-8', newline='') as f:
        fieldnames = ['RhymeSegment','Characters','OC_RhymeGroup','Tones','Source','Page','TableID']
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t')
        writer.writeheader()
        for rhyme in rhymes:
            row = {k: rhyme.get(k, '') for k in fieldnames}
            writer.writerow(row)
    print(f'Wrote rhyme groups to {rhyme_fname}')

    # annotation and lines
    all_warnings = []
    if args.mode in ('both','annotation-only'):
        ann, warnings, ann_summary = make_annotation_rows(rows)
        ann_fname = outdir / f'rhyme_annotations.{ts}.csv'
        with open(ann_fname, 'w', encoding='utf-8', newline='') as f:
            fieldnames = ['RowID','TableID','RhymeSegment','page','row_in_source','raw_text','rhyme_raw','OC_RhymeGroup','tones_raw','rhyme_tokens','tone_tokens','rhyme_numbers','source','alignment_issue','notes','image_refs']
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t')
            writer.writeheader()
            for a in ann:
                writer.writerow(a)
        all_warnings.extend(warnings)
        print(f'Wrote annotation table to {ann_fname}')

        # Build images manifest splitting glyph substitution images from other extracted images.
        try:
            referenced_images = set()
            for a in ann:
                try:
                    notes = json.loads(a.get('notes', '{}'))
                except Exception:
                    notes = {}
                for sub in notes.get('token_image_map', []) or []:
                    if isinstance(sub, list):
                        for x in sub:
                            if x:
                                referenced_images.add(x)
                for sub in notes.get('cell_image_map', []) or []:
                    if isinstance(sub, list):
                        for x in sub:
                            if x:
                                referenced_images.add(x)
            ext_idx_path = Path('extracted_images') / 'images_index.json'
            if ext_idx_path.exists():
                try:
                    idx = json.load(open(ext_idx_path, encoding='utf-8'))
                    img_index = {it['id']: it['file'] for it in idx}
                except Exception:
                    img_index = {}
                glyph_sub_images = [k for k in img_index.keys() if k in referenced_images]
                other_images = [k for k in img_index.keys() if k not in referenced_images]
                images_manifest = {
                    'mapping': [{'id': k, 'file': img_index[k]} for k in img_index.keys()],
                    'glyph_substitution_images': glyph_sub_images,
                    'other_extracted_images': other_images
                }
                with open(Path('extracted_images') / 'images_manifest.json', 'w', encoding='utf-8') as imf:
                    json.dump(images_manifest, imf, ensure_ascii=False, indent=2)
        except Exception:
            pass


    if args.mode in ('both','lines-only'):
        lines = make_line_rows(rows)
        lines_fname = outdir / f'poem_lines.{ts}.csv'
        with open(lines_fname, 'w', encoding='utf-8', newline='') as f:
            fieldnames = ['RowID','TableID','RhymeSegment','page','row_in_source','text','row_type','rhymeids','source','notes','image_refs']
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t')
            writer.writeheader()
            for l in lines:
                writer.writerow(l)
        print(f'Wrote line export to {lines_fname}')

    warnings_fname = outdir / f'warnings.{ts}.json'
    with open(warnings_fname, 'w', encoding='utf-8') as wf:
        json.dump(all_warnings, wf, ensure_ascii=False, indent=2)
    if all_warnings:
        print(f'Warnings present: wrote {len(all_warnings)} issues to {warnings_fname}')
    else:
        print(f'Wrote empty warnings file to {warnings_fname}')

    # metadata + README
    meta_ts = {
        'title': 'hudie2023_qin_rhymes - CLDF exports',
        'description': 'Conservative extraction of line-level text and table-level rhyme annotations from hudie2023_qin_rhymes.pdf. Tables only; prose and non-table material are not reconstructed.',
        'license': 'CC-BY-4.0',
        'generated': ts,
        'files': []
    }
    if args.mode in ('both','annotation-only'):
        meta_ts['files'].append(str(ann_fname.name))
    if args.mode in ('both','lines-only'):
        meta_ts['files'].append(str(lines_fname.name))
    meta_ts['files'].append(str(rhyme_fname.name))
    meta_ts['files'].append(str(warnings_fname.name))
    with open(outdir / f'metadata.{ts}.json', 'w', encoding='utf-8') as f:
        json.dump(meta_ts, f, ensure_ascii=False, indent=2)

    meta_latest = {
        'dc:title': meta_ts['title'],
        'dc:description': meta_ts['description'],
        'dc:license': meta_ts['license'],
        'files': meta_ts['files'],
        'note': 'This dataset contains conservative extractions. Most poems are not present and are not reconstructed.'
    }
    with open(outdir / 'metadata.json', 'w', encoding='utf-8') as f:
        json.dump(meta_latest, f, ensure_ascii=False, indent=2)

    with open(outdir / 'README.txt', 'w', encoding='utf-8') as f:
        f.write('This folder contains CLDF-friendly exports derived from hudie2023_qin_rhymes.pdf.\n')
        f.write('NOTE: This extraction is conservative. Many poems are not present in the source tables and have not been reconstructed. Only table-structured rhyme rows are emitted to the annotation outputs.\n')
        f.write('Files: rhyme_output.<ts>.txt (rhyme segments with RhymeSegment, Characters, OC_RhymeGroup, Tones), rhyme_annotations.<ts>.csv (row-level annotations), poem_lines.<ts>.csv (conservative line-level text), warnings.<ts>.json (alignment/header issues).\n')
        f.write('Images: extracted_images/images_index.json maps stable IDs (IMG_pXXX_NNN) to image files; individual rows contain image_refs and in-situ [IMG:ID] placeholders where non-Unicode glyphs occurred.\n')
    print('Wrote metadata and README')

    # build final summary and unresolved files
    parse_forward_filled_count = sum(1 for r in rows if r.get('group_forward_filled'))
    parse_source_forward_filled_count = sum(1 for r in rows if r.get('source_forward_filled'))
    final_summary = {
        'generated': ts,
        'total_rows': len(rows),
        'total_rhyme_groups': len(rhymes),
        'parse_forward_filled_count': parse_forward_filled_count,
        'parse_source_forward_filled_count': parse_source_forward_filled_count,
        'annotation_summary': ann_summary if 'ann_summary' in locals() else {}
    }
    # enrich final_summary with unreferenced images info
    img_idx_path = Path('extracted_images') / 'images_index.json'
    unreferenced = []
    if img_idx_path.exists():
        try:
            idx = json.load(open(img_idx_path, encoding='utf-8'))
            img_index = {it['id']: it['file'] for it in idx}
            # find unreferenced (we already computed these earlier if ann_summary exists)
            referenced = set()
            for rwi in (ann_summary.get('rows_with_images', []) if 'ann_summary' in locals() else []):
                for sub in (rwi.get('token_image_map') or []):
                    if isinstance(sub, list):
                        for x in sub:
                            if x:
                                referenced.add(x)
                for sub in (rwi.get('cell_image_map') or []):
                    if isinstance(sub, list):
                        for x in sub:
                            if x:
                                referenced.add(x)
            for row in rows_list if 'rows_list' in locals() else []:
                ir = (row.get('image_refs') or '').strip()
                if ir:
                    for x in ir.split(','):
                        x = x.strip()
                        if x:
                            referenced.add(x)
            unreferenced = [k for k in img_index.keys() if k not in referenced]
            final_summary['unreferenced_extractions'] = unreferenced
            final_summary['unreferenced_extractions_count'] = len(unreferenced)
        except Exception:
            final_summary['unreferenced_extractions'] = []
            final_summary['unreferenced_extractions_count'] = 0

    summary_fname = outdir / f'summary.{ts}.json'
    with open(summary_fname, 'w', encoding='utf-8') as sf:
        json.dump(final_summary, sf, ensure_ascii=False, indent=2)
    # write manifest describing this run and exact output files
    manifest = {'run_id': ts, 'files': meta_ts.get('files', []) + [summary_fname.name]}
    with open(outdir / f'manifest.{ts}.json', 'w', encoding='utf-8') as mf:
        json.dump(manifest, mf, ensure_ascii=False, indent=2)
    # update latest manifest for deterministic lookup by canary
    with open(outdir / 'latest_manifest.json', 'w', encoding='utf-8') as mf:
        json.dump(manifest, mf, ensure_ascii=False, indent=2)
    print(f'Wrote summary to {summary_fname}')


    # write unresolved file and warn if present
    unresolved = final_summary.get('annotation_summary', {}).get('unresolved_rows', [])
    if unresolved:
        unresolved_fname = outdir / f'unresolved.{ts}.json'
        with open(unresolved_fname, 'w', encoding='utf-8') as uf:
            json.dump(unresolved, uf, ensure_ascii=False, indent=2)
        print(f'Unresolved rows present: {len(unresolved)}. Wrote {unresolved_fname}. Exiting with failure code.')
        sys.exit(2)
    else:
        # if unreferenced images exist, warn but continue
        if final_summary.get('unreferenced_extractions_count', 0) > 0:
            print(f"Warning: {final_summary['unreferenced_extractions_count']} unreferenced extracted images: {final_summary['unreferenced_extractions'][:10]}")

