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

# Image placeholder constants and helpers
def make_img_token(image_id: str) -> str:
    """Create a canonical image placeholder token."""
    return f"⟦IMG:{image_id}⟧"

def is_img_token(tok: str) -> bool:
    """Check if a token is an image placeholder."""
    if not tok:
        return False
    return tok.startswith("⟦IMG:") and tok.endswith("⟧")

def extract_image_id(tok: str) -> str:
    """Extract image ID from a placeholder token."""
    if is_img_token(tok):
        return tok[5:-1]  # Remove ⟦IMG: and ⟧
    return None

def process_rhyme_char_cell_with_images(cell_text, image_refs):
    """Process a rhyme character cell that may have missing characters.
    Detects "、 、" patterns (delimiter-space-delimiter) and inserts image placeholders.
    Also detects leading "、 " which indicates missing first character.
    Returns (processed_text, image_ids_used).
    """
    if not cell_text or not image_refs:
        return cell_text, []
    
    # Check for leading "、 " pattern - missing first character
    if cell_text.startswith('、 '):
        img_iter = iter(image_refs)
        try:
            img_id = next(img_iter)
            placeholder = make_img_token(img_id)
            cell_text = placeholder + cell_text
            image_ids_used = [img_id]
        except StopIteration:
            image_ids_used = []
    else:
        image_ids_used = []
    
    # Now check for "、 、" patterns (consecutive delimiters with space)
    # This is a simple heuristic - look for delimiter followed by space and delimiter
    parts = cell_text.split('、')
    result_parts = []
    img_iter = iter(image_refs[len(image_ids_used):])  # Continue from where we left off
    
    for i, part in enumerate(parts):
        if i == 0:
            # First part - already handled above
            if part:
                result_parts.append(part)
            continue
        
        part_stripped = part.strip()
        if not part_stripped:
            # Empty part in middle - missing character
            try:
                img_id = next(img_iter)
            except StopIteration:
                img_id = image_refs[-1] if image_refs else None
            if img_id:
                result_parts.append(make_img_token(img_id))
                image_ids_used.append(img_id)
        else:
            result_parts.append(part_stripped)
    
    return '、'.join(result_parts), image_ids_used


def get_line_length(text):
    """
    Get the content length of a line, excluding:
    - Parenthetical glosses like （微）
    - Slip IDs
    - Footnote markers [1]
    - Whitespace
    
    Returns: character count
    """
    # Remove parenthetical content
    cleaned = re.sub(r'[（(][^）)]+[）)]', '', text)
    # Remove slip IDs
    cleaned = re.sub(r'\d{1,3}-\d{1,2}(?:背)?', '', cleaned)
    # Remove footnote markers
    cleaned = re.sub(r'\[\d+\]', '', cleaned)
    # Remove whitespace
    cleaned = re.sub(r'\s', '', cleaned)
    # Count CJK characters and punctuation
    return len(cleaned)


def is_verse_like(text, context_lines=None):
    """
    Determine if a line is verse-like based on:
    1. Line length (short, typically 4-10 characters)
    2. Ends with verse punctuation (，。；)
    3. Not a footnote or commentary
    4. If context provided, similar length to nearby lines
    
    Returns: (is_verse, confidence_score)
    """
    if not text or not text.strip():
        return False, 0.0
    
    text = text.strip()
    
    # Definitely not verse
    if is_footnote_or_commentary(text):
        return False, 0.0
    
    # Check for section markers
    if text in ['【用韻情況】', '【註釋】', '【注释】']:
        return False, 0.0
    if strip_note_section_header(text) is not None:
        return False, 0.0
    
    length = get_line_length(text)
    confidence = 0.0
    
    # Length heuristics
    if 1 <= length <= 2 and re.search(r'[，。；：]$', text):
        confidence += 0.25
    elif 3 <= length <= 12:
        confidence += 0.3
    elif 13 <= length <= 20:
        confidence += 0.1
    else:
        # Very long or very short - less likely verse
        confidence -= 0.2
    
    # Ends with verse punctuation
    if re.search(r'[，。；：][^，。；：]*$', text):
        confidence += 0.3
    
    # Has CJK characters
    if re.search(r'[\u4E00-\u9FFF]', text):
        confidence += 0.2
    else:
        return False, 0.0
    
    # Compare with context lines if provided
    if context_lines:
        lengths = [get_line_length(line) for line in context_lines if line]
        if lengths:
            avg_length = sum(lengths) / len(lengths)
            # Similar length to context (within 3 characters)
            if abs(length - avg_length) <= 3:
                confidence += 0.3
            # Check for consistent pattern (all within small range)
            if max(lengths) - min(lengths) <= 4:
                confidence += 0.2
    
    is_verse = confidence > 0.4
    return is_verse, confidence


def is_rhyme_metadata_line(text):
    """Detect rhyme metadata lines that should not be treated as verse."""
    if not text:
        return False

    text = text.strip()

    rhyme_patterns = [
        r'^[\u4E00-\u9FFF、，,\s]{1,16}部(?:獨韻|合韻|通韻)?[——-].+',
        r'^[\u4E00-\u9FFF、，,\s]{1,16}部(?:旁轉|轉韻|葉韻)[。，]?$',
    ]

    return any(re.search(pattern, text) for pattern in rhyme_patterns)


def is_footnote_or_commentary(text):
    """
    Detect footnote markers and scholarly commentary that shouldn't be treated as poetry.
    
    Returns: True if text is a footnote/commentary, False if it's likely a verse line.
    """
    if not text or not text.strip():
        return True
    
    text = text.strip()
    
    # Chinese circled numerals (footnote markers): ①②③④⑤⑥⑦⑧⑨⑩
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', text):
        return True
    
    # Bibliography references with author names and citations
    if re.search(r'[:：]《[^》]+》[,，].*?(年|頁)', text):
        return True
    
    # Rhyme info lines (these are metadata, not verse)
    if is_rhyme_metadata_line(text):
        return True
    
    # Lines that are section titles/headers
    if re.search(r'韻讀$', text):  # e.g., "睡虎地秦墓竹簡韻讀"
        return True
    
    # Chapter/section header patterns
    if re.match(r'^第[一二三四五六七八九十]+章', text):
        return True
    
    # Scholarly commentary phrases
    commentary_patterns = [
        r'^今按[：:]',
        r'^整理者[：:]',
        r'^本段文字',
        r'簡\d+-\d+作[：:]',  # References to other slip versions
        r'嶽麓秦簡.*?作',
        r'睡虎地秦簡.*?作',
        r'從押韻',
        r'用韻.*?需要',
        r'韻式.*?不同',
        r'^關於本段',  # "Regarding this section..."
        r'^關於本韻段',
        r'^具體可參看',  # "For details, see..."
        r'^待考[。，]?$',  # "Pending investigation" - often standalone
        r'^筆者',  # "This author..."
        r'^暫無法',  # "Temporarily unable to..."
        r'^暫且存疑',  # "Temporarily uncertain..."
        r'^值得注意',  # "Worth noting..."
        r'韻語.*?少見',  # Meta-commentary about rhyme usage, e.g., "擇術中韻語少見"
        r'^.*?中韻語',  # Commentary about rhyme words in a text
        r'^此段',  # "This section..."
        r'^例[，,]',  # e.g. "例，我們很難說..."
        r'我們很難說',  # analytical prose
        r'尾句.*?入韻',  # rhyme analysis commentary
        r'^部旁轉[。，]?$',  # short rhyme-analysis note
    ]
    
    for pattern in commentary_patterns:
        if re.search(pattern, text):
            return True
    
    # Lines that are mostly bibliography (author: title, publisher, page)
    # Example: "劉信芳:《〈天水放馬灘秦簡綜述〉質疑》,《文物》1990年第9期"
    if re.search(r'^[^，。]{2,10}[：:]《.+》.*?\d+年', text):
        return True
    
    # Bibliography with URLs
    if re.search(r'http[s]?://', text):
        return True
    
    # Citation format: Name:《Title》
    if re.search(r'^[^，。]{2,10}[：:]《[^》]+》', text):
        return True
    
    # Date patterns (e.g., "1年4月9日", "2010年", "26日")
    if re.search(r'^\d+(?:年\d*月?\d*日?|月\d+日|日)[。，]?$', text):
        return True
    
    # Long commentary-style lines (scholarly discussion, not verse)
    # If line is very long and contains phrases like "認爲", "觀點", "意見"
    if len(text) > 50 and re.search(r'(認爲|觀點|意見|看作|筆者|因而|故而|今按|然而|考其|由此)', text):
        return True
    
    return False


def extract_leading_title(text):
    """Extract a leading local title like 《馬心》 from a line."""
    if not text:
        return '', text

    match = re.match(r'^\s*《([^》]+)》\s*(.*)$', text.strip())
    if not match:
        return '', text

    title = match.group(1).strip()
    remainder = match.group(2).strip()

    # Ignore long bibliographic/article titles that are not local poem labels.
    if (
        len(title) > 20 or
        re.search(r'(研究|修訂|整理|刊佈|版本|編聯|頁|年|博士|學位|論文|的)', title)
    ):
        return '', text

    return title, remainder


def normalize_rhyme_category(label):
    """Normalize rhyme category labels for matching/export."""
    if not label:
        return ''

    category = re.sub(r'(部|獨韻|通韻|合韻)+$', '', label.strip())
    category = re.sub(r'[、,，\s]+', '', category)
    return category


def extract_rhyme_type(label):
    """Extract rhyme type shorthand from a label."""
    if not label:
        return ''
    if '獨韻' in label:
        return '獨'
    if '合韻' in label:
        return '合'
    if '通韻' in label:
        return '通'
    return ''


def extract_titles_from_source(source_text):
    """Extract normalized titles from annotation source strings."""
    if not source_text:
        return []
    return [title.strip() for title in re.findall(r'《([^》]+)》', source_text)]


def strip_note_section_header(text):
    """Remove a leading note-section header and return the remaining payload."""
    if text is None:
        return None
    match = re.match(r'^【(?:註釋|注釋|注释)】\s*(.*)$', text.strip())
    if not match:
        return None
    return match.group(1).strip()


def infer_single_group_rhyme_type(label):
    """Infer 獨 for clear single-group labels like 陽部 when no type is explicit."""
    if not label or extract_rhyme_type(label):
        return ''

    category = normalize_rhyme_category(label)
    if len(category) != 1:
        return ''

    if '部' in label and not any(sep in label for sep in ['、', '，', ',', '/', '／']):
        return '獨'

    return ''


def extract_embedded_note_text(text):
    """Return note text for stray note lines that leaked into verse segments."""
    if not text:
        return ''

    stripped = text.strip()
    payload = strip_note_section_header(stripped)
    if payload is not None:
        stripped = payload

    if not re.match(r'^\[\d+\]', stripped):
        return ''

    note_body = re.sub(r'^\[\d+\]\s*', '', stripped)
    if is_footnote_or_commentary(note_body) or any(
        marker in note_body for marker in ['本段文字', '今按', '整理者', '李零', '王寧', '子居', '：', ':', '《']
    ):
        return stripped

    return ''


def restore_terminal_note_marker(segment, note_number):
    """Restore a bare terminal note number like 吉。11 to 吉。[11] before export."""
    if not segment or not note_number:
        return False

    marker = f'[{note_number}]'
    for line_data in reversed(segment.get('lines', [])):
        text = line_data.get('text', '')
        if not text or marker in text:
            continue
        if re.search(rf'{re.escape(note_number)}\s*$', text):
            line_data['text'] = re.sub(rf'{re.escape(note_number)}\s*$', marker, text)
            return True

    return False


def build_tone_lookup(annotation_rows):
    """Build a title-keyed lookup of rhyme words to tones from table annotations."""
    tone_index = defaultdict(list)

    for row in annotation_rows or []:
        titles = extract_titles_from_source(row.get('Source', ''))
        if not titles:
            continue

        rhyme_tokens = [t.strip() for t in row.get('rhyme_tokens', '').split('|') if t.strip()]
        tone_tokens = [t.strip() for t in row.get('tone_tokens', '').split('|') if t.strip()]
        if not rhyme_tokens or not tone_tokens:
            continue

        record = {
            'group': normalize_rhyme_category(row.get('OC_RhymeGroup', '')),
            'ordered_words': rhyme_tokens,
            'ordered_tones': tone_tokens,
            'word_tones': {word: tone for word, tone in zip(rhyme_tokens, tone_tokens) if word and tone},
            'words': set(rhyme_tokens),
            'rhyme_type': infer_single_group_rhyme_type(row.get('OC_RhymeGroup', '')),
        }

        for title in titles:
            tone_index[title].append(record)

    return tone_index

def filter_rhyme_chars(tokens):
    """Filter a list of tokens to keep only valid rhyme characters and image placeholders.
    Removes line numbers, metadata, parenthetical content, sentence punctuation, counting words.
    Preserves image placeholder tokens.
    """
    filtered = []
    for tok in tokens:
        # Keep image placeholders
        if is_img_token(tok):
            filtered.append(tok)
            continue
        
        # Skip empty
        if not tok or not tok.strip():
            continue
        
        tok_clean = tok.strip()
        
        # Skip line numbers and metadata patterns
        if re.match(r'^\d+$', tok_clean):  # Pure numbers
            continue
        if re.match(r'^\d+\s*[（(]', tok_clean):  # "01（" pattern
            continue
        if re.match(r'^[（(]\d+', tok_clean):  # "（4" pattern
            continue
        
        # Skip parenthetical content (tone markers, etc)
        if tok_clean.startswith('（') and tok_clean.endswith('）'):
            continue
        if tok_clean.startswith('(') and tok_clean.endswith(')'):
            continue
        
        # Skip sentence punctuation
        if tok_clean in ('。', '，', '、', '：', '；', '！', '？', '…'):
            continue
        
        # Skip tone/rhyme markers
        if tok_clean in ('平', '上', '去', '入', '—', '–', '-', '部'):
            continue
        
        # Skip counting words
        if tok_clean in ('一曰', '二曰', '三曰', '四曰', '五曰', '六曰', '七曰', '八曰', '九曰', '十曰'):
            continue
        
        # Otherwise keep it
        filtered.append(tok_clean)
    
    return filtered

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
    """Replace non-allowed glyphs in cells with placeholders ⟦IMG:ID⟧.
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
                    placeholder = make_img_token(img_id)
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

                    # Apply image placeholder replacement to cells before processing
                    cells_with_img, cell_image_map = replace_images_in_cells(cells, img_ids if bool(page_images) else [])

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
                            # Collect image IDs from Characters
                            image_ids_in_chars = []
                            for ch in current['Characters']:
                                if is_img_token(ch):
                                    img_id = extract_image_id(ch)
                                    if img_id and img_id not in image_ids_in_chars:
                                        image_ids_in_chars.append(img_id)
                            
                            # Only include in rhyme_output if Characters field has content
                            # AND doesn't look like a complete poem
                            if current['Characters']:
                                chars_text = '、'.join(current['Characters'])
                                
                                # Exclude if it looks like a poem (not just rhyme characters):
                                # - Contains sentence-ending punctuation
                                # - Has multiple "lines" indicated by repetitive patterns like 一曰、二曰、三曰
                                # - Contains long runs of text without enumeration delimiters
                                is_poem = False
                                
                                if re.search(r'[。．！？…]', chars_text):
                                    is_poem = True
                                elif chars_text.count('一曰') >= 2 or chars_text.count('二曰') >= 2:
                                    is_poem = True
                                elif len(chars_text) > 200:  # Very long suggests concatenated poem lines
                                    is_poem = True
                                
                                if not is_poem:
                                    rhyme_groups.append({
                                        'RhymeSegment': current['RhymeSegment'],
                                        'Characters': chars_text,
                                        'OC_RhymeGroup': '、'.join(current['OC_RhymeGroup']),
                                        'Tones': ' '.join(current['Tones']),
                                        'Source': current['Source'],
                                        'Page': current['Page'],
                                        'TableID': current.get('TableID',''),
                                        'image_refs': '|'.join(image_ids_in_chars)
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
                        # Use cells_with_img which have image placeholders
                        for c in cells_with_img[1:]:
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
                            # Tone cell detection: must be PRIMARILY tone markers
                            tone_chars = len(re.findall(r'[—–\-平上去入]', text))
                            total_chars = len(re.sub(r'\s', '', text))
                            is_tone_cell = (tone_chars > 0 and tone_chars / max(total_chars, 1) > 0.5)
                            
                            if is_tone_cell:
                                tparts = split_tones(text)
                                append_preserve(current['Tones'], tparts)
                                continue
                            # otherwise characters
                            # If row has images, check for missing characters and insert placeholders
                            if img_ids and ('、' in text or ' ' in text):
                                processed_text, _ = process_rhyme_char_cell_with_images(text, img_ids)
                                tokens = split_chars(processed_text)
                            else:
                                tokens = split_chars(text)
                            # Filter to keep only valid rhyme characters and image placeholders
                            filtered_tokens = filter_rhyme_chars(tokens)
                            append_preserve(current['Characters'], filtered_tokens)
                    else:
                        # continuation row: attach to current group deterministically if present
                        if current:
                            for c in cells_with_img[1:]:
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
                                # Tone cell detection: must be PRIMARILY tone markers
                                tone_chars = len(re.findall(r'[—–\-平上去入]', text))
                                total_chars = len(re.sub(r'\s', '', text))
                                is_tone_cell = (tone_chars > 0 and tone_chars / max(total_chars, 1) > 0.5)
                                
                                if is_tone_cell:
                                    tparts = split_tones(text)
                                    append_preserve(current['Tones'], tparts)
                                    continue
                                # Check for character cells with missing characters
                                if img_ids and ('、' in text or ' ' in text):
                                    processed_text, _ = process_rhyme_char_cell_with_images(text, img_ids)
                                    tokens = split_chars(processed_text)
                                else:
                                    tokens = split_chars(text)
                                # Filter to keep only valid rhyme characters and image placeholders
                                filtered_tokens = filter_rhyme_chars(tokens)
                                append_preserve(current['Characters'], filtered_tokens)
                        else:
                            # orphan continuation row -> keep record but do not merge
                            pass
        # after pages: flush current group
        if current:
            # Collect image IDs from Characters
            image_ids_in_chars = []
            for ch in current['Characters']:
                if is_img_token(ch):
                    img_id = extract_image_id(ch)
                    if img_id and img_id not in image_ids_in_chars:
                        image_ids_in_chars.append(img_id)
            
            # Only include if Characters field has content and doesn't look like a poem
            if current['Characters']:
                chars_text = '、'.join(current['Characters'])
                
                # Exclude if it looks like a poem (not just rhyme characters)
                is_poem = False
                
                if re.search(r'[。．！？…]', chars_text):
                    is_poem = True
                elif chars_text.count('一曰') >= 2 or chars_text.count('二曰') >= 2:
                    is_poem = True
                elif len(chars_text) > 200:  # Very long suggests concatenated poem lines
                    is_poem = True
                
                if not is_poem:
                    rhyme_groups.append({
                        'RhymeSegment': current['RhymeSegment'],
                        'Characters': chars_text,
                        'OC_RhymeGroup': '、'.join(current['OC_RhymeGroup']),
                        'Tones': ' '.join(current['Tones']),
                        'Source': current['Source'],
                        'Page': current['Page'],
                        'TableID': current.get('TableID',''),
                        'image_refs': '|'.join(image_ids_in_chars)
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
            # Tone cell detection: must be PRIMARILY tone markers, not just contain one
            # Count tone-like content vs total content
            tone_chars = len(re.findall(r'[—–\-平上去入]', c))
            total_chars = len(re.sub(r'\s', '', c))  # Exclude whitespace
            is_tone_cell = (tone_chars > 0 and tone_chars / max(total_chars, 1) > 0.5)
            
            if is_tone_cell:
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
            summary['unresolved_rows'].append({'RowID': rowid, 'TableID': table_id, 'page': page, 'Characters': rhyme_raw_after, 'rhyme_tokens': rhyme_tokens})
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
                'Characters': rhyme_raw_after,
                'OC_RhymeGroup': oc_rhyme,
                'Tones': tones_raw,
                'rhyme_tokens': ' | '.join(rhyme_tokens),
                'tone_tokens': ' | '.join(tone_tokens),
                'rhyme_numbers': ' '.join(re.findall(r'\b(\d+)\b', raw_text)),
                'Source': source,
                'alignment_issue': '1' if alignment_issue else '',
                'notes': json.dumps(notes_obj, ensure_ascii=False),
                'image_refs': ','.join(image_refs)
            })
    # after row loop, produce sample snippets for prose drops
    prose_samples = summary.get('prose_dropped_rows', [])[:10]
    summary['prose_samples'] = prose_samples
    return ann, warnings, summary


def make_line_rows(rows):
    """Extract poem lines in long format: one CSV row per line.
    Parses line text, rhyme groups, and preserves image placeholders.
    """
    # First pass: collect all unique poems (Table + RhymeSegment combinations) and assign sequential IDs
    poems_seen = {}  # (table_id, rhyme_segment) -> sequential poem number
    poem_counter = 1
    
    for r in rows:
        table_id = r.get('table_id','')
        rhyme_segment = r.get('group_id', '')
        text = r.get('row_text','')
        
        if not text:
            continue
        
        # Apply same filters as before
        if '部' in text:
            continue
        if re.search(r'[—–\-]', text) and re.search(r'[平上去入]', text):
            continue
        if text.count('、') >= 3:
            continue
        if not re.search(r'[。．！？…]', text):
            continue
        
        total_cjk = len(re.findall(r'[\u4E00-\u9FFF]', text))
        if total_cjk < 6:
            continue
        
        # Register this poem if not seen
        poem_key = (table_id, rhyme_segment)
        if poem_key not in poems_seen:
            poems_seen[poem_key] = poem_counter
            poem_counter += 1
    
    # Second pass: actually extract lines
    line_rows = []
    
    for r in rows:
        page = r.get('page')
        table_id = r.get('table_id','')
        text = r.get('row_text','')
        if not text:
            continue
        
        # Filter 1: Exclude rhyme word lists
        if '部' in text:
            continue
        if re.search(r'[—–\-]', text) and re.search(r'[平上去入]', text):
            continue
        if text.count('、') >= 3:
            continue
        
        # Filter 2: Require sentence-ending punctuation
        if not re.search(r'[。．！？…]', text):
            continue
        
        # Filter 3: Minimum length
        total_cjk = len(re.findall(r'[\u4E00-\u9FFF]', text))
        if total_cjk < 6:
            continue
        
        # Extract source from row (already extracted during parsing)
        source = r.get('source', '')
        
        needs_images = row_needs_images(r.get('cells', [])) and r.get('has_images', False)
        image_refs = r.get('image_refs', []) if needs_images else []
        
        # Apply image replacement to cells if needed
        cells = r.get('cells', [])
        if image_refs:
            cells, _ = replace_images_in_cells(cells, image_refs)
        
        # Find the cell with the actual poem lines
        # Priority: cell with image placeholders > longest cell > first substantial cell
        poem_cell = None
        poem_cell_idx = -1
        
        # First pass: look for cells with image placeholders
        if image_refs:
            for idx, c in enumerate(cells[1:], start=1):
                if c and '⟦IMG:' in c:
                    poem_cell = c
                    poem_cell_idx = idx
                    break
        
        # Second pass: if no image placeholders found, take the longest substantial cell
        if not poem_cell:
            max_len = 0
            for idx, c in enumerate(cells[1:], start=1):
                if c and len(c.strip()) > max_len:
                    poem_cell = c
                    poem_cell_idx = idx
                    max_len = len(c.strip())
        
        if not poem_cell:
            continue
        
        # Split into individual lines
        raw_lines = [line.strip() for line in poem_cell.split('\n') if line.strip()]
        
        # Get poem ID (sequential number assigned in first pass)
        rhyme_segment = r.get('group_id', '')
        poem_key = (table_id, rhyme_segment)
        poem_num = poems_seen.get(poem_key, 0)
        poem_id = f"POEM_{poem_num:03d}"
        
        # Track rhyme groups encountered for rhyme ID assignment
        rhyme_group_map = {}  # Maps rhyme group to rhyme ID
        next_rhyme_id = 1
        
        # Parse each line
        line_order = 1
        for raw_line in raw_lines:
            # Parse format: "line_text,line_number （rhyme_group）" or just "line_text （rhyme_group）"
            
            # Extract rhyme group marker if present: （XXX）
            rhyme_group = ''
            rhyme_group_match = re.search(r'[（(]([^）)]+)[）)]', raw_line)
            if rhyme_group_match:
                rhyme_group = rhyme_group_match.group(1).strip()
                # Clean numbers and extra spaces from rhyme group
                rhyme_group = re.sub(r'\d+\s*', '', rhyme_group).strip()
            
            # Remove rhyme group marker from line text
            line_text = re.sub(r'\s*[（(][^）)]+[）)]\s*', '', raw_line)
            
            # Remove line numbers like "01-3" or "59-4" from line text
            # Try multiple patterns
            line_text = re.sub(r',\s*\d+-\d+\s*', '', line_text)  # ,01-3
            line_text = re.sub(r'\s+\d+-\d+\s*', ' ', line_text)  # space 01-3
            line_text = re.sub(r'\d+-\d+', '', line_text)      # 01-3 anywhere
            
            # Also remove standalone numbers like "024", "056" (but preserve image placeholders)
            # First mark image placeholders to protect them
            protected_imgs = []
            for match in re.finditer(r'⟦IMG:[^⟧]+⟧', line_text):
                protected_imgs.append(match.group())
            
            # Replace with temporary markers
            for i, img in enumerate(protected_imgs):
                line_text = line_text.replace(img, f'<<PROTECTED_{i}>>', 1)
            
            # Now remove numbers
            line_text = re.sub(r',\s*\d+', '', line_text)  # ,024
            line_text = re.sub(r'\d+\s*$', '', line_text)  # trailing 024
            line_text = re.sub(r'[。．]\s*\d+', '。', line_text)  # 。024
            
            # Restore protected images
            for i, img in enumerate(protected_imgs):
                line_text = line_text.replace(f'<<PROTECTED_{i}>>', img)
            
            # Clean up extra punctuation and whitespace
            line_text = line_text.strip().strip(',，。；：')
            
            # Skip if this is just a rhyme group marker with no actual line content
            if not line_text:
                continue
            
            # Skip if this looks like it's just metadata
            if line_text.startswith('《') or re.match(r'^\d+$', line_text):
                continue
            
            # Assign rhyme ID based on rhyme group
            rhyme_id = ''
            if rhyme_group:
                if rhyme_group not in rhyme_group_map:
                    rhyme_group_map[rhyme_group] = next_rhyme_id
                    next_rhyme_id += 1
                rhyme_id = str(rhyme_group_map[rhyme_group])
            
            line_rows.append({
                'POEM': poem_id,
                'LINE_ORDER': line_order,
                'LINE': line_text,
                'RHYME_GROUP': rhyme_group,
                'RHYME_ID': rhyme_id,
                'SOURCE': source,
                'PAGE': str(page) if page else '',
                'TABLE_ID': table_id,
                'RHYME_SEGMENT': str(rhyme_segment) if rhyme_segment else ''
            })
            
            line_order += 1
    
    return line_rows


def split_verse_lines_by_slip_id(text):
    """
    Split a text chunk into multiple verse lines if it contains slip IDs.
    
    Handles TWO formats:
    1. XX-Y format: "01-1", "47-2", "1-2背" (used in 睡虎地, 嶽麓書院, 周家臺)
    2. Simple numbers: "351", "252", "41" (used in 放馬灘, 北大)
    
    Rule: Each verse line MUST contain exactly ONE slip ID at the end.
    If text contains multiple slip IDs, split at each boundary.
    
    Returns: List of (line_text, slip_id) tuples
    """
    # --- Pattern 1: XX-Y format (e.g., 01-1, 47-2, 1-2背) ---
    SLIP_ID_PATTERN = re.compile(
        r'(?P<slip>\d{1,3}-\d{1,2}(?:背)?)(?P<fn>\[\d+\])?(?=[\s，。；：\u4E00-\u9FFF]|$)'
    )
    
    # Hard rejection: if followed by 頁, 年, % (false positives like "110-111頁")
    REJECT_PATTERN = re.compile(r'\d{1,3}-\d{1,2}[頁年%]')
    
    if REJECT_PATTERN.search(text):
        return []
    
    # Try XX-Y format first
    matches = list(SLIP_ID_PATTERN.finditer(text))
    
    if matches:
        # Found XX-Y slip IDs
        if len(matches) == 1:
            match = matches[0]
            line_text = text[:match.start()].strip()
            slip_id = match.group('slip')
            footnote = match.group('fn') or ''
            return [(f'{line_text}{footnote}', slip_id)]
        
        # Multiple slip IDs - split into multiple verse lines
        lines = []
        for i, match in enumerate(matches):
            slip_id = match.group('slip')
            
            if i == 0:
                line_text = text[:match.start()].strip()
            else:
                prev_end = matches[i-1].end()
                line_text = text[prev_end:match.start()].strip()
            footnote = match.group('fn') or ''
            
            line_text = line_text.lstrip('，。；：、 ')
            
            if line_text:
                lines.append((f'{line_text}{footnote}', slip_id))
        
        return lines
    
    # --- Pattern 2: Simple numbers at end of line (e.g., 351, 252, 41) ---
    # Must follow CJK text or punctuation, be 2-3 digits, and be at line end or followed by punctuation
    SIMPLE_NUM_PATTERN = re.compile(
        r'(?P<slip>\d{2,3})(?P<fn>\[\d+\])?(?=[\s，。；：]|$)'
    )
    
    # Skip lines that look like footnotes, bibliography, or commentary
    skip_patterns = [
        r'^\[\d+\]',           # Footnote lines [1] ...
        r'^【(?:註釋|注釋|注释)】',  # Note-section headers merged with note text
        r'^\d+年',             # Year references
        r'頁[。，]',           # Page references
        r'第\d+期',            # Journal issue numbers
        r'簡\d+-\d+作',        # References to other slips
        r'（\d{4}[年）]',      # Years in parentheses
        r'^整理者',            # Editor commentary
        r'^今按',              # Author commentary
        r'^關於',              # "Regarding..." commentary
        r'[①②③④⑤⑥⑦⑧⑨⑩]',  # Circled footnote markers
        r'《[^》]+》[,，].*?http',  # Bibliography with URL
        r'[:：]《[^》]+》',    # Citation format: Name:《Title》
        r'復旦大學',           # Specific university names (bibliography)
        r'出土文獻',           # Documentary source references
        r'研究中心',           # Research center references
    ]
    
    for pattern in skip_patterns:
        if re.search(pattern, text):
            return []
    
    # Look for simple numbers at end
    # The line should have CJK content before the number
    if not re.search(r'[\u4E00-\u9FFF]', text):
        return []
    
    # Find numbers at or near end of line (may have punctuation after)
    match = re.search(r'(.+?)(\d{2,3})(\[\d+\])?\s*[，。；：]?\s*$', text)
    if match:
        line_text = match.group(1).strip()
        slip_id = match.group(2)
        footnote = match.group(3) or ''
        
        # Line must have meaningful CJK content (single-character lines are valid in this corpus)
        cjk_count = len(re.findall(r'[\u4E00-\u9FFF]', line_text))
        if cjk_count >= 1:
            # Avoid matching years (e.g., "2010年" -> "2010")
            if not re.search(r'\d{4}$', line_text):
                return [(f"{line_text.rstrip('，。；：、 ')}{footnote}", slip_id)]
    
    return []


def find_terminal_token(line_text):
    """
    Find the terminal (rhyme) token by scanning from end.
    
    Rule D: If line ends with X（Y）, tag after X, match using X or Y.
    
    Skip:
    - Footnote markers like [5]
    - Punctuation "，。；："
    - Trailing parenthetical gloss "（…）" comes AFTER the base character
    
    Return: (token, gloss_char, insert_position)
    """
    # Remove trailing footnotes and punctuation
    text = line_text
    text = re.sub(r'\[\d+\]\s*$', '', text)  # Remove [5]
    text = text.rstrip('，。；： \t')
    
    # Check for pattern X（Y） at end
    # Match: last_char（gloss）
    parenthetical_match = re.search(r'([\u4E00-\u9FFF⟦])（([^）]+)）\s*$', text)
    
    if parenthetical_match:
        # Found X（Y）
        base_char_or_img = parenthetical_match.group(1)
        gloss_text = parenthetical_match.group(2)
        
        # Extract gloss character
        gloss_char = None
        if base_char_or_img.startswith('⟦'):
            # It's an image placeholder like ⟦IMG:...⟧（Y）
            # Need to extract the full image placeholder
            img_match = re.search(r'(⟦IMG:[^⟧]+⟧)（([^）]+)）\s*$', text)
            if img_match:
                token = img_match.group(1)
                gloss_text = img_match.group(2)
                cjk_in_gloss = re.findall(r'[\u4E00-\u9FFF]', gloss_text)
                if cjk_in_gloss:
                    gloss_char = cjk_in_gloss[0]
                # Insert position is after the image token, before （
                insert_pos = img_match.start(1) + len(token)
                return token, gloss_char, insert_pos
        else:
            # Regular character
            token = base_char_or_img
            cjk_in_gloss = re.findall(r'[\u4E00-\u9FFF]', gloss_text)
            if cjk_in_gloss:
                gloss_char = cjk_in_gloss[0]
            # Insert position is after the base character
            insert_pos = parenthetical_match.start(1) + 1
            return token, gloss_char, insert_pos
    
    # No X（Y） pattern - find last token normally
    # Remove any remaining parentheticals
    text = re.sub(r'（[^）]+）\s*$', '', text)
    
    # Find last image token
    img_match = re.search(r'(⟦IMG:[^⟧]+⟧)[^⟦]*$', text)
    if img_match:
        token = img_match.group(1)
        insert_pos = img_match.end()
        return token, None, insert_pos
    
    # Define common Classical Chinese particles
    particles = '之也焉矣乎哉耶邪與歟'
    
    # Find all CJK characters
    cjk_chars = list(re.finditer(r'[\u4E00-\u9FFF]', text))
    if cjk_chars:
        last_match = cjk_chars[-1]
        token = last_match.group()
        insert_pos = last_match.end()
        
        # If last character is a particle, try second-to-last
        if token in particles and len(cjk_chars) >= 2:
            second_last_match = cjk_chars[-2]
            second_token = second_last_match.group()
            # Return second-to-last as primary token, last as potential alternate
            # Insert position is after the second-to-last character
            return second_token, token, second_last_match.end()
        
        return token, None, insert_pos
    
    return None, None, len(text)


def find_rhyme_set(token, gloss_char, rhyme_sets):
    """
    Find which rhyme set this token belongs to.
    
    Rule D: Try matching token first, then gloss_char.
    """
    if not rhyme_sets:
        return None
    
    # Try exact token match first
    for rset in rhyme_sets:
        if token in rset['words']:
            return rset
    
    # Try gloss character if available
    if gloss_char:
        for rset in rhyme_sets:
            if gloss_char in rset['words']:
                return rset
    
    return None


def parse_rhyme_info(rhyme_info_raw):
    """
    Parse rhyme info into rhyme sets with IDs.
    
    Input: "職蒸部通韻——塞、力、能；真元部合韻——身、願。"
    Output: [
        {'id': 'a', 'label': '職蒸部通韻', 'words': ['塞', '力', '能']},
        {'id': 'b', 'label': '真元部合韻', 'words': ['身', '願']},
    ]
    """
    if not rhyme_info_raw:
        return []
    
    rhyme_sets = []
    set_ids = 'abcdefghijklmnopqrstuvwxyz'
    set_idx = 0
    
    # Split on sentence delimiters
    clauses = re.split(r'[；。]', rhyme_info_raw)
    
    for clause in clauses:
        clause = clause.strip()
        if not clause or '——' not in clause:
            continue
        
        # Split into label and wordlist
        parts = clause.split('——', 1)
        if len(parts) != 2:
            continue
        
        label = parts[0].strip()
        wordlist = parts[1].strip()
        
        # Split wordlist on enumeration comma
        words = []
        for word in re.split(r'[、,，;；\s]+', wordlist):
            word = word.strip()
            # Keep only actual characters and image placeholders
            if word and (re.search(r'[\u4E00-\u9FFF]', word) or '⟦IMG:' in word):
                words.append(word)

        if words and set_idx < len(set_ids):
            rhyme_sets.append({
                'id': set_ids[set_idx],
                'label': label,
                'words': words,
                'category': normalize_rhyme_category(label),
                'rhyme_type': extract_rhyme_type(label) or infer_single_group_rhyme_type(label),
                'tone_map': {}
            })
            set_idx += 1

    return rhyme_sets


def enrich_rhyme_sets_with_tones(rhyme_sets, text_name, tone_lookup):
    """Attach per-word tone data and safe single-group rhyme types from annotations."""
    if not rhyme_sets or not text_name or not tone_lookup:
        return rhyme_sets

    candidates = tone_lookup.get(text_name, [])
    if not candidates:
        return rhyme_sets

    for rset in rhyme_sets:
        words = rset.get('words', [])
        word_set = set(words)
        if not word_set:
            continue

        best_candidate = None
        best_score = 0
        category = rset.get('category', '')

        for candidate in candidates:
            overlap = len(word_set & candidate.get('words', set()))
            if overlap == 0:
                continue

            score = overlap * 10
            candidate_group = candidate.get('group', '')
            if category and candidate_group:
                if category == candidate_group:
                    score += 5
                elif category in candidate_group or candidate_group in category:
                    score += 2

            if score > best_score:
                best_candidate = candidate
                best_score = score

        if not best_candidate:
            continue

        tone_map = {}

        # First prefer exact word matches.
        for word in words:
            tone = best_candidate.get('word_tones', {}).get(word)
            if tone and word not in tone_map:
                tone_map[word] = tone

        # If exact token matching misses image placeholders or alternate glyph forms,
        # fall back to positional alignment when counts line up.
        ordered_words = best_candidate.get('ordered_words', [])
        ordered_tones = best_candidate.get('ordered_tones', [])
        if len(words) == len(ordered_words):
            for idx, word in enumerate(words):
                if word not in tone_map and idx < len(ordered_tones):
                    tone = ordered_tones[idx]
                    if tone:
                        tone_map[word] = tone

        rset['tone_map'] = tone_map
        if not rset.get('rhyme_type'):
            candidate_type = best_candidate.get('rhyme_type', '')
            candidate_group = best_candidate.get('group', '')
            if candidate_type and (not category or category == candidate_group):
                rset['rhyme_type'] = candidate_type

    return rhyme_sets


def build_rhyme_marker(rset, matched_word=None, gloss_char=None, include_category=True):
    """Build inline rhyme markers like [a@魚歌-平-合]."""
    set_id = rset.get('id', '')
    if not set_id:
        return ''

    if not include_category:
        return f'[{set_id}]'

    category = rset.get('category') or normalize_rhyme_category(rset.get('label', ''))
    tone_map = rset.get('tone_map', {}) or {}
    tone = ''
    if matched_word:
        tone = tone_map.get(matched_word, '')
    if not tone and gloss_char:
        tone = tone_map.get(gloss_char, '')
    rhyme_type = rset.get('rhyme_type') or extract_rhyme_type(rset.get('label', ''))

    parts = [part for part in (category, tone, rhyme_type) if part]
    if not parts:
        return f'[{set_id}]'

    return f'[{set_id}@{"-".join(parts)}]'


def annotate_line_with_rhyme(line_text, rhyme_sets, slip_id='', include_category=True):
    """
    Annotate a line with rhyme set marker [a]/[b]/etc.
    
    If include_category=True (new default), uses Suzuki-style format: [a@陽部]
    Otherwise uses simple format: [a]
    
    This function now handles BOTH terminal and non-terminal rhyme words.
    For lines like "欲富大（太）甚，貧不可得", it will annotate both 甚 and 得
    if both are rhyme words.
    
    Also handles gloss characters in parentheticals like ⟦IMG:X⟧（Y） where Y is the rhyme word.
    
    Returns: (annotated_text, rhyme_set_id, rhyme_word, rhyme_label, missed_tag_reason)
    """
    if not rhyme_sets:
        return line_text, '', '', '', ''
    
    # Build a map of all rhyme words to their sets
    word_to_set = {}
    for rset in rhyme_sets:
        for word in rset['words']:
            if word not in word_to_set:  # First match wins
                word_to_set[word] = rset
    
    # Find ALL occurrences of rhyme words in the line and annotate them
    # Process from right to left to avoid index shifting
    annotated = line_text
    all_matches = []
    
    for word, rset in word_to_set.items():
        # Skip image tokens for direct matching (handle via gloss)
        if '⟦IMG:' in word:
            continue
        
        # Find all occurrences of this word in the line
        for match in re.finditer(re.escape(word), line_text):
            pos = match.end()
            # Check if next char is already a rhyme marker
            if pos < len(line_text) and line_text[pos] == '[':
                continue
            
            # Check if this is inside a parenthetical (gloss)
            prefix = line_text[:match.start()]
            paren_opens = prefix.count('（') - prefix.count('）')
            
            if paren_opens > 0:
                # This word is inside a parenthetical - treat it as a gloss match
                # Find the position to insert marker (after the closing ）)
                suffix = line_text[match.end():]
                close_paren_match = re.match(r'[^）]*）', suffix)
                if close_paren_match:
                    insert_end = match.end() + close_paren_match.end()
                    all_matches.append({
                        'word': word,
                        'start': match.start(),
                        'end': insert_end,
                        'rset': rset,
                        'is_gloss': True,
                        'gloss_char': word
                    })
            else:
                all_matches.append({
                    'word': word,
                    'start': match.start(),
                    'end': match.end(),
                    'rset': rset,
                    'is_gloss': False,
                    'gloss_char': None
                })
    
    # Also check for image tokens that are rhyme words (exact match)
    for word, rset in word_to_set.items():
        if '⟦IMG:' in word:
            for match in re.finditer(re.escape(word), line_text):
                pos = match.end()
                if pos < len(line_text) and line_text[pos] == '[':
                    continue
                # Check for following gloss: ⟦IMG:X⟧（Y）
                suffix = line_text[match.end():]
                gloss_match = re.match(r'（[^）]+）', suffix)
                if gloss_match:
                    # Include gloss in the match so marker goes after ）
                    all_matches.append({
                        'word': word,
                        'start': match.start(),
                        'end': match.end() + gloss_match.end(),
                        'rset': rset,
                        'is_gloss': False,
                        'gloss_char': None
                    })
                else:
                    all_matches.append({
                        'word': word,
                        'start': match.start(),
                        'end': match.end(),
                        'rset': rset,
                        'is_gloss': False,
                        'gloss_char': None
                    })
    
    # Special: Look for patterns ⟦IMG:X⟧（Y） where Y might be a rhyme word
    # even if the image in rhyme_info is different
    img_gloss_pattern = re.compile(r'⟦IMG:[^⟧]+⟧（([^）]+)）')
    for match in img_gloss_pattern.finditer(line_text):
        gloss_text = match.group(1)
        # Extract CJK chars from gloss
        for ch in re.findall(r'[\u4E00-\u9FFF]', gloss_text):
            if ch in word_to_set:
                rset = word_to_set[ch]
                # Check if not already matched
                pos_already_matched = any(m['start'] <= match.start() < m['end'] for m in all_matches)
                if not pos_already_matched:
                    all_matches.append({
                        'word': ch,
                        'start': match.start(),
                        'end': match.end(),
                        'rset': rset,
                        'is_gloss': True,
                        'gloss_char': ch
                    })
    
    # Special case: if rhyme_info contains image tokens, allow near-terminal image forms
    # to stand for the rhyme even when the image ID differs across extractions.
    rhyme_sets_with_images = [rset for rset in rhyme_sets if any('⟦IMG:' in w for w in rset['words'])]
    if rhyme_sets_with_images:
        trailing_img_pattern = re.compile(r'⟦IMG:[^⟧]+⟧(?:（([^）]+)）)?')
        trailing_img_suffix = re.compile(r'(?:\[\d+\]|\d+|[之也者矣兮乎焉耳]|[，。；：、\s])*$')
        for match in trailing_img_pattern.finditer(line_text):
            suffix = line_text[match.end():]
            if not trailing_img_suffix.fullmatch(suffix):
                continue
            pos_already_matched = any(m['start'] <= match.start() < m['end'] for m in all_matches)
            if pos_already_matched:
                continue

            gloss_char = ''
            if match.group(1):
                cjk_chars = re.findall(r'[\u4E00-\u9FFF]', match.group(1))
                if cjk_chars:
                    gloss_char = cjk_chars[0]

            all_matches.append({
                'word': match.group(0),
                'start': match.start(),
                'end': match.end(),
                'rset': rhyme_sets_with_images[0],
                'is_gloss': bool(gloss_char),
                'gloss_char': gloss_char
            })
    
    # Sort matches by position (descending) to annotate from right to left
    all_matches.sort(key=lambda m: m['start'], reverse=True)
    
    # Remove overlapping matches (keep rightmost)
    filtered_matches = []
    last_start = len(line_text)
    for match in all_matches:
        if match['end'] <= last_start:
            filtered_matches.append(match)
            last_start = match['start']
    
    # Apply annotations from right to left
    primary_set_id = ''
    primary_word = ''
    primary_label = ''
    
    for match in filtered_matches:
        rset = match['rset']
        set_id = rset['id']
        label = rset.get('label', '')
        marker = build_rhyme_marker(
            rset,
            matched_word=match.get('word'),
            gloss_char=match.get('gloss_char'),
            include_category=include_category,
        )

        # Insert marker after the word
        annotated = annotated[:match['end']] + marker + annotated[match['end']:]
        
        # Track the rightmost (primary) annotation
        if not primary_set_id:
            primary_set_id = set_id
            primary_word = match['word']
            primary_label = label
    
    # If no matches, try original terminal-only logic as fallback
    if not filtered_matches:
        token, gloss_char, insert_pos = find_terminal_token(line_text)
        
        if not token:
            return line_text, '', '', '', ''
        
        rhyme_set = find_rhyme_set(token, gloss_char, rhyme_sets)
        
        if not rhyme_set:
            missed_reason = ''
            for rset in rhyme_sets:
                if token in rset['words'] or (gloss_char and gloss_char in rset['words']):
                    missed_reason = f"Token '{token}' found in rhyme set '{rset['label']}' but not tagged"
                    break
            return line_text, '', token, '', missed_reason
        
        label = rhyme_set.get('label', '')
        marker = build_rhyme_marker(
            rhyme_set,
            matched_word=token,
            gloss_char=gloss_char,
            include_category=include_category,
        )

        annotated = line_text[:insert_pos] + marker + line_text[insert_pos:]
        return annotated, rhyme_set['id'], token, label, ''
    
    return annotated, primary_set_id, primary_word, primary_label, ''


def parse_slip_id_key(slip_id):
    """Parse slip ids like 37-2 or 326 into sortable tuples."""
    if not slip_id:
        return None

    slip_id = slip_id.replace('簡', '').strip()

    match = re.match(r'^(\d+)-(\d+)(背)?$', slip_id)
    if match:
        return ('compound', int(match.group(1)), int(match.group(2)), 1 if match.group(3) else 0)

    match = re.match(r'^(\d+)$', slip_id)
    if match:
        return ('simple', int(match.group(1)))

    return None


def parse_slip_range_bounds(slip_range):
    """Return (start_key, end_key) parsed from a slip range string."""
    if not slip_range:
        return None, None

    slip_range = slip_range.replace('簡', '').strip()
    if '至' in slip_range:
        start_text, end_text = [part.strip() for part in slip_range.split('至', 1)]
    else:
        start_text = end_text = slip_range

    return parse_slip_id_key(start_text), parse_slip_id_key(end_text)


def slip_ranges_are_adjacent(left_range, right_range):
    """Check whether two slip ranges are directly adjacent."""
    _, left_end = parse_slip_range_bounds(left_range)
    right_start, _ = parse_slip_range_bounds(right_range)

    if not left_end or not right_start:
        return False

    if left_end[0] != right_start[0]:
        return False

    if left_end[0] == 'compound':
        return left_end[2] == right_start[2] and left_end[3] == right_start[3] and right_start[1] == left_end[1] + 1

    return right_start[1] == left_end[1] + 1


def segment_looks_like_intro(segment):
    """Detect short introductory fragments that belong with the next verse segment."""
    texts = [line.get('text', '').strip() for line in segment.get('lines', []) if line.get('text', '').strip()]
    if not texts or len(texts) > 3:
        return False

    intro_patterns = [
        r'曰[:：]?$',
        r'祝曰[:：]?$',
        r'禹步.*曰[:：]?$',
        r'見車[，,].*曰[:：]?$',
    ]

    return any(re.search(pattern, text) for text in texts for pattern in intro_patterns)


def segment_looks_like_page_break_continuation(left, right):
    """Detect verse fragments stranded on the previous page before rhyme info resumes."""
    if not left.get('lines') or not right.get('lines'):
        return False

    left_page = left.get('start_page') or 0
    right_page = right.get('start_page') or 0
    if right_page - left_page != 1:
        return False

    left_lines = left.get('lines', [])
    if len(left_lines) > 9:
        return False

    # Prefer cases where the fragment looks like a genuine continuation,
    # not a fresh section opening or commentary.
    left_texts = [line.get('text', '').strip() for line in left_lines if line.get('text', '').strip()]
    if not left_texts:
        return False

    if any(is_footnote_or_commentary(text) for text in left_texts):
        return False

    return True


def merge_segment_pair(left, right):
    """Merge two adjacent segments that belong to the same poem."""
    merged = dict(left)
    merged['lines'] = left.get('lines', []) + right.get('lines', [])
    merged['rhyme_info_raw'] = left.get('rhyme_info_raw') or right.get('rhyme_info_raw', '')
    merged['notes'] = left.get('notes', []) + [note for note in right.get('notes', []) if note not in left.get('notes', [])]
    merged['img_count'] = sum(line.get('text', '').count('⟦IMG:') for line in merged['lines'])
    merged['start_page'] = min(left.get('start_page') or 0, right.get('start_page') or 0)

    if not merged.get('slip_range'):
        merged['slip_range'] = right.get('slip_range', '')
    elif right.get('slip_range'):
        left_start, left_end = parse_slip_range_bounds(left.get('slip_range', ''))
        right_start, right_end = parse_slip_range_bounds(right.get('slip_range', ''))
        if left_start and right_end:
            merged['slip_range'] = f"簡{left.get('slip_range', '').replace('簡', '').split('至', 1)[0]}至{right.get('slip_range', '').replace('簡', '').split('至', 1)[-1]}"

    right_title = right.get('text_name', '')
    left_title = left.get('text_name', '')
    if right_title and (not left_title or len(right_title) > len(left_title)):
        merged['text_name'] = right_title

    return merged


def merge_related_segments(segments):
    """Merge adjacent segments that were incorrectly split between intro/verse/rhyme states."""
    if not segments:
        return segments

    merged = []
    i = 0
    while i < len(segments):
        current = segments[i]
        if i + 1 < len(segments):
            nxt = segments[i + 1]
            same_context = (
                current.get('collection') == nxt.get('collection') and
                current.get('section') == nxt.get('section')
            )
            same_title = (
                current.get('text_name') and
                current.get('text_name') == nxt.get('text_name')
            )
            current_needs_rhyme = not current.get('rhyme_info_raw')
            next_has_rhyme = bool(nxt.get('rhyme_info_raw'))
            contiguous_slips = slip_ranges_are_adjacent(current.get('slip_range', ''), nxt.get('slip_range', ''))
            intro_fragment = segment_looks_like_intro(current)
            page_break_fragment = segment_looks_like_page_break_continuation(current, nxt)

            if same_context and current_needs_rhyme and next_has_rhyme and (
                contiguous_slips or
                (same_title and intro_fragment) or
                (same_title and page_break_fragment)
            ):
                merged.append(merge_segment_pair(current, nxt))
                i += 2
                continue

        merged.append(current)
        i += 1

    for idx, segment in enumerate(merged, 100):
        segment['poem_id'] = f'POEM_{idx:03d}'

    return merged


def extract_chapter2_poems(pdf_path, start_page=21, end_page=141, tone_lookup=None):
    """
    Extract poems from Chapter 2 using layout-based parsing with state machine.
    
    State A: Waiting for 【用韻情況】
    State B: Collecting rhyme info, then verse lines
    
    Builds lines from layout tokens (chars + small images).
    """
    import os
    from pathlib import Path
    
    # Ensure extracted_images directory exists
    img_dir = Path('extracted_images')
    img_dir.mkdir(exist_ok=True)
    
    segments = []
    segment_counter = 100
    
    # State machine with backward rhyme attachment
    state = 'WAITING_VERSE'  # or 'COLLECTING_VERSE' or 'COLLECTING_RHYME_INFO' or 'IN_NOTES'
    current_verse_lines = []
    current_text_name = None
    pending_local_title = None
    current_segment_title = None
    current_start_page = None
    pending_rhyme_target_idx = None  # Index of segment waiting for rhyme info
    
    # New: Track collection/excavation site and section numbers
    current_collection = None  # e.g., "睡虎地秦墓竹簡"
    current_section = None     # e.g., "2.1.1"
    current_notes = []         # Footnotes for current segment
    
    # Collection mapping from section number prefix
    COLLECTION_MAP = {
        '2.1': '睡虎地秦墓竹簡',
        '2.2': '放馬灘秦墓簡牘',
        '2.3': '王家臺秦墓竹簡',
        '2.4': '周家臺秦墓簡牘',
        '2.5': '嶽麓書院藏秦簡',
        '2.6': '北京大學藏秦簡牘',
    }
    
    def flush_segment(attach_rhyme_info_lines=None):
        """Flush accumulated verse lines as a segment."""
        nonlocal segment_counter, current_verse_lines, current_text_name, pending_local_title, current_segment_title
        nonlocal current_start_page, pending_rhyme_target_idx, current_notes
        
        if not current_verse_lines:
            return
        
        # Don't attach rhyme info here - it will come later
        rhyme_info_raw = ''
        if attach_rhyme_info_lines:
            rhyme_info_raw = '\n'.join(attach_rhyme_info_lines).strip()
        
        img_count = sum(line['text'].count('⟦IMG:') for line in current_verse_lines)
        
        # Calculate slip range from verse lines
        slip_ids = [line.get('slip_id', '') for line in current_verse_lines if line.get('slip_id')]
        slip_range = ''
        if slip_ids:
            first_slip = slip_ids[0]
            last_slip = slip_ids[-1]
            if first_slip == last_slip:
                slip_range = f"簡{first_slip}"
            else:
                slip_range = f"簡{first_slip}至{last_slip}"
        
        segment_idx = len(segments)
        segments.append({
            'poem_id': f"POEM_{segment_counter:03d}",
            'text_name': current_segment_title or pending_local_title or current_text_name or '',
            'collection': current_collection or '',
            'section': current_section or '',
            'slip_range': slip_range,
            'start_page': current_start_page or 0,
            'lines': current_verse_lines[:],
            'rhyme_info_raw': rhyme_info_raw,
            'notes': current_notes[:],  # Copy current notes
            'img_count': img_count
        })
        
        # This segment is pending rhyme info
        pending_rhyme_target_idx = segment_idx
        
        segment_counter += 1
        current_verse_lines = []
        current_start_page = None
        current_notes = []  # Clear notes for next segment
        pending_local_title = None
        current_segment_title = None
    
    def build_line_from_layout(page, page_num):
        """Build visual lines from layout objects (chars + small images)."""
        chars = page.chars
        images = getattr(page, 'images', [])
        
        if not chars:
            return []
        
        # Group chars by y-position (top coordinate)
        Y_TOLERANCE = 3
        lines_by_y = {}
        
        for ch in chars:
            y = round(ch['top'] / Y_TOLERANCE) * Y_TOLERANCE
            if y not in lines_by_y:
                lines_by_y[y] = {'chars': [], 'images': []}
            lines_by_y[y]['chars'].append(ch)
        
        # Add small images (≤14pt width/height) to appropriate y-bands
        for idx, img in enumerate(images):
            width = img.get('width', 0)
            height = img.get('height', 0)
            if width <= 14 and height <= 14:
                img_y = img.get('top', 0)
                y = round(img_y / Y_TOLERANCE) * Y_TOLERANCE
                
                # Find closest y-band
                if y not in lines_by_y:
                    closest_y = min(lines_by_y.keys(), key=lambda k: abs(k - y), default=y)
                    if abs(closest_y - y) < 10:
                        y = closest_y
                
                if y in lines_by_y:
                    img_id = f"IMG_p{page_num:03d}_{idx+1:03d}"
                    lines_by_y[y]['images'].append({
                        'x0': img.get('x0', 0),
                        'img_id': img_id,
                        'img_obj': img
                    })
        
        # Build text lines by merging chars and images by x-position
        visual_lines = []
        for y in sorted(lines_by_y.keys()):
            band = lines_by_y[y]
            
            # Merge chars and image placeholders
            tokens = []
            for ch in band['chars']:
                tokens.append({'x0': ch['x0'], 'text': ch['text'], 'type': 'char'})
            for img_data in band['images']:
                tokens.append({'x0': img_data['x0'], 'text': f"⟦IMG:{img_data['img_id']}⟧", 
                              'type': 'img', 'img_obj': img_data['img_obj'], 'img_id': img_data['img_id']})
            
            # Sort by x-position
            tokens.sort(key=lambda t: t['x0'])
            
            # Build line text
            line_text = ''.join(t['text'] for t in tokens)
            
            # Save image crops
            for t in tokens:
                if t['type'] == 'img':
                    try:
                        img_obj = t['img_obj']
                        img_crop = page.crop((img_obj['x0'], img_obj['top'], 
                                             img_obj['x0'] + img_obj['width'],
                                             img_obj['top'] + img_obj['height']))
                        img_path = img_dir / f"{t['img_id']}.png"
                        img_crop.to_image(resolution=150).save(str(img_path))
                    except Exception:
                        pass
            
            visual_lines.append(line_text.strip())
        
        return visual_lines
    
    # Verse line pattern: ends with slip-id like 02-1, 17-3, 1-2背, optionally [5]
    VERSE_PATTERN = re.compile(r'^(.+?)[,，]?\s*(\d{1,3}-\d+[^,，。；\s]*?)(\[\d+\])?\s*$')
    
    with pdfplumber.open(pdf_path) as pdf:
        section_title_map = {}
        for page_num in range(start_page, min(end_page, len(pdf.pages) + 1)):
            text_lines = (pdf.pages[page_num - 1].extract_text() or '').splitlines()
            for raw_line in text_lines:
                raw_line = raw_line.strip()
                header_match = re.match(r'^(\d+(?:\.\d+)+)《([^》]+)》', raw_line)
                if header_match:
                    section_title_map[header_match.group(1)] = header_match.group(2).strip()

        for page_num in range(start_page, min(end_page, len(pdf.pages) + 1)):
            page = pdf.pages[page_num - 1]
            
            visual_lines = build_line_from_layout(page, page_num)
            
            for line in visual_lines:
                if not line:
                    continue
                
                # Detect section headers (collection and text name)
                section_match = re.match(r'^(\d+(?:\.\d+)+)', line)
                if section_match:
                    section_num = section_match.group(1)
                    current_section = section_num
                    
                    # A new section ends any previous state
                    if state == 'COLLECTING_VERSE' and current_verse_lines:
                        flush_segment()
                        state = 'WAITING_VERSE'
                    elif state in ('COLLECTING_RHYME_INFO', 'IN_NOTES'):
                        # Flush any pending rhyme info
                        if state == 'COLLECTING_RHYME_INFO' and pending_rhyme_target_idx is not None:
                            if pending_rhyme_target_idx < len(segments):
                                rhyme_info_raw = '\n'.join(current_rhyme_info_lines).strip()
                                segments[pending_rhyme_target_idx]['rhyme_info_raw'] = rhyme_info_raw
                                pending_rhyme_target_idx = None
                            current_rhyme_info_lines = []
                        state = 'WAITING_VERSE'
                    pending_local_title = None
                    
                    # Update collection based on major section number (e.g., 2.1 -> 睡虎地)
                    major_section = '.'.join(section_num.split('.')[:2])
                    if major_section in COLLECTION_MAP:
                        current_collection = COLLECTION_MAP[major_section]
                    
                    # Extract text name if present
                    text_match = re.search(r'《([^》]+)》', line)
                    if section_num in section_title_map:
                        current_text_name = section_title_map[section_num]
                    elif text_match:
                        current_text_name = text_match.group(1)
                    else:
                        header_text = line[section_match.end():].strip()
                        current_text_name = header_text or None
                    continue
                
                # Skip page headers/footers
                if any(skip in line for skip in ['秦簡牘韻文整理與研究', '济南大学硕士学位论文']):
                    continue

                if state != 'IN_NOTES':
                    local_title, remainder = extract_leading_title(line)
                    if local_title:
                        if state == 'COLLECTING_VERSE' and current_verse_lines:
                            flush_segment()
                            state = 'WAITING_VERSE'
                        pending_local_title = local_title
                        line = remainder
                        if not line:
                            continue

                # State transitions
                if line == '【用韻情況】':
                    # Flush current verse segment before collecting rhyme info
                    if state == 'COLLECTING_VERSE' and current_verse_lines:
                        flush_segment()
                    # Now collect rhyme info - it will be attached to the segment we just flushed
                    state = 'COLLECTING_RHYME_INFO'
                    current_rhyme_info_lines = []
                    continue

                note_payload = strip_note_section_header(line)
                if note_payload is not None:
                    # Start notes section - collect footnotes for PRECEDING segment
                    if state == 'COLLECTING_RHYME_INFO':
                        # Attach collected rhyme info to pending segment
                        if pending_rhyme_target_idx is not None and pending_rhyme_target_idx < len(segments):
                            rhyme_info_raw = '\n'.join(current_rhyme_info_lines).strip()
                            segments[pending_rhyme_target_idx]['rhyme_info_raw'] = rhyme_info_raw
                        current_rhyme_info_lines = []
                    elif state == 'COLLECTING_VERSE':
                        # End current verse segment
                        flush_segment()
                    # Notes will be attached to pending_rhyme_target_idx (the segment we just finished)
                    state = 'IN_NOTES'
                    line = note_payload
                    if not line:
                        continue
                
                # Collect footnotes in notes section - attach to PRECEDING segment
                if state == 'IN_NOTES':
                    # Footnotes start with [n] pattern
                    if re.match(r'^\[\d+\]', line.strip()):
                        note_match = re.match(r'^\[(\d+)\]', line.strip())
                        # Attach this note directly to the pending segment
                        if pending_rhyme_target_idx is not None and pending_rhyme_target_idx < len(segments):
                            if note_match:
                                restore_terminal_note_marker(segments[pending_rhyme_target_idx], note_match.group(1))
                            segments[pending_rhyme_target_idx]['notes'].append(line.strip())
                        else:
                            # Fallback to current_notes for next segment (shouldn't happen normally)
                            current_notes.append(line.strip())
                    elif pending_rhyme_target_idx is not None and pending_rhyme_target_idx < len(segments):
                        # Continuation of previous footnote
                        seg_notes = segments[pending_rhyme_target_idx]['notes']
                        if seg_notes and not line.startswith('①') and not line.startswith('②'):
                            if not any(skip in line for skip in ['http:', 'www.', '年第', '頁。']):
                                seg_notes[-1] += ' ' + line.strip()
                    continue
                
                # Handle COLLECTING_RHYME_INFO state first - we need to capture rhyme info lines
                # which would otherwise be filtered by is_footnote_or_commentary
                if state == 'COLLECTING_RHYME_INFO':
                    if is_footnote_or_commentary(line):
                        current_rhyme_info_lines.append(line)
                        continue

                    # Check if this line starts a new verse (ends our rhyme collection)
                    verse_lines = split_verse_lines_by_slip_id(line)
                    
                    if verse_lines:
                        # Hit verse line(s) while collecting rhyme - attach rhyme to previous segment
                        if pending_rhyme_target_idx is not None and pending_rhyme_target_idx < len(segments):
                            rhyme_info_raw = '\n'.join(current_rhyme_info_lines).strip()
                            segments[pending_rhyme_target_idx]['rhyme_info_raw'] = rhyme_info_raw
                            pending_rhyme_target_idx = None
                        
                        # Start new verse segment with first verse line
                        state = 'COLLECTING_VERSE'
                        current_start_page = page_num
                        current_segment_title = pending_local_title or current_text_name
                        pending_local_title = None
                        current_verse_lines = []
                        for line_text, slip_id in verse_lines:
                            current_verse_lines.append({
                                'text': line_text,
                                'slip_id': slip_id,
                                'page': page_num
                            })
                        current_rhyme_info_lines = []
                    else:
                        # No slip ID - check if this is a new verse line (pure verse)
                        is_verse, confidence = is_verse_like(line, [])
                        
                        if is_verse and confidence >= 0.5:
                            # Looks like verse - attach collected rhyme and start new segment
                            if pending_rhyme_target_idx is not None and pending_rhyme_target_idx < len(segments):
                                rhyme_info_raw = '\n'.join(current_rhyme_info_lines).strip()
                                segments[pending_rhyme_target_idx]['rhyme_info_raw'] = rhyme_info_raw
                                pending_rhyme_target_idx = None
                            
                            state = 'COLLECTING_VERSE'
                            current_start_page = page_num
                            current_segment_title = pending_local_title or current_text_name
                            pending_local_title = None
                            current_verse_lines = [{
                                'text': line,
                                'slip_id': '',  # No slip ID
                                'page': page_num
                            }]
                            current_rhyme_info_lines = []
                        else:
                            # Still collecting rhyme info - add this line
                            current_rhyme_info_lines.append(line)
                    continue
                
                # Skip footnotes and commentary lines (not in COLLECTING_RHYME_INFO state)
                if is_footnote_or_commentary(line):
                    continue
                
                # Check if this line contains verse (slip IDs)
                # Use Rule A: Split by slip-id boundaries
                verse_lines = split_verse_lines_by_slip_id(line)
                
                # If no slip IDs found, check if line is verse-like using smart detection
                if not verse_lines and state == 'COLLECTING_VERSE':
                    # Get context from recent verse lines
                    context = [v['text'] for v in current_verse_lines[-5:]] if current_verse_lines else []
                    is_verse, confidence = is_verse_like(line, context)
                    
                    if is_verse:
                        # Treat as verse line without slip ID
                        current_verse_lines.append({
                            'text': line,
                            'slip_id': '',  # No slip ID
                            'page': page_num
                        })
                        continue
                
                if state == 'COLLECTING_VERSE':
                    if verse_lines:
                        # Continue collecting verse lines
                        for line_text, slip_id in verse_lines:
                            current_verse_lines.append({
                                'text': line_text,
                                'slip_id': slip_id,
                                'page': page_num
                            })
                    else:
                        # No slip ID found - already handled above with verse detection
                        # If we get here, it's truly not a verse line
                        # Skip footnote markers like [3], [4], etc. - don't flush segment
                        if re.match(r'^\[\d+\]\s*$', line):
                            continue
                        # Otherwise end segment
                        flush_segment()
                        state = 'WAITING_VERSE'
                
                elif state == 'WAITING_VERSE':
                    section_allows_verse = bool(current_section and len(current_section.split('.')) >= 3)
                    if not section_allows_verse:
                        continue

                    if verse_lines:
                        # Start collecting verse
                        state = 'COLLECTING_VERSE'
                        current_start_page = page_num
                        current_segment_title = pending_local_title or current_text_name
                        pending_local_title = None
                        current_verse_lines = []
                        for line_text, slip_id in verse_lines:
                            current_verse_lines.append({
                                'text': line_text,
                                'slip_id': slip_id,
                                'page': page_num
                            })
                    else:
                        # No slip ID - check if this is a pure verse line (common in 王家臺 歸藏)
                        is_verse, confidence = is_verse_like(line, [])
                        
                        if is_verse and confidence >= 0.5:
                            # Start collecting verse with this line
                            state = 'COLLECTING_VERSE'
                            current_start_page = page_num
                            current_segment_title = pending_local_title or current_text_name
                            pending_local_title = None
                            current_verse_lines = [{
                                'text': line,
                                'slip_id': '',  # No slip ID
                                'page': page_num
                            }]
        
        # Flush last segment
        flush_segment()

    segments = merge_related_segments(segments)

    for segment in segments:
        cleaned_lines = []
        hoisted_notes = []
        for line_data in segment.get('lines', []):
            note_text = extract_embedded_note_text(line_data.get('text', ''))
            if note_text:
                hoisted_notes.append(note_text)
                continue
            cleaned_lines.append(line_data)
        segment['lines'] = cleaned_lines
        for note in hoisted_notes:
            if note not in segment['notes']:
                segment['notes'].append(note)

    # Post-process: annotate lines with rhyme patterns and collect diagnostics
    missed_tags = []
    for segment in segments:
        rhyme_sets = parse_rhyme_info(segment['rhyme_info_raw'])
        rhyme_sets = enrich_rhyme_sets_with_tones(rhyme_sets, segment.get('text_name', ''), tone_lookup)
        
        for line_data in segment['lines']:
            annotated, set_id, rhyme_word, rhyme_label, missed_reason = annotate_line_with_rhyme(
                line_data['text'], rhyme_sets, line_data.get('slip_id', '')
            )
            line_data['text_annotated'] = annotated
            line_data['rhyme_set_id'] = set_id
            line_data['rhyme_word'] = rhyme_word
            line_data['rhyme_label'] = rhyme_label
            
            # Rule E: Log missed tags
            if missed_reason:
                missed_tags.append({
                    'poem_id': segment['poem_id'],
                    'line_text': line_data['text'][:60],
                    'slip_id': line_data.get('slip_id', ''),
                    'terminal_token': rhyme_word,
                    'reason': missed_reason
                })
    
    return segments, missed_tags


def export_annotated_poems(segments, output_dir):
    """
    Export clean, book-format text file with all poems with rhyme annotations.
    
    Creates a single consolidated file with all poems.
    
    Only exports segments where:
    - Rhyme info is present
    - At least one line has a rhyme tag [a]/[b]/etc.
    - Lines pass footnote filter
    
    Improvements (2026-02-04):
    1. Treat � as hard error - replace with explicit placeholder
    2. Enforce slip-id boundary - split lines with multiple slip IDs
    3. Fix dangling short lines (PDF wrapping artifacts)
    4. Preserve all markup tokens unchanged
    5. Add integrity checks and issue logging
    
    Format:
    @title: <text_name> <poem_id>
    @annotator: 胡蝶 (from dissertation)
    LINE_TEXT[a]
    LINE_TEXT[b]
    ...
    """
    import re
    from pathlib import Path
    
    exported_count = 0
    all_poems = []  # Collect all poems for consolidated file
    skipped_fragments = 0
    issues = []  # Issue log for integrity checks
    img_tokens_found = set()  # Track all IMG tokens referenced
    
    # Slip-ID pattern
    slip_id_pattern = r'\d+-\d+(?:背)?'
    
    def replace_replacement_char(text, poem_id, line_num):
        """Replace � with explicit placeholder and log it."""
        if '�' in text:
            issues.append({
                'type': 'REPLACEMENT_CHAR',
                'poem_id': poem_id,
                'line_num': line_num,
                'text': text
            })
            # Replace with explicit marker
            text = text.replace('�', '⟦UNK⟧')
        return text
    
    def split_by_slip_ids(text):
        """Split text into multiple lines if it contains multiple slip IDs."""
        matches = list(re.finditer(slip_id_pattern, text))
        
        if len(matches) <= 1:
            return [text]
        
        lines = []
        start = 0
        for match in matches:
            end = match.end()
            line = text[start:end].strip()
            if line:
                lines.append(line)
            start = end
        
        # Add any remaining text after last slip ID
        if start < len(text):
            remainder = text[start:].strip()
            if remainder:
                # If remainder has no slip ID and isn't very short, might be continuation
                # For safety, keep it separate
                lines.append(remainder)
        
        return lines if lines else [text]
    
    def join_dangling_lines(lines):
        """Join very short lines that look like PDF wrapping artifacts."""
        if len(lines) <= 1:
            return lines
        
        result = []
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Check if this is a dangling short line
            # Criteria: <= 3 CJK chars, no punctuation/slip-id/markers
            cjk_count = len(re.findall(r'[\u4E00-\u9FFF]', line))
            has_markers = bool(re.search(r'[，。；：]|' + slip_id_pattern + r'|\[|\]|（|）|⟦|⟧|〖|〗', line))
            
            if cjk_count <= 3 and not has_markers and i + 1 < len(lines):
                # Potentially a dangling line - join with next
                next_line = lines[i + 1]
                joined = line + next_line  # No space for Chinese
                result.append(joined)
                i += 2
            else:
                result.append(line)
                i += 1
        
        return result
    
    def extract_img_tokens(text):
        """Extract all IMG tokens from text."""
        return set(re.findall(r'⟦IMG:[^⟧]+⟧', text))
    
    def check_line_integrity(line, poem_id, line_num):
        """Check line integrity and log issues."""
        # Check: no more than one slip ID per line
        slip_ids = re.findall(slip_id_pattern, line)
        if len(slip_ids) > 1:
            issues.append({
                'type': 'MULTIPLE_SLIP_IDS',
                'poem_id': poem_id,
                'line_num': line_num,
                'slip_ids': slip_ids,
                'text': line
            })
        
        # Check: no remaining �
        if '�' in line:
            issues.append({
                'type': 'REMAINING_REPLACEMENT_CHAR',
                'poem_id': poem_id,
                'line_num': line_num,
                'text': line
            })

    
    for segment in segments:
        # Only export if we have rhyme info
        if not segment.get('rhyme_info_raw'):
            continue
        
        # Check if any lines have tags
        has_tags = any(line.get('rhyme_set_id') for line in segment['lines'])
        
        # Count rhyme words per rhyme set
        rhyme_counts = {}
        for line_data in segment['lines']:
            rid = line_data.get('rhyme_set_id', '')
            if rid:
                rhyme_counts[rid] = rhyme_counts.get(rid, 0) + 1
        
        # Require at least one rhyme set with 2+ lines, OR has rhyme info with 2+ lines
        # (relaxed: poems with rhyme info should be included even if annotation didn't work perfectly)
        has_valid_rhyme = any(count >= 2 for count in rhyme_counts.values())
        has_rhyme_with_lines = segment.get('rhyme_info_raw') and len(segment['lines']) >= 2
        
        if not has_valid_rhyme and not has_rhyme_with_lines:
            skipped_fragments += 1
            continue
        
        # Process lines with all improvements
        # Keep both text and metadata for intelligent processing
        line_items = []
        for line_data in segment['lines']:
            text = line_data['text']
            if is_footnote_or_commentary(text):
                continue
            
            # Use annotated version if available
            line_output = line_data.get('text_annotated', text)
            slip_id = line_data.get('slip_id', '')
            line_items.append({'text': line_output, 'slip_id': slip_id})
        
        if not line_items:
            continue
        
        poem_id = segment['poem_id']
        
        # Step 1: Replace � with explicit placeholder
        for i, item in enumerate(line_items):
            item['text'] = replace_replacement_char(item['text'], poem_id, i+1)
        
        # Step 2: Join dangling short lines (fix PDF wrapping)
        # But DON'T join if both lines have slip IDs (they're separate verses)
        joined_items = []
        i = 0
        while i < len(line_items):
            item = line_items[i]
            line = item['text']
            has_slip = bool(item['slip_id'])
            
            # Check if this is a dangling short line
            cjk_count = len(re.findall(r'[\u4E00-\u9FFF]', line))
            has_markers = bool(re.search(r'[，。；：]|' + slip_id_pattern + r'|\[|\]|（|）|⟦|⟧|〖|〗', line))
            
            # Only join if:
            # 1. Line is short (<= 3 CJK chars) and has no markers
            # 2. Neither current nor next line has slip ID (otherwise they're real verse boundaries)
            # 3. Next line exists
            if (cjk_count <= 3 and not has_markers and not has_slip and 
                i + 1 < len(line_items) and not line_items[i + 1]['slip_id']):
                # Join with next line
                next_line = line_items[i + 1]['text']
                joined = line + next_line  # No space for Chinese
                joined_items.append(joined)
                i += 2
            else:
                joined_items.append(line)
                i += 1
        
        # Step 3: Split by slip-ID boundaries (shouldn't be needed now, but keep for safety)
        final_lines = []
        for line in joined_items:
            split_lines = split_by_slip_ids(line)
            final_lines.extend(split_lines)
        
        # Step 4: Extract IMG tokens and check integrity
        for i, line in enumerate(final_lines):
            img_tokens_found.update(extract_img_tokens(line))
            check_line_integrity(line, poem_id, i+1)
        
        if not final_lines:
            continue
        
        # Prepare poem content with enhanced metadata
        text_name = segment.get('text_name', 'Unknown')
        collection = segment.get('collection', '')
        section = segment.get('section', '')
        slip_range = segment.get('slip_range', '')
        notes = segment.get('notes', [])
        
        poem_content = []
        poem_content.append(f"@title: {text_name} ({poem_id})")
        if collection:
            poem_content.append(f"@collection: {collection}")
        if section:
            poem_content.append(f"@section: {section}")
        if slip_range:
            poem_content.append(f"@slip_range: {slip_range}")
        poem_content.append("@annotator: 胡蝶")
        
        # Add notes if present (footnotes with parallel text references)
        if notes:
            for note in notes:
                poem_content.append(f"@note: {note}")
        
        poem_content.append("")
        poem_content.extend(final_lines)
        
        # Add to consolidated list
        all_poems.append('\n'.join(poem_content))
        
        exported_count += 1
    
    # Write consolidated file with all poems
    if all_poems:
        consolidated_path = Path(output_dir) / 'all_annotated_poems.txt'
        with open(consolidated_path, 'w', encoding='utf-8') as f:
            f.write('\n\n---\n\n'.join(all_poems))
    
    # Write IMG token reference list
    if img_tokens_found:
        img_list_path = Path(output_dir) / 'image_tokens_referenced.txt'
        with open(img_list_path, 'w', encoding='utf-8') as f:
            for token in sorted(img_tokens_found):
                f.write(f'{token}\n')
    
    # Write issues report if any
    if issues:
        import json
        issues_path = Path(output_dir) / 'poem_export_issues.json'
        with open(issues_path, 'w', encoding='utf-8') as f:
            json.dump(issues, f, ensure_ascii=False, indent=2)
        print(f'  ⚠️  Found {len(issues)} integrity issues - see {issues_path}')
    
    return exported_count, skipped_fragments


# ---- CLI and IO ----------------------------------------------------------

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Parse Qin rhymes and export CLDF-friendly outputs')
    parser.add_argument('--mode', choices=['both','lines-only','annotation-only'], default='both')
    parser.add_argument('--outdir', default='outputs')
    parser.add_argument('--pdf', default='hudie2023_qin_rhymes.pdf')
    args = parser.parse_args()

    from datetime import timezone
    ts = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    run_dir = Path(args.outdir) / f'run_{ts}'
    run_dir.mkdir(parents=True, exist_ok=True)
    
    # Create symlink to extracted_images in the run directory
    extracted_images_root = Path('extracted_images')
    extracted_images_in_run = run_dir / 'extracted_images'
    if extracted_images_root.exists():
        # Create relative symlink
        try:
            if extracted_images_in_run.exists():
                extracted_images_in_run.unlink()
            extracted_images_in_run.symlink_to('../../extracted_images', target_is_directory=True)
        except Exception as e:
            print(f"Warning: Could not create symlink to extracted_images: {e}")
    
    # Update latest_run.txt to point to this run
    latest_run_file = Path(args.outdir) / 'latest_run.txt'
    with open(latest_run_file, 'w', encoding='utf-8') as f:
        f.write(str(run_dir) + '\n')

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

    # ---- TSV Validation and Writing (2026-02-04) ----
    
    def normalize_cell(text):
        """Remove embedded newlines and tabs from TSV cell content."""
        if not isinstance(text, str):
            text = str(text)
        # Replace newlines and carriage returns with space
        text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        # Replace tabs with space
        text = text.replace('\t', ' ')
        # Collapse multiple spaces and strip
        text = re.sub(r'\s+', ' ', text).strip()
        return text
    
    def extract_witness_labels(char_string):
        """
        Extract apparatus/witness labels (e.g. BIII, AII) from end of character string.
        Can be: standalone token after 、, or attached to last character.
        Returns: (cleaned_chars, witness_label)
        """
        if not char_string:
            return char_string, ''
        
        witness_pattern = r'([A-Z]{1,3}[IVX]{1,6})$'
        
        # First check if entire string ends with witness label (could be attached)
        match = re.search(witness_pattern, char_string)
        if match:
            witness = match.group(1)
            # Remove witness from end
            cleaned = char_string[:match.start()].rstrip()
            return cleaned, witness
        
        # Fallback: split on separator and check last token
        tokens = [t.strip() for t in char_string.split('、') if t.strip()]
        if not tokens:
            return char_string, ''
        
        last_token = tokens[-1]
        if re.match(r'^[A-Z]{1,3}[IVX]{1,6}$', last_token):
            # Last token is standalone witness label
            cleaned = '、'.join(tokens[:-1])
            return cleaned, last_token
        
        return char_string, ''
    
    def validate_char_tone_alignment(chars, tones):
        """
        Validate that character and tone counts align.
        Returns: (is_valid, char_count, tone_count)
        """
        if not chars or not tones:
            return True, 0, 0
        
        # Split characters on enumeration separator
        char_tokens = [t.strip() for t in chars.split('、') if t.strip()]
        
        # Split tones - try multiple separators
        # Tones might use: space, —, –, -, |
        tone_tokens = []
        for sep in [' ', '—', '–', '-', '|']:
            if sep in tones:
                tone_tokens = [t.strip() for t in tones.split(sep) if t.strip() and t in '平上去入']
                if tone_tokens:
                    break
        
        if not tone_tokens:
            # No recognizable tone separators, count individual tone chars
            tone_tokens = [c for c in tones if c in '平上去入']
        
        char_count = len(char_tokens)
        tone_count = len(tone_tokens)
        
        is_valid = (char_count == tone_count) or char_count == 0 or tone_count == 0
        
        return is_valid, char_count, tone_count
    
    def write_rhyme_tsv_with_validation(rhymes, output_path, rejects_path):
        """
        Write rhyme_output TSV with strict validation.
        
        Returns: (valid_count, reject_count, self_check_passed)
        """
        fieldnames = ['RhymeSegment', 'Characters', 'OC_RhymeGroup', 'Tones', 
                      'Source', 'Page', 'TableID', 'image_refs', 'Witness', 'needs_review']
        
        valid_rows = []
        reject_rows = []
        
        for rhyme in rhymes:
            # Step 1: Normalize all cells
            normalized = {}
            for key in fieldnames:
                if key in ['Witness', 'needs_review']:
                    normalized[key] = ''  # Will be populated
                else:
                    normalized[key] = normalize_cell(rhyme.get(key, ''))
            
            # Step 2: Extract witness labels from Characters
            chars, witness = extract_witness_labels(normalized['Characters'])
            normalized['Characters'] = chars
            normalized['Witness'] = witness
            
            # Step 3: Validate character-tone alignment
            is_valid, char_count, tone_count = validate_char_tone_alignment(
                normalized['Characters'], 
                normalized['Tones']
            )
            
            if not is_valid:
                normalized['needs_review'] = f'ALIGNMENT:{char_count}chars_{tone_count}tones'
            else:
                normalized['needs_review'] = ''
            
            # Step 4: Structural validation - ensure exactly N columns
            row_values = [normalized.get(k, '') for k in fieldnames]
            if len(row_values) != len(fieldnames):
                # Malformed row - reject it
                reject_rows.append({
                    'error': 'COLUMN_COUNT_MISMATCH',
                    'expected': len(fieldnames),
                    'got': len(row_values),
                    'data': str(rhyme)
                })
                continue
            
            valid_rows.append(normalized)
        
        # Write valid rows
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t', 
                                   lineterminator='\n')
            writer.writeheader()
            for row in valid_rows:
                writer.writerow(row)
        
        # Write rejects if any
        if reject_rows:
            with open(rejects_path, 'w', encoding='utf-8', newline='') as f:
                reject_fieldnames = ['error', 'expected', 'got', 'data']
                writer = csv.DictWriter(f, fieldnames=reject_fieldnames, delimiter='\t',
                                       lineterminator='\n')
                writer.writeheader()
                for reject in reject_rows:
                    writer.writerow(reject)
        
        # Self-check: re-read and verify structure
        self_check_passed = True
        with open(output_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            for i, line in enumerate(lines):
                if i == 0:  # Skip header
                    continue
                cols = line.rstrip('\n').split('\t')
                if len(cols) != len(fieldnames):
                    self_check_passed = False
                    print(f'  ⚠️  Self-check failed: line {i+1} has {len(cols)} columns, expected {len(fieldnames)}')
        
        return len(valid_rows), len(reject_rows), self_check_passed

    # write rhyme segments with validation
    rhyme_fname = run_dir / f'rhyme_output.{ts}.tsv'
    rejects_fname = run_dir / f'rhyme_output_rejects.{ts}.tsv'
    valid_count, reject_count, self_check = write_rhyme_tsv_with_validation(
        rhymes, rhyme_fname, rejects_fname
    )
    
    print(f'Wrote rhyme groups to {rhyme_fname}')
    print(f'  - Valid rows: {valid_count}')
    if reject_count > 0:
        print(f'  - Rejected rows: {reject_count} (see {rejects_fname})')
    if not self_check:
        print(f'  ⚠️  Self-check FAILED - TSV structure may be corrupt!')
    else:
        print(f'  ✓ Self-check passed')
    
    # Count rows needing review
    review_count = sum(1 for r in rhymes if validate_char_tone_alignment(
        r.get('Characters', ''), r.get('Tones', ''))[0] == False)
    if review_count > 0:
        print(f'  - Rows needing review (alignment): {review_count}')


    # annotation and lines
    all_warnings = []
    ann = []
    if args.mode in ('both','annotation-only'):
        ann, warnings, ann_summary = make_annotation_rows(rows)
        ann_fname = run_dir / f'rhyme_annotations.{ts}.tsv'
        with open(ann_fname, 'w', encoding='utf-8', newline='') as f:
            fieldnames = ['RowID','TableID','RhymeSegment','page','row_in_source','raw_text','Characters','OC_RhymeGroup','Tones','rhyme_tokens','tone_tokens','rhyme_numbers','Source','alignment_issue','notes','image_refs']
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
        tone_lookup = build_tone_lookup(ann)
        if not tone_lookup:
            ann_for_tones, _, _ = make_annotation_rows(rows)
            tone_lookup = build_tone_lookup(ann_for_tones)

        # Extract table-based poems (from appendix)
        table_lines = make_line_rows(rows)
        
        # Extract prose-embedded poems (from Chapter 2) using layout-based state machine
        chapter2_segments, missed_tags = extract_chapter2_poems(
            args.pdf,
            start_page=21,
            end_page=141,
            tone_lookup=tone_lookup,
        )
        
        # Convert Chapter 2 segments to same format as table lines
        prose_lines = []
        for segment in chapter2_segments:
            for i, line_data in enumerate(segment['lines'], 1):
                prose_lines.append({
                    'POEM': segment['poem_id'],
                    'LINE_ORDER': i,
                    'LINE': line_data['text'],
                    'LINE_ANNOTATED': line_data.get('text_annotated', line_data['text']),
                    'RHYME_SET_ID': line_data.get('rhyme_set_id', ''),
                    'RHYME_WORD': line_data.get('rhyme_word', ''),
                    'RHYME_LABEL': line_data.get('rhyme_label', ''),
                    'RHYME_GROUP': '',  # From table format
                    'RHYME_ID': '',     # From table format
                    'SOURCE': segment['text_name'],
                    'PAGE': line_data['page'],
                    'TABLE_ID': f"CH2_P{line_data['page']:03d}",
                    'RHYME_SEGMENT': '',
                    'SLIP_ID': line_data.get('slip_id', ''),
                    'RHYME_INFO': segment['rhyme_info_raw']
                })
        
        # Write segment index for debugging
        segment_index_fname = run_dir / f'segment_index.{ts}.csv'
        with open(segment_index_fname, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['poem_id', 'text_name', 'collection', 'section', 'slip_range', 'start_page', 'n_lines', 'first_line', 'rhyme_info_raw', 'img_token_count'])
            writer.writeheader()
            for seg in chapter2_segments:
                writer.writerow({
                    'poem_id': seg['poem_id'],
                    'text_name': seg.get('text_name', ''),
                    'collection': seg.get('collection', ''),
                    'section': seg.get('section', ''),
                    'slip_range': seg.get('slip_range', ''),
                    'start_page': seg['start_page'],
                    'n_lines': len(seg['lines']),
                    'first_line': seg['lines'][0]['text'][:50].replace('\n', ' ') if seg['lines'] else '',
                    'rhyme_info_raw': seg['rhyme_info_raw'][:100].replace('\n', ' '),
                    'img_token_count': seg['img_count']
                })
        
        # Combine both sources
        all_lines = table_lines + prose_lines
        
        lines_fname = run_dir / f'poem_lines.{ts}.csv'
        with open(lines_fname, 'w', encoding='utf-8-sig', newline='') as f:
            # Expanded fieldnames to include Chapter 2 metadata and rhyme annotations
            fieldnames = ['POEM', 'LINE_ORDER', 'LINE', 'LINE_ANNOTATED', 'RHYME_SET_ID', 
                         'RHYME_WORD', 'RHYME_LABEL', 'RHYME_GROUP', 'RHYME_ID', 
                         'SOURCE', 'PAGE', 'TABLE_ID', 'RHYME_SEGMENT', 'SLIP_ID', 'RHYME_INFO']
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            for l in all_lines:
                # Fill in missing fields for table lines
                if 'SLIP_ID' not in l:
                    l['SLIP_ID'] = ''
                if 'RHYME_INFO' not in l:
                    l['RHYME_INFO'] = ''
                if 'LINE_ANNOTATED' not in l:
                    l['LINE_ANNOTATED'] = l['LINE']
                if 'RHYME_SET_ID' not in l:
                    l['RHYME_SET_ID'] = ''
                if 'RHYME_WORD' not in l:
                    l['RHYME_WORD'] = ''
                if 'RHYME_LABEL' not in l:
                    l['RHYME_LABEL'] = ''
                writer.writerow(l)
        
        print(f'Wrote line export to {lines_fname}')
        print(f'  - Table-based poems: {len(table_lines)} lines')
        print(f'  - Chapter 2 segments: {len(prose_lines)} lines from {len(chapter2_segments)} segments')
        print(f'  - Total: {len(all_lines)} lines')
        print(f'Wrote segment index to {segment_index_fname}')
        
        # Write missed tags diagnostic file
        if missed_tags:
            missed_tags_fname = run_dir / f'missed_tags.{ts}.json'
            with open(missed_tags_fname, 'w', encoding='utf-8') as f:
                json.dump(missed_tags, f, ensure_ascii=False, indent=2)
            print(f'Missed tags diagnostic: {len(missed_tags)} potential missed rhyme tags in {missed_tags_fname}')
        
        # Validation checks
        segments_with_rhyme = sum(1 for s in chapter2_segments if s['rhyme_info_raw'])
        segments_with_images = sum(1 for s in chapter2_segments if s['img_count'] > 0)
        lines_with_tags = sum(1 for l in prose_lines if l.get('RHYME_SET_ID'))
        print(f'Validation:')
        print(f'  - Segments with rhyme info: {segments_with_rhyme}/{len(chapter2_segments)}')
        print(f'  - Segments with image tokens: {segments_with_images}/{len(chapter2_segments)}')
        print(f'  - Lines with rhyme tags [a/b]: {lines_with_tags}/{len(prose_lines)}')
        if segments_with_images == 0:
            print(f'  ⚠️ WARNING: No image tokens found in Chapter 2 - image merge may not be working')
        
        # Export clean annotated poems to text file
        exported_count, skipped_fragments = export_annotated_poems(chapter2_segments, run_dir)
        consolidated_file = run_dir / 'all_annotated_poems.txt'
        print(f'Exported {exported_count} annotated poems to: {consolidated_file}')
        if skipped_fragments > 0:
            print(f'  (Skipped {skipped_fragments} single-rhyme fragments)')

    warnings_fname = run_dir / f'warnings.{ts}.json'
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
    with open(run_dir / f'metadata.{ts}.json', 'w', encoding='utf-8') as f:
        json.dump(meta_ts, f, ensure_ascii=False, indent=2)

    meta_latest = {
        'dc:title': meta_ts['title'],
        'dc:description': meta_ts['description'],
        'dc:license': meta_ts['license'],
        'files': meta_ts['files'],
        'note': 'This dataset contains conservative extractions. Most poems are not present and are not reconstructed.'
    }
    with open(Path(args.outdir) / 'metadata.json', 'w', encoding='utf-8') as f:
        json.dump(meta_latest, f, ensure_ascii=False, indent=2)

    with open(run_dir / 'README.txt', 'w', encoding='utf-8') as f:
        f.write('This folder contains CLDF-friendly exports derived from hudie2023_qin_rhymes.pdf.\n')
        f.write('NOTE: This extraction is conservative. Many poems are not present in the source tables and have not been reconstructed. Only table-structured rhyme rows are emitted to the annotation outputs.\n')
        f.write('Files: rhyme_output.<ts>.tsv (rhyme segments with RhymeSegment, Characters, OC_RhymeGroup, Tones), rhyme_annotations.<ts>.tsv (row-level annotations), poem_lines.<ts>.csv (UTF-8 with BOM, one line per row with POEM, LINE_ORDER, LINE, RHYME_GROUP, RHYME_ID, NOTES columns), warnings.<ts>.json (alignment/header issues).\n')
        f.write('Images: extracted_images/images_index.json maps stable IDs (IMG_pXXX_NNN) to image files; individual rows contain image_refs and in-situ ⟦IMG:ID⟧ placeholders where non-Unicode glyphs occurred.\n')
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

    summary_fname = run_dir / f'summary.{ts}.json'
    with open(summary_fname, 'w', encoding='utf-8') as sf:
        json.dump(final_summary, sf, ensure_ascii=False, indent=2)
    # write manifest describing this run and exact output files
    manifest = {'run_id': ts, 'files': meta_ts.get('files', []) + [summary_fname.name]}
    with open(run_dir / f'manifest.{ts}.json', 'w', encoding='utf-8') as mf:
        json.dump(manifest, mf, ensure_ascii=False, indent=2)
    # update latest manifest for deterministic lookup by canary
    with open(Path(args.outdir) / 'latest_manifest.json', 'w', encoding='utf-8') as mf:
        json.dump(manifest, mf, ensure_ascii=False, indent=2)
    print(f'Wrote summary to {summary_fname}')


    # write unresolved file and warn if present
    unresolved = final_summary.get('annotation_summary', {}).get('unresolved_rows', [])
    if unresolved:
        unresolved_fname = run_dir / f'unresolved.{ts}.json'
        with open(unresolved_fname, 'w', encoding='utf-8') as uf:
            json.dump(unresolved, uf, ensure_ascii=False, indent=2)
        print(f'Unresolved rows present: {len(unresolved)}. Wrote {unresolved_fname}. Exiting with failure code.')
        sys.exit(2)
    else:
        # if unreferenced images exist, warn but continue
        if final_summary.get('unreferenced_extractions_count', 0) > 0:
            print(f"Warning: {final_summary['unreferenced_extractions_count']} unreferenced extracted images: {final_summary['unreferenced_extractions'][:10]}")
