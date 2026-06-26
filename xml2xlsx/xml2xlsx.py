"""
xml2xlsx.py — Convert annotated XML corpora to Excel.

Tag extraction is driven by the DTD: pass --dtd path/to/schema.dtd to
automatically discover element hierarchy and attributes. Without --dtd the
script falls back to the built-in PREFAB schema (legacy behaviour).
"""

import pandas as pd
import xml.etree.ElementTree as ET
import sys
import os
import re
import argparse
import numpy as np
from xml2xlsx.format_excel import format_excel


# ---------------------------------------------------------------------------
# DTD parsing
# ---------------------------------------------------------------------------

def parse_dtd(dtd_path):
    """
    Parse a DTD file and return:
      - element_children : {elem: [child_elem, ...]}   (content model)
      - element_attribs  : {elem: [attr_name, ...]}    (declared attributes,
                            excluding xmlns boilerplate)
    Order follows declaration order in the DTD.
    """
    element_children = {}
    element_attribs = {}

    with open(dtd_path, encoding='utf-8') as f:
        text = f.read()

    # ELEMENT declarations
    for m in re.finditer(r'<!ELEMENT\s+(\w+)\s+\(([^)]+)\)', text):
        elem = m.group(1)
        content = m.group(2)
        # Extract child element names (skip #PCDATA)
        children = [c.strip().lstrip('(').rstrip(')*+?')
                    for c in re.split(r'[|,]', content)
                    if c.strip() and '#PCDATA' not in c]
        children = [c for c in children if c]
        element_children[elem] = children

    # ATTLIST declarations
    current_elem = None
    for line in text.splitlines():
        line = line.strip()
        m_attlist = re.match(r'<!ATTLIST\s+(\w+)', line)
        if m_attlist:
            current_elem = m_attlist.group(1)
            element_attribs.setdefault(current_elem, [])
            continue
        if current_elem and line and not line.startswith('<!'):
            # Attribute line: name type default
            m_attr = re.match(r'(\w+)\s+', line)
            if m_attr:
                attr_name = m_attr.group(1)
                if attr_name != 'xmlns':
                    element_attribs[current_elem].append(attr_name)
        if line.endswith('>'):
            current_elem = None

    return element_children, element_attribs


# ---------------------------------------------------------------------------
# Built-in PREFAB schema (fallback when no DTD supplied)
# ---------------------------------------------------------------------------

PREFAB_CHILDREN = {
    'text':     ['p'],
    'p':        ['INTRODD', 'PPI', 'NONPPI', 'MD', 'APP'],
    'INTRODD':  ['EXPANSION', 'VDD', 'MOD'],
    'PPI':      ['MD'],
    'VDD':      ['EXPANSION'],
    'NONPPI':   [],
    'MD':       [],
    'APP':      [],
    'MOD':      [],
    'EXPANSION':[],
}

PREFAB_ATTRIBS = {
    'INTRODD':  ['position'],
    'PPI':      ['decl', 'type', 'position'],
    'VDD':      ['type', 'lemme'],
    'EXPANSION':['constr', 'type'],
}

# Columns whose text must NOT be lowercased
NO_LOWER = {'INTRODD_text', 'paragraph_text', 'APP_text',
            'INTRODD_EXPANSION_text', 'INTRODD_EXPANSION_2_text',
            'INTRODD_EXPANSION_3_text'}

# Tags serialised inline in paragraph_text
SERIALIZE_SKIP = {'p'}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def clean_text(text):
    if not text:
        return None
    result = ' '.join(text.split())
    return result if result else None


def serialize_paragraph(p_elem, skip_tags=None):
    """Rebuild paragraph text preserving XML tags as literal strings."""
    skip_tags = skip_tags or {'p'}

    def _serialize(elem):
        result = ''
        tag = elem.tag
        if callable(tag):  # Extract comment with tags
            if elem.text:
                result += f'<!-- {elem.text} -->'
            return result
        if tag not in skip_tags:
            attrs = ''.join(f' {k}="{v}"' for k, v in elem.attrib.items())
            result += f'<{tag}{attrs}>'
        if elem.text:
            result += elem.text
        for child in elem:
            result += _serialize(child)
        if tag not in skip_tags:
            result += f'</{tag}>'
        if elem.tail:
            result += elem.tail
        return result

    raw = _serialize(p_elem)
    return ' '.join(raw.split()) if raw.strip() else None
# ---------------------------------------------------------------------------
# Generic recursive extraction
# ---------------------------------------------------------------------------

def _extract_element(elem, prefix, children_map, attribs_map, row, suffix=''):
    """
    Recursively extract text and attributes from *elem* into *row*.
    prefix  : column name prefix built so far (e.g. 'INTRODD_VDD')
    suffix  : numeric suffix for repeated siblings (e.g. '' or '_2')
    """
    col_base = f'{prefix}{suffix}'

    # Text content
    text = clean_text(''.join(elem.itertext()))
    row[f'{col_base}_text'] = text

    # Declared attributes
    for attr in attribs_map.get(elem.tag, []):
        row[f'{col_base}_{attr}'] = elem.attrib.get(attr)

    # Recurse into known children
    for child_tag in children_map.get(elem.tag, []):
        child_elems = elem.findall(child_tag)
        child_prefix = f'{col_base}_{child_tag}'
        if child_elems:
            for idx, child in enumerate(child_elems):
                child_suffix = '' if idx == 0 else f'_{idx + 1}'
                _extract_element(child, child_prefix, children_map,
                                 attribs_map, row, suffix=child_suffix)
        else:
            # Guarantee columns exist even when absent
            _fill_absent(child_tag, child_prefix, children_map, attribs_map, row)


def _fill_absent(tag, prefix, children_map, attribs_map, row):
    """Fill None for all columns that would be produced by a missing element."""
    row[f'{prefix}_text'] = None
    for attr in attribs_map.get(tag, []):
        row[f'{prefix}_{attr}'] = None
    for child_tag in children_map.get(tag, []):
        _fill_absent(child_tag, f'{prefix}_{child_tag}',
                     children_map, attribs_map, row)


# ---------------------------------------------------------------------------
# POSITION_INTRODD
# ---------------------------------------------------------------------------

def get_introdd_position(p_elem):
    tags = [child.tag for child in p_elem]
    if 'INTRODD' not in tags:
        return None
    if 'PPI' not in tags:
        return 'AUTRE'
    return 'ANTE' if tags.index('INTRODD') < tags.index('PPI') else 'POST'


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------

from lxml import etree

def extract_paragraphs(xml_path, children_map, attribs_map,
                       compute_position=False):
    tree = etree.parse(xml_path)
    root = tree.getroot()
    paragraphs = root.findall('p') or root.findall('.//p')
    p_children = children_map.get('p', [])
    rows = []
    for p_index, p in enumerate(paragraphs):
        full_text = ''.join(p.itertext())
        match = re.search(r'[«"]["\s]*([^"»]+)["\s]*[»"]', full_text)
        p_id = match.group(1).strip() if match else None
        row = {
            'p_index': p_index,
            'p_id': p_id,
            'paragraph_text': serialize_paragraph(p),
        }
        if compute_position:
            row['POSITION_INTRODD'] = get_introdd_position(p)

        xml_comments = [
            child.text.strip()
            for child in p
            if isinstance(child.tag, str) is False and 'Comment' in str(type(child))
            and child.text
        ]
        row['xml_comments'] = ' | '.join(xml_comments) if xml_comments else None

        for child_tag in p_children:
            child_elems = p.findall(child_tag)
            child_prefix = child_tag
            if child_elems:
                for idx, child in enumerate(child_elems):
                    child_suffix = '' if idx == 0 else f'_{idx + 1}'
                    _extract_element(child, child_prefix, children_map,
                                     attribs_map, row, suffix=child_suffix)
            else:
                _fill_absent(child_tag, child_prefix,
                             children_map, attribs_map, row)
        rows.append(row)
    return rows

# ---------------------------------------------------------------------------
# Excel to XML (reverse)
# ---------------------------------------------------------------------------

def parse_column_name(col_name):
    """Parse column name to extract element path with indices."""
    skip_cols = {'p_id', 'paragraph_text', 'source_file', 'POSITION_INTRODD', 'xml_comments'}
    if col_name in skip_cols:
        return None
    
    parts = col_name.split('_')
    
    # Separate content type (text or attribute) from path
    if parts[-1] == 'text':
        content_parts = parts[:-1]
        content_type = 'text'
    else:
        attr_name = parts[-1]
        content_parts = parts[:-1]
        content_type = 'attr'
    
    # Parse path with indices: [tag, [index], tag, [index], ...]
    path = []
    indices = {}  # {tag_name: index}
    
    i = 0
    while i < len(content_parts):
        tag = content_parts[i]
        if tag.isdigit():
            # This is an index for the previous tag
            if path:
                indices[path[-1]] = int(tag)
            i += 1
        else:
            path.append(tag)
            i += 1
    
    result = {'path': path, 'type': content_type, 'indices': indices}
    if content_type == 'attr':
        result['attr_name'] = attr_name
    return result


def xlsx2xml(excel_path, xml_path, output_path, children_map):
    """Read modified Excel, apply changes back to XML."""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    df = pd.read_excel(excel_path)
    paragraphs = root.findall('p') or root.findall('.//p')
    has_p_index = 'p_index' in df.columns

    for _, row in df.iterrows():
        p_id = row.get('p_id')
        if pd.isna(p_id):
            continue

        # Locate the target <p> element
        p_elem = None
        if has_p_index and not pd.isna(row.get('p_index')):
            idx = int(row['p_index'])
            if idx < len(paragraphs):
                p_elem = paragraphs[idx]
        if p_elem is None:
            # Fallback: match by p_id (legacy Excel without p_index)
            for p in paragraphs:
                full_text = ''.join(p.itertext())
                match = re.search(r'[«"]["\s]*([^"»]+)["\s]*[»"]', full_text)
                found_id = match.group(1).strip() if match else None
                if found_id == p_id:
                    p_elem = p
                    break

        if p_elem is None:
            continue

        # Pre-compute columns that have children → skip their _text
        non_empty_cols = {c for c, v in row.items() if not pd.isna(v)}
        has_children = set()
        for col in non_empty_cols:
            parsed = parse_column_name(col)
            if not parsed or parsed['type'] != 'text':
                continue
            prefix = '_'.join(col.split('_')[:-1])
            for other in non_empty_cols:
                if other != col and other.startswith(prefix + '_'):
                    has_children.add(col)
                    break

        for col_name, value in row.items():
            if pd.isna(value):
                continue

            parsed = parse_column_name(col_name)
            if not parsed:
                continue

            if parsed['type'] == 'text' and col_name in has_children:
                continue

            current = p_elem
            for tag in parsed['path']:
                target_index = parsed['indices'].get(tag, 0)
                matches = current.findall(tag)

                if len(matches) > target_index:
                    elem = matches[target_index]
                elif len(matches) > 0:
                    elem = matches[-1]
                else:
                    expected_children = children_map.get(current.tag, [])
                    if tag in expected_children:
                        insert_index = 0
                        for expected_tag in expected_children:
                            if expected_tag == tag:
                                break
                            insert_index += len(current.findall(expected_tag))
                    else:
                        insert_index = len(current)
                    elem = ET.Element(tag)
                    print(f"[DEBUG] Création de <{tag}> dans <{current.tag}> (p_id={p_id})")
                    current.insert(insert_index, elem)

                current = elem

            if parsed['type'] == 'text':
                current.text = str(value)
            else:
                current.attrib[parsed['attr_name']] = str(value)

    tree.write(output_path, encoding='utf-8')


# ---------------------------------------------------------------------------
# Column ordering
# ---------------------------------------------------------------------------

def reorder_columns(df, p_children):
    """Order columns: meta first, then by p-child declaration order."""
    """
    this stupid function used to mess up the order that was already good, its kept here just as a dummy, maybe i'd need it later. It does nothing for now
    """
    META = ['source_file', 'p_id', 'paragraph_text', 'POSITION_INTRODD']
   

    return df


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert annotated XML to Excel")
    parser.add_argument('input', help=".xml file or folder of .xml files")
    parser.add_argument('output', nargs='?', help="Output .xlsx file (optional)")
    parser.add_argument('--dtd', dest='dtd', default=None,
                        help="Path to DTD file for schema-driven extraction")
    parser.add_argument('--position', dest='compute_position',
                        default=False, action='store_true',
                        help="Compute POSITION_INTRODD column")
    parser.add_argument('--no-individual', dest='no_individual',
                        default=False, action='store_true',
                        help="Skip individual per-file xlsx (folder mode only)")
    args = parser.parse_args()

    # Load schema
    if args.dtd:
        children_map, attribs_map = parse_dtd(args.dtd)
        print(f"[schema] Loaded DTD: {args.dtd}")
    else:
        children_map, attribs_map = PREFAB_CHILDREN, PREFAB_ATTRIBS
        print("[schema] Using built-in PREFAB schema")

    p_children = children_map.get('p', [])
    input_path = args.input
    compute_position = args.compute_position

    def process_rows(rows, df_path, xml_path=None, lowercased=True):
        df = pd.DataFrame(rows)
        if lowercased:
            text_cols = [c for c in df.columns
                         if c.endswith('_text') and c not in NO_LOWER]
            df[text_cols] = df[text_cols].apply(
                lambda col: col.map(
                    lambda v: v.lower() if isinstance(v, str) else v))
        df = reorder_columns(df, p_children)
        from xml2xlsx.integrity_check import check_counts
        check_counts(df, xml_path, args.dtd)
        format_excel(df, df_path, p_children)
        print(f"  Saved: {df_path}")
        return df

    all_rows = []

    if os.path.isfile(input_path) and input_path.endswith('.xml'):
        rows = extract_paragraphs(input_path, children_map, attribs_map,
                                  compute_position)
        for row in rows:
            row['source_file'] = os.path.basename(input_path)
        out = args.output or input_path.replace('.xml', '.xlsx')
        process_rows(rows, out, xml_path=input_path,lowercased=False)

    elif os.path.isdir(input_path):
        xml_files = sorted(f for f in os.listdir(input_path) if f.endswith('.xml'))
        if not xml_files:
            print("No XML files found.")
            sys.exit(1)
        for f in xml_files:
            file_path = os.path.join(input_path, f)
            print(f"  Processing: {f}")
            rows = extract_paragraphs(file_path, children_map, attribs_map,
                                      compute_position)
            if rows and not args.no_individual:
                ind_rows = [dict(r, source_file=f) for r in rows]
                process_rows(ind_rows,
                             file_path.replace('.xml', '.xlsx'),
                             xml_path=f,
                             lowercased=True)
            for row in rows:
                row['source_file'] = f
            all_rows.extend(rows)

        if all_rows:
            master_out = args.output or os.path.join(input_path, 'master_output.xlsx')
            process_rows(all_rows, master_out, lowercased=True)
    else:
        print("Provide a .xml file or a folder containing .xml files.")
        sys.exit(1)


if __name__ == '__main__':
    main()
