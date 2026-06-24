"""
xml2xlsx.py — Convert annotated XML corpora to Excel.

Fully generalist: schema-driven extraction from DTD or XML discovery.
All column ordering, text rules, and serialization inferred from schema.
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
      - element_children : {elem: [child_elem, ...]}   (content model, order-preserving)
      - element_attribs  : {elem: [attr_name, ...]}    (declared attributes)
    """
    element_children = {}
    element_attribs = {}

    with open(dtd_path, encoding='utf-8') as f:
        text = f.read()

    # ELEMENT declarations
    for m in re.finditer(r'<!ELEMENT\s+(\w+)\s+\(([^)]+)\)', text):
        elem = m.group(1)
        content = m.group(2)
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
            m_attr = re.match(r'(\w+)\s+', line)
            if m_attr:
                attr_name = m_attr.group(1)
                if attr_name != 'xmlns':
                    element_attribs[current_elem].append(attr_name)
        if line.endswith('>'):
            current_elem = None

    return element_children, element_attribs


# ---------------------------------------------------------------------------
# Schema discovery from XML (fallback)
# ---------------------------------------------------------------------------

def discover_schema_from_xml(xml_path):
    """
    If no DTD: walk the XML tree and discover element hierarchy and attributes.
    Returns schema maps in same format as parse_dtd().
    """
    element_children = {}
    element_attribs = {}
    
    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    def walk(elem):
        tag = elem.tag
        
        # Collect unique child tags in order of first appearance
        if tag not in element_children:
            seen_children = []
            for child in elem:
                if child.tag not in seen_children:
                    seen_children.append(child.tag)
            element_children[tag] = seen_children
        
        # Collect unique attributes
        if tag not in element_attribs:
            element_attribs[tag] = list(elem.attrib.keys())
        
        for child in elem:
            walk(child)
    
    walk(root)
    return element_children, element_attribs


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def clean_text(text):
    if not text:
        return None
    result = ' '.join(text.split())
    return result if result else None


def serialize_paragraph(p_elem, skip_tags=None):
    """Rebuild paragraph preserving XML tags as literal strings."""
    skip_tags = skip_tags or {'p'}

    def _serialize(elem):
        result = ''
        tag = elem.tag
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
    """Recursively extract text and attributes from elem into row."""
    col_base = f'{prefix}{suffix}'

    text = clean_text(''.join(elem.itertext()))
    row[f'{col_base}_text'] = text

    for attr in attribs_map.get(elem.tag, []):
        row[f'{col_base}_{attr}'] = elem.attrib.get(attr)

    for child_tag in children_map.get(elem.tag, []):
        child_elems = elem.findall(child_tag)
        child_prefix = f'{col_base}_{child_tag}'
        if child_elems:
            for idx, child in enumerate(child_elems):
                child_suffix = '' if idx == 0 else f'_{idx + 1}'
                _extract_element(child, child_prefix, children_map,
                                 attribs_map, row, suffix=child_suffix)
        else:
            _fill_absent(child_tag, child_prefix, children_map, attribs_map, row)


def _fill_absent(tag, prefix, children_map, attribs_map, row):
    """Fill None for all columns that would be produced by missing element."""
    row[f'{prefix}_text'] = None
    for attr in attribs_map.get(tag, []):
        row[f'{prefix}_{attr}'] = None
    for child_tag in children_map.get(tag, []):
        _fill_absent(child_tag, f'{prefix}_{child_tag}',
                     children_map, attribs_map, row)


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------

def extract_paragraphs(xml_path, children_map, attribs_map):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    paragraphs = root.findall('p') or root.findall('.//p')
    p_children = children_map.get('p', [])

    rows = []
    for p in paragraphs:
        full_text = ''.join(p.itertext())
        match = re.search(r'[«"]["\s]*([^"»]+)["\s]*[»"]', full_text)
        p_id = match.group(1).strip() if match else None

        row = {
            'p_id': p_id,
            'paragraph_text': serialize_paragraph(p),
        }

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
# Column ordering (schema-driven hierarchy)
# ---------------------------------------------------------------------------

def reorder_columns(df, children_map):
    """
    Reorder columns respecting schema hierarchy.
    Build sort key from schema structure, not hardcoded prefixes.
    """
    META = ['source_file', 'p_id', 'paragraph_text', 'POSITION_INTRODD']
    
    def parse_column_path(col):
        """
        Parse column name into (element_path, attribute_name).
        e.g., 'INTRODD_2_EXPANSION_text' → (['INTRODD', 'EXPANSION'], 'text')
        """
        if col in META:
            return None, None
        
        # Split by uppercase letter boundaries, preserving numbers
        parts = []
        current = ''
        for char in col:
            if char.isupper() and current and current[-1].isalpha():
                parts.append(current)
                current = char
            else:
                current += char
        if current:
            parts.append(current)
        
        # Last part is attribute (e.g., 'text', 'decl') or numeric suffix
        if parts and not parts[-1][0].isupper():
            attr = parts.pop()
        else:
            attr = None
        
        # Remaining parts form element path with instance numbers
        path = []
        for part in parts:
            path.append(part)
        
        return path, attr

    def sort_key(col):
        if col in META:
            return (0, META.index(col), [], col)
        
        path, attr = parse_column_path(col)
        if path is None:
            return (2, 0, [], col)
        
        # Build hierarchical sort key from schema
        sort_hierarchy = []
        current_parent = 'p'
        
        for i, elem_part in enumerate(path):
            # elem_part might be "INTRODD", "2", "EXPANSION", etc.
            if elem_part.isdigit():
                # Instance number
                sort_hierarchy.append(int(elem_part))
            else:
                # Element name: find index in parent's children
                children = children_map.get(current_parent, [])
                try:
                    idx = children.index(elem_part)
                    sort_hierarchy.append(idx)
                except ValueError:
                    sort_hierarchy.append(999)  # Unknown, sort to end
                current_parent = elem_part
        
        return (1, 0, tuple(sort_hierarchy), col)
    
    ordered = sorted(df.columns, key=sort_key)
    return df[ordered]


# ---------------------------------------------------------------------------
# Lowercasing (smart: only text columns, discovered from schema)
# ---------------------------------------------------------------------------

def should_lowercase_element(tag, children_map, attribs_map):
    """
    Heuristic: lowercase text from leaf elements (no children).
    Never lowercase if element has attributes (likely metadata).
    """
    has_children = bool(children_map.get(tag, []))
    has_attrs = bool(attribs_map.get(tag, []))
    return not has_children and not has_attrs


def get_lowercase_columns(df, children_map, attribs_map):
    """Discover which text columns should be lowercased based on schema."""
    lowercase_cols = set()
    
    for col in df.columns:
        if not col.endswith('_text'):
            continue
        
        # Extract element name from column (e.g., 'INTRODD_EXPANSION_text' → 'EXPANSION')
        col_prefix = col[:-5]  # Remove '_text'
        parts = col_prefix.split('_')
        
        # Last uppercase part is the element
        elem = None
        for part in reversed(parts):
            if part and part[0].isupper() and not part.isdigit():
                elem = part
                break
        
        if elem and should_lowercase_element(elem, children_map, attribs_map):
            lowercase_cols.add(col)
    
    return lowercase_cols


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert annotated XML to Excel (fully generalist)")
    parser.add_argument('input', help=".xml file or folder of .xml files")
    parser.add_argument('output', nargs='?', help="Output .xlsx file (optional)")
    parser.add_argument('--dtd', dest='dtd', default=None,
                        help="Path to DTD file for schema")
    parser.add_argument('--no-lowercase', dest='no_lowercase',
                        default=False, action='store_true',
                        help="Don't lowercase text columns")
    parser.add_argument('--no-individual', dest='no_individual',
                        default=False, action='store_true',
                        help="Skip individual per-file xlsx (folder mode only)")
    args = parser.parse_args()

    input_path = args.input
    first_xml = None

    # Find first XML to discover schema if no DTD
    if not args.dtd:
        if os.path.isfile(input_path) and input_path.endswith('.xml'):
            first_xml = input_path
        elif os.path.isdir(input_path):
            xml_files = [f for f in os.listdir(input_path) if f.endswith('.xml')]
            if xml_files:
                first_xml = os.path.join(input_path, xml_files[0])

    # Load schema
    if args.dtd:
        children_map, attribs_map = parse_dtd(args.dtd)
        print(f"[schema] Loaded DTD: {args.dtd}")
    elif first_xml:
        children_map, attribs_map = discover_schema_from_xml(first_xml)
        print(f"[schema] Discovered from: {first_xml}")
    else:
        print("ERROR: Provide --dtd or a valid XML file/folder")
        sys.exit(1)

    def process_rows(rows, df_path):
        df = pd.DataFrame(rows)
        df = df.replace('', np.nan)
        df.dropna(axis=1, how='all', inplace=True)
        
        if not args.no_lowercase:
            lowercase_cols = get_lowercase_columns(df, children_map, attribs_map)
            for col in lowercase_cols:
                if col in df.columns:
                    df[col] = df[col].map(
                        lambda v: v.lower() if isinstance(v, str) else v)
        
        df = reorder_columns(df, children_map)
        format_excel(df, df_path, children_map.get('p', []))
        print(f"  Saved: {df_path}")
        return df

    all_rows = []

    if os.path.isfile(input_path) and input_path.endswith('.xml'):
        rows = extract_paragraphs(input_path, children_map, attribs_map)
        for row in rows:
            row['source_file'] = os.path.basename(input_path)
        out = args.output or input_path.replace('.xml', '.xlsx')
        process_rows(rows, out)

    elif os.path.isdir(input_path):
        xml_files = sorted(f for f in os.listdir(input_path) if f.endswith('.xml'))
        if not xml_files:
            print("No XML files found.")
            sys.exit(1)
        for f in xml_files:
            file_path = os.path.join(input_path, f)
            print(f"  Processing: {f}")
            rows = extract_paragraphs(file_path, children_map, attribs_map)
            if rows and not args.no_individual:
                ind_rows = [dict(r, source_file=f) for r in rows]
                process_rows(ind_rows, file_path.replace('.xml', '.xlsx'))
            for row in rows:
                row['source_file'] = f
            all_rows.extend(rows)

        if all_rows:
            master_out = args.output or os.path.join(input_path, 'master_output.xlsx')
            process_rows(all_rows, master_out)
    else:
        print("Provide a .xml file or a folder containing .xml files.")
        sys.exit(1)


if __name__ == '__main__':
    main()
