import pandas as pd
import xml.etree.ElementTree as ET
import sys
import os
import re
from xml2xlsx.format_excel import format_ppi_bold
import numpy as np


def clean_text(text):
    if not text:
        return None
    result = ' '.join(text.split())
    return result if result else None


def clean_lower(text):
    return ' '.join(text.split()).lower() if text else None


NO_LOWER = {'INTRODD_text', 'paragraph_text', 'APP_text', 'INTRODD_EXPANSION_text', 'INTRODD_EXPANSION_2_text','INTRODD_EXPANSION_3_text'}

# Tags to preserve in paragraph_text (order matters for nested: deepest first)
PRESERVE_TAGS = ['INTRODD', 'VDD', 'EXPANSION', 'MOD', 'PPI', 'NONPPI', 'MD', 'APP']


def serialize_paragraph(p_elem):
    """Rebuild paragraph text preserving XML tags (with attributes) as literal text."""
    def _serialize(elem):
        result = ''
        tag = elem.tag
        if tag != 'p':
            # Reconstruct opening tag WITH attributes
            attrs = ''.join(f' {k}="{v}"' for k, v in elem.attrib.items())
            result += f'<{tag}{attrs}>'
        if elem.text:
            result += elem.text
        for child in elem:
            result += _serialize(child)
        if tag != 'p':
            result += f'</{tag}>'
        if elem.tail:
            result += elem.tail
        return result

    raw = _serialize(p_elem)
    return ' '.join(raw.split()) if raw.strip() else None


def get_introdd_position(p_elem):
    """
    Returns 'ANTE' if INTRODD appears before the first PPI in element order,
    'POST' if INTRODD appears after the first PPI,
    'AUTRE' if INTRODD is present but no PPI exists,
    None if no INTRODD at all.
    """
    tags = [child.tag for child in p_elem]
    has_introdd = 'INTRODD' in tags
    has_ppi = 'PPI' in tags

    if not has_introdd:
        return None
    if not has_ppi:
        return 'AUTRE'

    introdd_idx = tags.index('INTRODD')
    ppi_idx = tags.index('PPI')

    if introdd_idx < ppi_idx:
        return 'ANTE'
    elif introdd_idx > ppi_idx:
        return 'POST'
    else:
        return 'AUTRE'


def get_introdd_position_dd(paragraph_text, ignore_multi_ppi=True):
    if not paragraph_text:
        return None

    if ignore_multi_ppi:
        if len(re.findall(r'<PPI>', paragraph_text)) > 1:
            return None

    introdd_start = re.search(r'<INTRODD>', paragraph_text)
    introdd_end   = re.search(r'</INTRODD>', paragraph_text)

    if not introdd_start:
        return None

    intro_open  = introdd_start.start()
    intro_close = introdd_end.end() if introdd_end else intro_open

    # Build all DD spans: (open_start, close_end, full_text_of_DD)
    dd_spans = []
    for m_open in re.finditer(r'<DD>', paragraph_text):
        m_close = re.search(r'</DD>', paragraph_text[m_open.end():])
        if m_close:
            close_end = m_open.end() + m_close.end()
            dd_text = paragraph_text[m_open.start():close_end]
            dd_spans.append((m_open.start(), close_end, dd_text))

    if not dd_spans:
        return None

    def has_ppi(dd_text):
        return bool(re.search(r'<PPI>', dd_text))

    # --- INCISE case 1: INTRODD is fully inside a single DD ---
    for (dd_open, dd_close_end, dd_text) in dd_spans:
        if dd_open <= intro_open and intro_close <= dd_close_end:
            if has_ppi(dd_text):
                return 'INCISE'

    # --- INCISE case 2: INTRODD is sandwiched between two immediate DDs ---
    # Find the closest DD ending at or before intro_open
    dds_before = [(o, ce, t) for (o, ce, t) in dd_spans if ce <= intro_open]
    # Find the closest DD starting at or after intro_close
    dds_after  = [(o, ce, t) for (o, ce, t) in dd_spans if o >= intro_close]

    if dds_before and dds_after:
        nearest_before = dds_before[-1]   # last DD before INTRODD
        nearest_after  = dds_after[0]     # first DD after INTRODD
        # "Immediate": no other DD boundary exists between them and INTRODD
        no_dd_between_before = not any(
            o > nearest_before[1] and ce <= intro_open
            for (o, ce, t) in dd_spans
        )
        no_dd_between_after = not any(
            o >= intro_close and o < nearest_after[0]
            for (o, ce, t) in dd_spans
        )
        if no_dd_between_before and no_dd_between_after:
            # INCISE only if there is nothing between the two DDs except the INTRODD
            # (and punctuation/whitespace). Any narrator text outside INTRODD → POST/ANTE.
            between = paragraph_text[nearest_before[1]:nearest_after[0]]
            without_introdd = re.sub(r'<INTRODD>.*?</INTRODD>', '', between, flags=re.DOTALL)
            without_tags = re.sub(r'<[^>]+>', '', without_introdd)

            # A dash after removing the INTRODD signals a new speaker turn → not INCISE
            has_turn_change = bool(re.search(r'[–—]', without_tags))

            residual = re.sub(r'[\s\W]+', '', without_tags)

            if residual == '' and not has_turn_change:
                if has_ppi(nearest_before[2]) or has_ppi(nearest_after[2]):
                    return 'INCISE'
    # --- ANTE: INTRODD comes before a DD that contains PPI ---
    if dds_after and any(has_ppi(t) for (o, ce, t) in dds_after):
        return 'ANTE'

    # --- POST: INTRODD comes after a DD that contains PPI ---
    if dds_before and any(has_ppi(t) for (o, ce, t) in dds_before):
        return 'POST'

    return None


def extract_paragraphs(xml_path, compute_position=False):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    rows = []

    for p in root.findall('p'):
        full_text = ''.join(p.itertext())
        match = re.search(r'[«"]["\s]*([^"»]+)["\s]*[»"]', full_text)
        p_id = match.group(1).strip() if match else None

        row = {
            'p_id': p_id,
            'paragraph_text': serialize_paragraph(p)
        }

        # --- INTRODD (multiple) ---
        introdds = p.findall('INTRODD')
        for i, introdd in enumerate(introdds):
            suffix = '' if i == 0 else f'_{i+1}'
            row[f'INTRODD{suffix}_text'] = clean_text(''.join(introdd.itertext()))
            
            # Get all INTRODD attributes
            props = introdd.attrib
            for key, value in props.items():
                row[f'INTRODD{suffix}_{key}'] = value
            
            # --- INTRODD EXPANSION ---
            expansions = introdd.findall('EXPANSION')
            for j, exp in enumerate(expansions):
                exp_suffix = '' if j == 0 else f'_{j+1}'
                row[f'INTRODD{suffix}_EXPANSION{exp_suffix}_text'] = clean_text(''.join(exp.itertext()))
                props = exp.attrib
                for key, value in props.items():
                    row[f'INTRODD{suffix}_EXPANSION{exp_suffix}_{key}'] = value
            if not expansions:
                row[f'INTRODD{suffix}_EXPANSION_text'] = None
                row[f'INTRODD{suffix}_EXPANSION_constr'] = None
                row[f'INTRODD{suffix}_EXPANSION_type'] = None
            
            # --- INTRODD MOD (multiple) ---
            mods = introdd.findall('MOD')
            for j, mod in enumerate(mods):
                mod_suffix = '' if j == 0 else f'_{j+1}'
                row[f'INTRODD{suffix}_MOD{mod_suffix}_text'] = clean_text(''.join(mod.itertext()))
                props = mod.attrib
                for key, value in props.items():
                    row[f'INTRODD{suffix}_MOD{mod_suffix}_{key}'] = value
            if not mods:
                row[f'INTRODD{suffix}_MOD_text'] = None
            
            # --- INTRODD VDD (multiple) ---
            vdds = introdd.findall('VDD')
            for j, vdd in enumerate(vdds):
                vdd_suffix = '' if j == 0 else f'_{j+1}'
                row[f'INTRODD{suffix}_VDD{vdd_suffix}_text'] = clean_text(''.join(vdd.itertext()))
                props = vdd.attrib
                for key, value in props.items():
                    row[f'INTRODD{suffix}_VDD{vdd_suffix}_{key}'] = value
                
                # --- VDD EXPANSION ---
                vdd_expansions = vdd.findall('EXPANSION')
                for k, exp in enumerate(vdd_expansions):
                    vdd_exp_suffix = '' if k == 0 else f'_{k+1}'
                    row[f'INTRODD{suffix}_VDD{vdd_suffix}_EXPANSION{vdd_exp_suffix}_text'] = clean_text(''.join(exp.itertext()))
                    props = exp.attrib
                    for key, value in props.items():
                        row[f'INTRODD{suffix}_VDD{vdd_suffix}_EXPANSION{vdd_exp_suffix}_{key}'] = value
                if not vdd_expansions:
                    row[f'INTRODD{suffix}_VDD{vdd_suffix}_EXPANSION_text'] = None
                    row[f'INTRODD{suffix}_VDD{vdd_suffix}_EXPANSION_constr'] = None
                    row[f'INTRODD{suffix}_VDD{vdd_suffix}_EXPANSION_type'] = None
            if not vdds:
                row[f'INTRODD{suffix}_VDD_text'] = None
                row[f'INTRODD{suffix}_VDD_type'] = None
                row[f'INTRODD{suffix}_VDD_EXPANSION_text'] = None
                row[f'INTRODD{suffix}_VDD_EXPANSION_constr'] = None
                row[f'INTRODD{suffix}_VDD_EXPANSION_type'] = None
        
        if not introdds:
            row['INTRODD_text'] = None
            row['INTRODD_EXPANSION_text'] = None
            row['INTRODD_EXPANSION_constr'] = None
            row['INTRODD_EXPANSION_type'] = None
            row['INTRODD_MOD_text'] = None
            row['INTRODD_VDD_text'] = None
            row['INTRODD_VDD_type'] = None
            row['INTRODD_VDD_EXPANSION_text'] = None
            row['INTRODD_VDD_EXPANSION_constr'] = None
            row['INTRODD_VDD_EXPANSION_type'] = None

        # --- POSITION_INTRODD ---
        if compute_position:
            row['POSITION_INTRODD'] = get_introdd_position(p)

        # --- All PPIs (multiple MD in each PPI) ---
        ppis = p.findall('PPI')
        for i, ppi in enumerate(ppis):
            suffix = '' if i == 0 else f'_{i+1}'
            row[f'PPI{suffix}_text'] = clean_text(''.join(ppi.itertext()))
            props = ppi.attrib
            for key, value in props.items():
                row[f'PPI{suffix}_{key}'] = value

            # --- MD in PPI (multiple) ---
            mds_in_ppi = ppi.findall('MD')
            for j, md in enumerate(mds_in_ppi):
                md_suffix = '' if j == 0 else f'_{j+1}'
                row[f'PPI{suffix}_MD{md_suffix}_text'] = clean_text(''.join(md.itertext()))
                props = md.attrib
                for key, value in props.items():
                    row[f'PPI{suffix}_MD{md_suffix}_{key}'] = value
            if not mds_in_ppi:
                row[f'PPI{suffix}_MD_text'] = None
        
        if not ppis:
            row['PPI_text'] = None
            row['PPI_decl'] = None
            row['PPI_type'] = None
            row['PPI_MD_text'] = None

        # --- All NONPPIs ---
        nonppis = p.findall('NONPPI')
        for i, nonppi in enumerate(nonppis):
            suffix = '' if i == 0 else f'_{i+1}'
            row[f'NONPPI{suffix}_text'] = clean_text(''.join(nonppi.itertext()))
            props = nonppi.attrib
            for key, value in props.items():
                row[f'NONPPI{suffix}_{key}'] = value
        if not nonppis:
            row['NONPPI_text'] = None

        # --- Standalone MD ---
        standalone_mds = [child for child in p if child.tag == 'MD']
        for i, md in enumerate(standalone_mds):
            suffix = '' if i == 0 else f'_{i+1}'
            row[f'MD{suffix}_text'] = clean_text(''.join(md.itertext()))
            props = md.attrib
            for key, value in props.items():
                row[f'MD{suffix}_{key}'] = value
        if not standalone_mds:
            row['MD_text'] = None

        # --- All APPs ---
        apps = p.findall('APP')
        for i, app in enumerate(apps):
            suffix = '' if i == 0 else f'_{i+1}'
            row[f'APP{suffix}_text'] = clean_text(''.join(app.itertext()))
            props = app.attrib
            for key, value in props.items():
                row[f'APP{suffix}_{key}'] = value
        if not apps:
            row['APP_text'] = None

        rows.append(row)

    return rows


def reorder_columns(df):
    """Reorder columns to reflect XML hierarchy."""
    PRIORITY = [
        'source_file',
        'p_id',
        'paragraph_text',
        'POSITION_INTRODD',
        'POSITION_INTRODD_DD',
        'INTRODD',
        'PPI',
        'NONPPI',
        'MD',
        'APP',
    ]

    def sort_key(col):
        for i, prefix in enumerate(PRIORITY):
            if col == prefix or col.startswith(prefix + '_'):
                return (i, col)
        return (len(PRIORITY), col)

    ordered = sorted(df.columns, key=sort_key)
    return df[ordered]


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Convert annotated XML to Excel")
    parser.add_argument('input', help=".xml file, folder, or .xlsx file")
    parser.add_argument('output', nargs='?', help="Output file (optional)")
    parser.add_argument(
        '--ignore-multi-ppi', dest='ignore_multi_ppi',
        default=True, action=argparse.BooleanOptionalAction,
        help="Ignore paragraphs with multiple PPIs for POSITION_INTRODD_DD (default: True)"
    )
    parser.add_argument(
        '--position', dest='compute_position',
        default=False, action='store_true',
        help="Compute POSITION_INTRODD and POSITION_INTRODD_DD columns (default: off)"
    )
    args = parser.parse_args()

    input_path = args.input
    ignore_multi_ppi = args.ignore_multi_ppi
    compute_position = args.compute_position

    print(f"[config] ignore_multi_ppi = {ignore_multi_ppi}")
    print(f"[config] compute_position = {compute_position}")

    all_rows = []

    if os.path.isfile(input_path) and input_path.endswith('.xml'):
        rows = extract_paragraphs(input_path, compute_position)
        for row in rows:
            row['source_file'] = os.path.basename(input_path)
        all_rows = rows
        output_file = args.output or input_path.replace(".xml", ".xlsx")

    elif os.path.isdir(input_path):
        output_file = args.output or os.path.join(input_path, "master_output.xlsx")
        for f in sorted(os.listdir(input_path)):
            if f.endswith('.xml'):
                file_path = os.path.join(input_path, f)
                rows = extract_paragraphs(file_path, compute_position)
                if rows:
                    individual_df = pd.DataFrame(rows)
                    individual_df = individual_df.replace('', np.nan)
                    individual_df.dropna(axis=1, how='all', inplace=True)
                    text_cols = [c for c in individual_df.columns if c.endswith('_text') and c not in NO_LOWER]
                    individual_df[text_cols] = individual_df[text_cols].apply(
                        lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v)
                    )
                    if compute_position and 'paragraph_text_dd' in individual_df.columns:
                        pos = individual_df.columns.get_loc('paragraph_text_dd') + 1
                        individual_df.insert(
                            pos, 'POSITION_INTRODD_DD',
                            individual_df['paragraph_text_dd'].map(
                                lambda t: get_introdd_position_dd(t, ignore_multi_ppi),
                                na_action='ignore'
                            )
                        )
                    individual_df = reorder_columns(individual_df)
                    individual_output = file_path.replace(".xml", ".xlsx")
                    format_ppi_bold(individual_df, individual_output)
                    print(f"Saved individual file: {individual_output}")
                for row in rows:
                    row['source_file'] = f
                all_rows.extend(rows)

    elif os.path.isfile(input_path) and input_path.endswith('.xlsx'):
        if not compute_position:
            print("Pass --position to recompute POSITION_INTRODD_DD on an existing xlsx.")
            sys.exit(0)
        df = pd.read_excel(input_path)
        if 'paragraph_text_dd' not in df.columns:
            print("Error: column 'paragraph_text_dd' not found in the xlsx file.")
            sys.exit(1)
        if 'POSITION_INTRODD_DD' in df.columns:
            df['POSITION_INTRODD_DD'] = df['paragraph_text_dd'].map(
                lambda t: get_introdd_position_dd(t, ignore_multi_ppi), na_action='ignore'
            )
        else:
            pos = df.columns.get_loc('paragraph_text_dd') + 1
            df.insert(
                pos, 'POSITION_INTRODD_DD',
                df['paragraph_text_dd'].map(
                    lambda t: get_introdd_position_dd(t, ignore_multi_ppi), na_action='ignore'
                )
            )
        df = reorder_columns(df)
        format_ppi_bold(df, input_path)
        print(f"Updated {input_path}")
        sys.exit(0)

    else:
        print("Please provide a valid .xml file, .xlsx file, or folder containing .xml files")
        sys.exit(1)

    if all_rows:
        df = pd.DataFrame(all_rows)
        cols = ['source_file'] + [c for c in df.columns if c != 'source_file']
        df = df[cols]
        text_cols = [c for c in df.columns if c.endswith('_text') and c not in NO_LOWER]
        df[text_cols] = df[text_cols].apply(
            lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v)
        )
        df = df.replace('', np.nan)
        df.dropna(axis=1, how='all', inplace=True)
        if compute_position and 'paragraph_text_dd' in df.columns:
            pos = df.columns.get_loc('paragraph_text_dd') + 1
            df.insert(
                pos, 'POSITION_INTRODD_DD',
                df['paragraph_text_dd'].map(
                    lambda t: get_introdd_position_dd(t, ignore_multi_ppi),
                    na_action='ignore'
                )
            )
        df = reorder_columns(df)
        format_ppi_bold(df, output_file)
        print(f"Saved master file to {output_file}")
    else:
        print("No data extracted from input")


if __name__ == '__main__':
    main()
if __name__ == '__main__':
    main()
