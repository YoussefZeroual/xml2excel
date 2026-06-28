"""
format_excel.py — Excel formatting for xml2xlsx output.

Tag colours are configurable via TAG_COLORS. Any tag not listed falls back
to TAG_COLOR_DEFAULT. No element names are hardcoded in the logic.
"""

import re
import math
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

# ---------------------------------------------------------------------------
# Tag colour configuration — edit here or pass a custom dict to format_excel()
# ---------------------------------------------------------------------------

TAG_COLORS = {
    'INTRODD':   '#1F4E79',  # dark blue
    'VDD':       '#7030A0',  # purple
    'EXPANSION': '#C55A11',  # dark orange
    'MOD':       '#833C00',  # brown
    'PPI':       '#C00000',  # dark red
    'NONPPI':    '#375623',  # dark green
    'MD':        '#2E75B6',  # medium blue
    'APP':       '#595959',  # dark grey
    'DD':        '#1D6B5E',  # teal
}
TAG_COLOR_DEFAULT = '#000000'  # black for unknown tags

# Column whose value is rendered with inline tag coloring
RICH_TEXT_COLUMNS = {'paragraph_text'}


# ---------------------------------------------------------------------------
# Rich-text helpers
# ---------------------------------------------------------------------------

def _parse_tagged_text(text):
    pattern = re.compile(r'(<(/?)(\w+)([^>]*)>)')
    segments = []
    pos = 0
    for m in pattern.finditer(text):
        start, end = m.start(), m.end()
        if pos < start:
            segments.append({'type': 'text', 'value': text[pos:start], 'tag': None})
        is_close = m.group(2) == '/'
        segments.append({
            'type': 'close_tag' if is_close else 'open_tag',
            'value': m.group(1),
            'tag': m.group(3),
        })
        pos = end
    if pos < len(text):
        segments.append({'type': 'text', 'value': text[pos:], 'tag': None})
    return segments


def _build_rich_args(text, workbook, tag_colors=None):
    """
    Build args list for worksheet.write_rich_string().
    Returns None if text has no tags (caller should use plain write).
    """
    tag_colors = tag_colors or TAG_COLORS
    segments = _parse_tagged_text(text)
    if not any(s['type'] in ('open_tag', 'close_tag') for s in segments):
        return None

    fmt_cache = {}
    plain_fmt = workbook.add_format({'font_color': '#000000', 'text_wrap': True, 'valign': 'top'})
    tag_label_fmt = workbook.add_format({'font_color': '#AAAAAA', 'italic': True,
                                         'text_wrap': True, 'valign': 'top'})

    def get_fmt(tag_name):
        if tag_name not in fmt_cache:
            color = tag_colors.get(tag_name, TAG_COLOR_DEFAULT)
            fmt_cache[tag_name] = workbook.add_format({
                'font_color': color, 'bold': True,
                'text_wrap': True, 'valign': 'top',
            })
        return fmt_cache[tag_name]

    args = []
    stack = []
    for seg in segments:
        if seg['type'] == 'open_tag':
            stack.append(seg['tag'])
            args += [tag_label_fmt, seg['value']]
        elif seg['type'] == 'close_tag':
            args += [tag_label_fmt, seg['value']]
            if stack and stack[-1] == seg['tag']:
                stack.pop()
        else:
            val = seg['value']
            if not val:
                continue
            fmt = get_fmt(stack[-1]) if stack else plain_fmt
            args += [fmt, val]

    # Filter empty strings
    filtered = []
    i = 0
    while i < len(args):
        if isinstance(args[i], str):
            if args[i]:
                filtered.append(args[i])
            i += 1
        else:
            if i + 1 < len(args) and isinstance(args[i + 1], str) and args[i + 1]:
                filtered += [args[i], args[i + 1]]
            i += 2

    return filtered if filtered else None


# ---------------------------------------------------------------------------
# Main formatting function
# ---------------------------------------------------------------------------

def format_excel(df, filename, p_children=None, tag_colors=None,
                 rich_text_columns=None):
    """
    Write *df* to *filename* (.xlsx) with:
      - header row (bold, grey background)
      - rich-text colour coding for columns in *rich_text_columns*
      - auto column widths (capped at 100)
      - auto row heights
    
    Parameters
    ----------
    df               : DataFrame to write
    filename         : output path
    p_children       : list of top-level element tags (used for col ordering,
                       not strictly needed here)
    tag_colors       : dict {tag_name: hex_color} — overrides TAG_COLORS
    rich_text_columns: set of column names to render with inline tag colours
                       (default: RICH_TEXT_COLUMNS = {'paragraph_text'})
    """
    tag_colors = tag_colors or TAG_COLORS
    rich_cols = rich_text_columns or RICH_TEXT_COLUMNS

    # Clean up column names
    df.columns = df.columns.astype(str)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.reset_index(drop=True)

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Sheet1')
        writer.sheets['Sheet1'] = worksheet

        regular = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        bold_fmt = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})
        header_fmt = workbook.add_format({'bold': True, 'valign': 'top',
                                          'bg_color': '#F2F2F2'})

        # Headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)

        # Column widths
        col_widths = {col: len(str(col)) for col in df.columns}
        for row_num in range(len(df)):
            for col_num, col_name in enumerate(df.columns):
                val = df.iloc[row_num][col_name]
                if pd.notna(val):
                    clean = re.sub(r'<[^>]+>', '', str(val))
                    col_widths[col_name] = max(col_widths[col_name], len(clean) + 2)
        for col_name in col_widths:
            col_widths[col_name] = min(col_widths[col_name], 100)

        # Row heights
        base_h = 15
        row_heights = {}
        for row_num in range(len(df)):
            max_h = base_h
            for col_name in df.columns:
                val = df.iloc[row_num][col_name]
                if pd.notna(val):
                    clean = re.sub(r'<[^>]+>', '', str(val))
                    w = col_widths[col_name]
                    lines = max(1, math.ceil(len(clean) / max(int(w * 1.1), 1)))
                    max_h = max(max_h, base_h * lines)
            row_heights[row_num + 1] = max_h

        # Set column widths
        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, col_widths[col_name], regular)

        # Write rows
        for row_num in range(len(df)):
            excel_row = row_num + 1
            worksheet.set_row(excel_row, row_heights[excel_row])

            for col_num, col_name in enumerate(df.columns):
                val = df.iloc[row_num][col_name]
                if pd.isna(val):
                    continue
                text = str(val)

                if col_name in rich_cols:
                    rich_args = _build_rich_args(text, workbook, tag_colors)
                    if rich_args:
                        worksheet.write_rich_string(excel_row, col_num, *rich_args)
                    else:
                        worksheet.write(excel_row, col_num, text, regular)
                else:
                    # Bold markup: **text** or <strong>text</strong>
                    text = re.sub(r'\*\*(.*?)\*\*', r'<B>\1</B>', text)
                    text = re.sub(r'<strong>(.*?)</strong>', r'<B>\1</B>', text,
                                  flags=re.IGNORECASE)
                    parts = re.split(r'(<B>|</B>)', text)
                    if len(parts) > 1:
                        rich = []
                        bold_on = False
                        for part in parts:
                            if part == '<B>':
                                bold_on = True
                            elif part == '</B>':
                                bold_on = False
                            elif part:
                                rich += [bold_fmt if bold_on else regular, part]
                        rich = [x for x in rich if x != '']
                        if rich:
                            worksheet.write_rich_string(excel_row, col_num, *rich)
                        else:
                            worksheet.write(excel_row, col_num, text, regular)
                    else:
                        worksheet.write(excel_row, col_num, text, regular)


# ---------------------------------------------------------------------------
# Comparison colouring (kept for evaluation workflows)
# ---------------------------------------------------------------------------

def color_compare_pairs(df, filename):
    """Green/red fill for _human vs _ia column pairs."""
    human_cols = [c for c in df.columns if c.endswith('_human')]
    pairs = [(h, h.replace('_human', '_ia'))
             for h in human_cols if h.replace('_human', '_ia') in df.columns]
    if not pairs:
        return

    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        wb = writer.book
        ws = writer.sheets['Sheet1']
        green = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        red   = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        for row_num in range(len(df)):
            er = row_num + 2
            for hc, ic in pairs:
                hv = str(df.iloc[row_num][hc]) if pd.notna(df.iloc[row_num][hc]) else ''
                iv = str(df.iloc[row_num][ic])  if pd.notna(df.iloc[row_num][ic])  else ''
                fill = green if hv.strip().lower() == iv.strip().lower() else red
                ws.cell(row=er, column=df.columns.get_loc(hc) + 1).fill = fill
                ws.cell(row=er, column=df.columns.get_loc(ic) + 1).fill = fill
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            width = min(max((len(str(c.value)) for c in col if c.value), default=10) + 2, 50)
            ws.column_dimensions[letter].width = width


if __name__ == '__main__':
    import sys
    df = pd.read_excel(sys.argv[1])
    format_excel(df, sys.argv[1].replace('.xlsx', '_formatted.xlsx'))
