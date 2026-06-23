import re
import pandas as pd
import math
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
import sys

# Color map: tag name -> hex RGB text color (for write_rich_string, xlsxwriter font color)
TAG_COLORS = {
    'INTRODD':   '#1F4E79',  # dark blue
    'VDD':       '#7030A0',  # purple
    'EXPANSION': '#C55A11',  # dark orange
    'MOD':       '#833C00',  # brown
    'PPI':       '#C00000',  # dark red
    'NONPPI':    '#375623',  # dark green
    'MD':        '#2E75B6',  # medium blue
    'APP':       '#595959',  # dark grey
    'DD': '#1D6B5E',  # teal foncé
}
TAG_COLOR_DEFAULT = '#000000'  # black for unknown tags


def _parse_tagged_text(text):
    """
    Parse a string containing XML-like tags and return a list of segments:
    [{'type': 'text'|'open_tag'|'close_tag', 'value': str, 'tag': str or None}]
    """
    pattern = re.compile(r'(</?(\w+)>)')
    segments = []
    pos = 0
    for m in pattern.finditer(text):
        start, end = m.start(), m.end()
        if pos < start:
            segments.append({'type': 'text', 'value': text[pos:start], 'tag': None})
        full_tag = m.group(1)
        tag_name = m.group(2)
        is_close = full_tag.startswith('</')
        segments.append({
            'type': 'close_tag' if is_close else 'open_tag',
            'value': full_tag,
            'tag': tag_name
        })
        pos = end
    if pos < len(text):
        segments.append({'type': 'text', 'value': text[pos:], 'tag': None})
    return segments


def _build_rich_string_args(text, workbook):
    """
    Build args list for worksheet.write_rich_string() from a tagged paragraph_text.
    Tags are shown in grey italic, tag content is colored per TAG_COLORS.
    Plain text (outside any tag) is black.
    Returns None if the text has no tags (use plain write instead).
    """
    segments = _parse_tagged_text(text)
    if not any(s['type'] in ('open_tag', 'close_tag') for s in segments):
        return None

    tag_fmt_cache = {}
    plain_fmt = workbook.add_format({'font_color': '#000000', 'text_wrap': True, 'valign': 'top'})
    tag_label_fmt = workbook.add_format({'font_color': '#AAAAAA', 'italic': True, 'text_wrap': True, 'valign': 'top'})

    def get_tag_fmt(tag_name):
        if tag_name not in tag_fmt_cache:
            color = TAG_COLORS.get(tag_name, TAG_COLOR_DEFAULT)
            tag_fmt_cache[tag_name] = workbook.add_format({
                'font_color': color,
                'text_wrap': True,
                'valign': 'top',
                'bold': True, 
            })
        return tag_fmt_cache[tag_name]

    args = []
    # Track current open tag for coloring content
    tag_stack = []

    for seg in segments:
        if seg['type'] == 'open_tag':
            tag_stack.append(seg['tag'])
            args.append(tag_label_fmt)
            args.append(seg['value'])
        elif seg['type'] == 'close_tag':
            args.append(tag_label_fmt)
            args.append(seg['value'])
            if tag_stack and tag_stack[-1] == seg['tag']:
                tag_stack.pop()
        else:  # plain text
            val = seg['value']
            if not val:
                continue
            if tag_stack:
                # Color by innermost tag
                args.append(get_tag_fmt(tag_stack[-1]))
            else:
                args.append(plain_fmt)
            args.append(val)

    # write_rich_string needs at least one format+string pair
    # Filter empty strings
    filtered = []
    i = 0
    while i < len(args):
        if isinstance(args[i], str):
            if args[i]:
                filtered.append(args[i])
            i += 1
        else:
            # it's a format object, pair with next string
            if i + 1 < len(args) and isinstance(args[i+1], str) and args[i+1]:
                filtered.append(args[i])
                filtered.append(args[i+1])
            i += 2

    if not filtered:
        return None
    return filtered


def format_ppi_bold(df, filename):
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(
                lambda v: v.replace('–', '\n–').replace('\n\n', '\n') if isinstance(v, str) else v
            )

    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.reset_index(drop=True)

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Sheet1')
        writer.sheets['Sheet1'] = worksheet

        bold = workbook.add_format({'bold': True})
        regular = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        header_format = workbook.add_format({'bold': True, 'valign': 'top', 'bg_color': '#F2F2F2'})

        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_format)

        # Calculate column widths
        col_widths = {col: len(str(col)) for col in df.columns}
        for row_num in range(len(df)):
            for col_num, col_name in enumerate(df.columns):
                cell_value = df.iloc[row_num][col_name]
                if pd.notna(cell_value):
                    text = str(cell_value)
                    if not text.strip():
                        continue
                    clean = re.sub(r'<[^>]+>', '', text)
                    clean = re.sub(r'\*\*', '', clean)
                    col_widths[col_name] = max(col_widths[col_name], len(clean) + 2)

        for col_name in col_widths:
            col_widths[col_name] = min(col_widths[col_name], 100)

        # Calculate row heights
        base_height = 15
        row_heights = {}
        for row_num in range(len(df)):
            excel_row = row_num + 1
            max_height = base_height
            for col_num, col_name in enumerate(df.columns):
                cell_value = df.iloc[row_num][col_name]
                if pd.notna(cell_value):
                    text = str(cell_value)
                    clean = re.sub(r'<[^>]+>', '', text)
                    clean = re.sub(r'\*\*', '', clean)
                    col_width = col_widths[col_name]
                    chars_per_line = int(col_width * 1.1)
                    lines_needed = max(1, math.ceil(len(clean) / chars_per_line)) if chars_per_line > 0 else 1
                    max_height = max(max_height, base_height * lines_needed)
            row_heights[excel_row] = max_height

        # Set column widths
        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, col_widths[col_name], regular)

        # Write content
        for row_num in range(len(df)):
            excel_row = row_num + 1
            worksheet.set_row(excel_row, row_heights[excel_row])

            for col_num, col_name in enumerate(df.columns):
                cell_value = df.iloc[row_num][col_name]

                if pd.notna(cell_value):
                    text = str(cell_value)

                    if col_name in ['paragraph_text','paragraph_text_dd']:
                        # Rich string with tag color coding
                        rich_args = _build_rich_string_args(text, workbook)
                        if rich_args:
                            worksheet.write_rich_string(excel_row, col_num, *rich_args)
                        else:
                            worksheet.write(excel_row, col_num, text, regular)
                    else:
                        # Existing PPI bold logic for other columns
                        text = re.sub(r'\*\*(.*?)\*\*', r'<PPI>\1</PPI>', text)
                        text = re.sub(r'<strong>(.*?)</strong>', r'<PPI>\1</PPI>', text)
                        parts = re.split(r'(<PPI>|</PPI>)', text)

                        if len(parts) > 1:
                            rich_string = []
                            is_bold = False
                            for part in parts:
                                if part == '<PPI>':
                                    is_bold = True
                                    rich_string.append('<PPI>')
                                elif part == '</PPI>':
                                    is_bold = False
                                    rich_string.append('</PPI>')
                                else:
                                    if part:
                                        if is_bold:
                                            rich_string.append(bold)
                                            rich_string.append(part)
                                        else:
                                            rich_string.append(part)
                            rich_string = [item for item in rich_string if item != '']
                            if rich_string:
                                worksheet.write_rich_string(excel_row, col_num, *rich_string)
                            else:
                                worksheet.write(excel_row, col_num, text, regular)
                        else:
                            worksheet.write(excel_row, col_num, text, regular)


def color_compare_pairs(df, filename):
    df_copy = df.copy()
    human_cols = [col for col in df_copy.columns if col.endswith('_human')]
    pairs = []
    for human_col in human_cols:
        ia_col = human_col.replace('_human', '_ia')
        if ia_col in df_copy.columns:
            pairs.append((human_col, ia_col))

    if not pairs:
        print("No '_human' and '_ia' column pairs found")
        return

    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        df_copy.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

        for row_num in range(len(df_copy)):
            excel_row = row_num + 2
            for human_col, ia_col in pairs:
                human_val = df_copy.iloc[row_num][human_col]
                ia_val = df_copy.iloc[row_num][ia_col]
                human_str = str(human_val) if pd.notna(human_val) else ""
                ia_str = str(ia_val) if pd.notna(ia_val) else ""
                are_equal = human_str.strip().lower() == ia_str.strip().lower()
                human_col_idx = df_copy.columns.get_loc(human_col)
                ia_col_idx = df_copy.columns.get_loc(ia_col)
                fill = green_fill if are_equal else red_fill
                worksheet.cell(row=excel_row, column=human_col_idx + 1).fill = fill
                worksheet.cell(row=excel_row, column=ia_col_idx + 1).fill = fill

        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)


def format_and_compare(df, filename):
    format_ppi_bold(df, filename)
    temp_df = pd.read_excel(filename)
    color_compare_pairs(temp_df, filename)


if __name__ == "__main__":
    df = pd.read_excel(sys.argv[1])
    format_ppi_bold(df, sys.argv[1].replace(".xlsx", "_formatted.xlsx"))
