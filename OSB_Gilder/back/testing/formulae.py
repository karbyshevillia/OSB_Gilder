import re

ALIAS_MAP = {
    'DT': 'Debit_Total', 'DNC': 'Debit_NC', 'DIC': 'Debit_IC',
    'CT': 'Credit_Total', 'CNC': 'Credit_NC', 'CIC': 'Credit_IC',
    'BT': 'Balance_Total', 'BNC': 'Balance_NC', 'BIC': 'Balance_IC'
}


# Helpers to seamlessly handle both Latin and Cyrillic keyboard layouts
def is_active_char(c):
    return c in ('A', 'А', 'a', 'а')


def is_passive_char(c):
    return c in ('P', 'П', 'p', 'п')


def group_into_ranges(coords, sheet_name="RawData", include_sheet_name=True):
    """
    Takes a list of coordinates like ['P50', 'P51', 'P53', 'P54']
    and groups them into Excel-friendly ranges.
    """
    if not coords:
        return "0"

    col_match = re.match(r"([A-Z]+)(\d+)", coords[0])
    col_letter = col_match.group(1)

    rows = sorted([int(re.search(r"\d+", c).group()) for c in coords])

    ranges = []
    start = rows[0]
    prev = rows[0]

    for r in rows[1:]:
        if r == prev + 1:
            prev = r
        else:
            ranges.append((start, prev))
            start = r
            prev = r
    ranges.append((start, prev))

    sheet_prefix = f"'{sheet_name}'!" if sheet_name and include_sheet_name else ""
    formatted_ranges = []

    for start_row, end_row in ranges:
        if start_row == end_row:
            formatted_ranges.append(f"{sheet_prefix}{col_letter}{start_row}")
        else:
            formatted_ranges.append(f"{sheet_prefix}{col_letter}{start_row}:{col_letter}{end_row}")

    if len(formatted_ranges) == 1 and ":" not in formatted_ranges[0]:
        return formatted_ranges[0]

    return f"SUM({', '.join(formatted_ranges)})"


def compile_formula(formula_str, clean_df, sheet_name="RawData", include_sheet_name=True):
    compiled_str = formula_str

    # REGEX UPDATE:
    # Group 2 now captures the digits/wildcards PLUS an optional A/P at the end.
    # Group 4 captures the optional third parameter for true Kind.
    pattern = r'(CODE_VAL|CODE_SUM)\(\s*([0-9\*]+[АПA-Pа-пa-p]?)\s*,\s*([A-Z_]+)\s*(?:,\s*([АПA-Pа-пa-p]))?\s*\)'
    matches = re.finditer(pattern, formula_str)

    for match in matches:
        full_token = match.group(0)
        func_type = match.group(1)
        target_code_raw = match.group(2)  # e.g., "1405A", "14*", "2600П"
        alias = match.group(3)
        true_kind_raw = match.group(4)  # e.g., "A", "П", or None

        target_col = f"Coord_{ALIAS_MAP.get(alias, alias)}"

        # 1. Parse the Code String for the Kind_Mark (Line Item Split)
        last_char = target_code_raw[-1]

        if is_active_char(last_char):
            kind_mark_filter = 'A'
            base_code = target_code_raw[:-1]
        elif is_passive_char(last_char):
            kind_mark_filter = 'П'
            base_code = target_code_raw[:-1]
        else:
            kind_mark_filter = None
            base_code = target_code_raw

        # 2. Filter by Balance Code
        if '*' in base_code:
            search_pattern = f"^{base_code.replace('*', '.*')}$"
            mask = clean_df['Balance'].astype(str).str.match(search_pattern)
        else:
            mask = clean_df['Balance'].astype(str) == base_code

        # 3. Filter by Kind_Mark (Physical Row Split)
        if kind_mark_filter:
            if kind_mark_filter == 'A':
                mask = mask & clean_df['Kind_Mark'].astype(str).str.strip().str.upper().isin(['A', 'А'])
            else:
                mask = mask & clean_df['Kind_Mark'].astype(str).str.strip().str.upper().isin(['P', 'П'])

        # 4. Filter by True Kind (Third Parameter)
        if true_kind_raw:
            if is_active_char(true_kind_raw):
                mask = mask & (clean_df['Kind'] == 'Active')
            elif is_passive_char(true_kind_raw):
                mask = mask & (clean_df['Kind'] == 'Passive')

        matching_rows = clean_df[mask]

        # 5. Generate Excel Syntax
        if matching_rows.empty:
            excel_syntax = "0"
        else:
            coords = matching_rows[target_col].tolist()
            excel_syntax = group_into_ranges(coords, sheet_name, include_sheet_name)

        compiled_str = compiled_str.replace(full_token, excel_syntax)

    if not compiled_str.startswith('='):
        compiled_str = '=' + compiled_str

    return compiled_str
