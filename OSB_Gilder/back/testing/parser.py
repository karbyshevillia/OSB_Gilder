import pandas as pd
import numpy as np
import re
import formulae


def get_col_letter(col_idx):
    """Converts a 0-based Pandas column index into an Excel column letter (0 -> 'A', 26 -> 'AA')"""
    string = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        string = chr(65 + remainder) + string
    return string


def parse_to_sane(filepath, sheet_name):
    # 1. Fast Load with Calamine
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, dtype=str, engine='calamine')

    # NEW: Lock in the original Excel Row Numbers before we do any slicing!
    # Pandas index is 0-based, Excel is 1-based.
    df_raw['Excel_Row'] = df_raw.index + 1

    # 2. Locate the Anchor
    anchor_mask = df_raw.apply(
        lambda row: row.astype(str).str.contains('Балансові рахунки', na=False, case=False).any(), axis=1)
    if anchor_mask.empty:
        raise ValueError("Could not find data anchor.")
    data_start_row = anchor_mask.idxmax()

    # ==========================================
    # 3. HEADER MERGED-CELL RESOLUTION & MAPPING
    # ==========================================
    headers = df_raw.iloc[:data_start_row].copy().astype(str).apply(lambda x: x.str.lower().str.strip())
    headers.replace({'nan': np.nan, 'none': np.nan, '': np.nan}, inplace=True)

    for idx, row in headers.iterrows():
        if any(kw in row.values for kw in ['дебет', 'кредит', 'сальдо']):
            headers.loc[idx] = row.ffill()

    col_mapping = {}
    col_letters = {}  # NEW: Track the Excel Column letters!

    for col in df_raw.columns:
        if col == 'Excel_Row': continue
        text = " ".join(headers[col].dropna().tolist())

        mapped_name = None
        if 'дебет' in text:
            if 'усього' in text:
                mapped_name = 'Debit_Total'
            elif 'нв' in text:
                mapped_name = 'Debit_NC'
            elif 'ів' in text or 'iв' in text:
                mapped_name = 'Debit_IC'
        elif 'кредит' in text:
            if 'усього' in text:
                mapped_name = 'Credit_Total'
            elif 'нв' in text:
                mapped_name = 'Credit_NC'
            elif 'ів' in text or 'iв' in text:
                mapped_name = 'Credit_IC'
        elif 'сальдо' in text:
            if 'усього' in text:
                mapped_name = 'Balance_Total'
            elif 'нв' in text:
                mapped_name = 'Balance_NC'
            elif 'ів' in text or 'iв' in text:
                mapped_name = 'Balance_IC'

        if mapped_name:
            col_mapping[col] = mapped_name
            col_letters[mapped_name] = get_col_letter(col)

    # ==========================================
    # 4. DATA BLOCK PROCESSING
    # ==========================================
    # Instead of resetting the index, we just copy. The 'Excel_Row' column travels with it safely.
    df_data = df_raw.iloc[data_start_row:].copy()

    # COLUMN 1: "Kind"
    df_data['Kind'] = pd.Series(np.nan, dtype='object', index=df_data.index)
    active_pattern = r'^\s*Актив(и|ні)?\s*$'
    passive_pattern = r'^\s*(Зобов\'язання|Пасив(и|ні)?|Капітал)\s*$'

    def is_exclusive_header_row(row, pattern):
        vals = [str(x).strip() for x in row if pd.notna(x) and str(x).strip().lower() not in ('', 'nan', 'none')]
        text_vals = [v for v in vals if not re.match(r'^-?[\d\s\.,]+$', v)]
        if not text_vals: return False
        return all(re.fullmatch(pattern, v, flags=re.IGNORECASE) for v in text_vals)

    # Note: drop 'Excel_Row' from the exclusivity check so it doesn't trigger false negatives!
    check_data = df_data.drop(columns=['Excel_Row'])
    active_mask = check_data.apply(lambda r: is_exclusive_header_row(r, active_pattern), axis=1)
    passive_mask = check_data.apply(lambda r: is_exclusive_header_row(r, passive_pattern), axis=1)

    df_data.loc[active_mask, 'Kind'] = 'Active'
    df_data.loc[passive_mask, 'Kind'] = 'Passive'
    df_data['Kind'] = df_data['Kind'].ffill()

    # COLUMNS 2-5: "Balance" via Purity Ratio & ffill
    best_balance_col = None
    max_purity = 0

    for col in df_data.columns:
        if col in col_mapping or col in ('Kind', 'Excel_Row'): continue

        sample = df_data[col].dropna().astype(str).str.replace(r'\.0$', '', regex=True)
        sample = sample[sample != 'nan']
        if sample.empty: continue

        four_digit_count = sample.str.match(r'^\d{4}$').sum()
        purity = four_digit_count / len(sample)

        if purity > max_purity and purity > 0.5:
            max_purity = purity
            best_balance_col = col

    if best_balance_col is None:
        raise ValueError("Could not find the Balance (Account Code) column.")

    col_mapping[best_balance_col] = 'Balance'
    col_letters['Balance'] = get_col_letter(best_balance_col)

    df_data[best_balance_col] = df_data[best_balance_col].astype(str).str.replace(r'\.0$', '', regex=True)
    df_data[best_balance_col] = df_data[best_balance_col].replace({'nan': np.nan, 'None': np.nan})
    df_data[best_balance_col] = df_data[best_balance_col].ffill()

    # COLUMN 7: "Kind_Mark" (А/П)
    ap_cols = []
    for col in df_data.columns:
        if col in col_mapping or col in ('Kind', 'Excel_Row'): continue
        sample = df_data[col].dropna().astype(str)
        if sample.str.match(r'^[АПA-P]{1,2}$', flags=re.IGNORECASE).sum() > len(sample) * 0.05:
            ap_cols.append(col)

    if not ap_cols:
        raise ValueError("Could not find the Kind_Mark (А/П) column.")
    col_mapping[ap_cols[0]] = 'Kind_Mark'
    col_letters['Kind_Mark'] = get_col_letter(ap_cols[0])

    # COLUMN 6: "Name"
    best_name_col = None
    max_name_len = 0
    for col in df_data.columns:
        if col not in col_mapping and col not in ('Kind', 'Excel_Row'):
            sample = df_data[col].dropna().astype(str)
            if not sample.empty:
                avg_len = sample.str.len().mean()
                if avg_len > max_name_len:
                    max_name_len = avg_len
                    best_name_col = col

    if best_name_col is not None:
        col_mapping[best_name_col] = 'Name'
        col_letters['Name'] = get_col_letter(best_name_col)

    df_data.rename(columns=col_mapping, inplace=True)

    if 'Name' in df_data.columns:
        df_data['Name'] = df_data['Name'].replace({'nan': np.nan, 'None': np.nan})
        df_data['Name'] = df_data['Name'].ffill()

    # ==========================================
    # 5. FILTERING AND ASSEMBLY
    # ==========================================

    line_items = df_data[
        df_data['Kind_Mark'].astype(str).str.match(r'^[АПA-P]{1,2}$', na=False, flags=re.IGNORECASE)].copy()

    line_items['Class'] = line_items['Balance'].str[0]
    line_items['Division'] = line_items['Balance'].str[:2]
    line_items['Group'] = line_items['Balance'].str[:3]

    kind_series = line_items['Kind'].copy()
    kind_series.loc[line_items['Class'] == '5'] = 'Capital'
    kind_series.loc[line_items['Class'] == '6'] = 'Revenue'
    kind_series.loc[line_items['Class'] == '7'] = 'Expenditures'

    invalid_class_mask = ~line_items['Class'].isin(['1', '2', '3', '4', '5', '6', '7', '9'])
    kind_series.loc[invalid_class_mask] = np.nan
    line_items['Kind'] = kind_series
    line_items['Kind'] = line_items['Kind'].replace({'nan': np.nan, 'None': np.nan})

    final_cols = [
        "Kind", "Class", "Division", "Group", "Balance", "Name", "Kind_Mark",
        "Debit_Total", "Debit_NC", "Debit_IC",
        "Credit_Total", "Credit_NC", "Credit_IC",
        "Balance_Total", "Balance_NC", "Balance_IC"
    ]

    for col in final_cols:
        if col not in line_items.columns:
            line_items[col] = np.nan

    # NEW: Inject Coordinates!
    # We will loop through columns that have an original physical location (like Balance_Total, Name, etc.)
    # and combine their Excel Letter + their Excel Row to create columns like 'Coord_Balance_Total'.
    coord_cols = []
    for col in final_cols:
        if col in col_letters:
            coord_col_name = f"Coord_{col}"
            coord_cols.append(coord_col_name)
            # Combine the tracked Letter (e.g., 'M') with the Row string (e.g., '15') -> 'M15'
            line_items[coord_col_name] = col_letters[col] + line_items['Excel_Row'].astype(str)

    # Assembly: Bring it all together (16 Data cols + Coordinate Cols)
    clean_df = line_items[final_cols + coord_cols].copy()

    for col in final_cols[7:]:  # Safely cast the 9 value columns to floats
        clean_df[col] = pd.to_numeric(clean_df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    return clean_df

if __name__ == '__main__':
    df = parse_to_sane("/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/test_files/TEST_03.xlsx", sheet_name='  2АТ "Укрексімбанк" ')
    pd.set_option('display.max_columns', None)
    print(df.head(25))
    print(df.tail(25))
    formula_str = "=CODE_SUM(1*, BT, A) + CODE_VAL(5000, DT)"
    upd = formulae.compile_formula(formula_str, df, '  2АТ "Укрексімбанк" ', include_sheet_name=False)
    print(upd)
    pass