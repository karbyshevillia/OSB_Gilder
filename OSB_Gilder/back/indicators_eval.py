import pandas as pd
import re
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import defaultdict


class IndicatorsFrameBuilder:
    # Mapping CSV mask aliases to the DataFrame coordinate columns
    ALIAS_MAP = {
        'DT': 'Debit_Total', 'DNC': 'Debit_NC', 'DIC': 'Debit_IC',
        'CT': 'Credit_Total', 'CNC': 'Credit_NC', 'CIC': 'Credit_IC',
        'BT': 'Balance_Total', 'BNC': 'Balance_NC', 'BIC': 'Balance_IC'
    }

    def __init__(self, indicators_csv, balance_codes_frame):
        # Accepting a pre-loaded DataFrame of coordinates
        self.rules = pd.read_csv(indicators_csv, sep=";", skipinitialspace=True)
        self.balance_df = balance_codes_frame
        self.output_cols = [
            'Debit_Total', 'Debit_NC', 'Debit_IC',
            'Credit_Total', 'Credit_NC', 'Credit_IC',
            'Balance_Total', 'Balance_NC', 'Balance_IC'
        ]

        # ADDED: Mapping for O(1) lookups: ID -> (Row Index in DataFrame, Mask string)
        self.id_meta = {
            int(row['ID']): (idx, str(row['MASK']).upper())
            for idx, row in self.rules.iterrows()
        }

        # NEW: Track which DataFrame indices are used by which Indicator IDs
        self.code_usage = defaultdict(set)

    def _is_active(self, c):
        return str(c).upper() in ('A', 'А')

    def _is_passive(self, c):
        return str(c).upper() in ('P', 'П')

    def _group_into_ranges(self, coords):
        """Converts ['P50', 'P51', 'P53'] into 'SUM(P50:P51, P53)'"""
        if not coords: return "0"
        # Extract column letter and row numbers
        col = re.match(r"([A-Z]+)", coords[0]).group(1)
        rows = sorted([int(re.search(r"\d+", c).group()) for c in coords])

        ranges = []
        if not rows: return "0"

        start = prev = rows[0]
        for r in rows[1:]:
            if r == prev + 1:
                prev = r
            else:
                ranges.append(f"{col}{start}:{col}{prev}" if start != prev else f"{col}{start}")
                start = prev = r
        ranges.append(f"{col}{start}:{col}{prev}" if start != prev else f"{col}{start}")

        return f"SUM({', '.join(ranges)})" if len(ranges) > 1 or ":" in ranges[0] else ranges[0]

    def _compile(self, ind_id, expr, alias, start_col, start_row):
        # Regex Update:
        # 1. Added (?:UE)? to match both CODE_VAL and CODE_VALUE
        # 2. We don't replace X here anymore; we handle it in the resolver
        pattern = r'(CODE_VAL(?:UE)?|CODE_SUM)\(\s*([0-9\*]+[АПA-Pа-пa-p]?)\s*,\s*([A-Z_]+)\s*(?:,\s*([АПA-Pа-пa-p]))?\s*\)'

        def resolver(match):
            func_type, raw_code, target_alias, true_kind_param = match.groups()

            # KEY FIX: If target_alias is 'X', use the alias passed to _compile
            actual_alias = alias if target_alias == 'X' else target_alias

            # Now verify the alias exists in your map
            if actual_alias not in self.ALIAS_MAP:
                # If it's still not found, it's a typo in the CSV (e.g., CODE_VAL(1000, INVALID))
                return "0"

            # 1. Handle Suffix (Kind_Mark)
            last = raw_code[-1]
            if self._is_active(last):
                kind_mark, code = 'A', raw_code[:-1]
            elif self._is_passive(last):
                kind_mark, code = 'П', raw_code[:-1]
            else:
                kind_mark, code = None, raw_code

            # 2. Filter coordinates
            mask = self.balance_df['Balance'].astype(str).str.match(f"^{code.replace('*', '.*')}$")

            if kind_mark:
                km_list = ['A', 'А'] if kind_mark == 'A' else ['P', 'П']
                mask &= self.balance_df['Kind_Mark'].str.upper().isin(km_list)

            if true_kind_param:
                k_str = 'Active' if self._is_active(true_kind_param) else 'Passive'
                mask &= (self.balance_df['Kind'] == k_str)

            # NEW: Track the usage!
            # Since 'mask' is a boolean series, we find the indices of True values
            # and map those row indices to the current Indicator ID.
            matched_indices = self.balance_df[mask].index
            # print(matched_indices)
            for idx in matched_indices:
                # print(idx)
                self.code_usage[idx].add(ind_id)



            # Use actual_alias (which is now 'BT', 'DT', etc.)
            coord_col = f"Coord_{self.ALIAS_MAP[actual_alias]}"
            coords = self.balance_df.loc[mask, coord_col].tolist()

            return self._group_into_ranges(coords)

        res = re.sub(pattern, resolver, expr)

        # ADDED: RESOLVE SYNTHETIC INDICATORS (IND_REF)
        pattern_refs = r'IND_REF\(\s*(\d+)\s*,\s*([A-Z_]+)\s*\)'

        def resolver_refs(match):
            target_id = int(match.group(1))
            target_alias = match.group(2).strip()

            actual_alias = alias if target_alias == 'X' else target_alias

            if target_id not in self.id_meta:
                return "#REF_ID_NOT_FOUND!"

            target_idx, target_mask = self.id_meta[target_id]

            # SAFETY CHECK: If the target cell wasn't generated in its mask, return 0
            if actual_alias not in [m.strip() for m in target_mask.split(',')]:
                return "0"

            col_name = self.ALIAS_MAP.get(actual_alias)
            if not col_name:
                return "#REF_ALIAS_NOT_FOUND!"

            # CALCULATE ABSOLUTE EXCEL COORDINATES
            col_keys = ['ID', 'NAME'] + list(self.ALIAS_MAP.values())
            target_excel_col = start_col + col_keys.index(col_name)
            target_excel_row = start_row + 1 + target_idx

            return f"{get_column_letter(target_excel_col)}{target_excel_row}"

        res = re.sub(pattern_refs, resolver_refs, res)

        return f"={res}" if not res.startswith('=') else res

    def build_frame(self, start_coord="A1"):
        """Loops through rules and produces the final formula table."""
        # ADDED: Parse starting coordinate for IND_REF calculation
        col_match = re.match(r"([A-Z]+)(\d+)", start_coord)
        start_col = column_index_from_string(col_match.group(1))
        start_row = int(col_match.group(2))

        table = []
        for _, rule in self.rules.iterrows():
            row_data = {'ID': rule['ID'], 'NAME': rule['NAME']}
            mask_list = [m.strip() for m in str(rule['MASK']).split(',')]
            ind_id = int(rule['ID'])  # NEW: Extract ID to pass to compiler

            for alias, col_name in self.ALIAS_MAP.items():
                if alias in mask_list:
                    # UPDATE: Pass ind_id into _compile
                    row_data[col_name] = self._compile(ind_id, rule['EXPR'], alias, start_col, start_row)
                else:
                    row_data[col_name] = "-"
            table.append(row_data)

        return pd.DataFrame(table)