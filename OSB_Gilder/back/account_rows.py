import openpyxl as xl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import Cell
import pandas as pd
from .indicators_eval import Indicator, IndicatorKind, LogicalIndicator
from .script.utils import STATE_BANKS, FOREIGN_BANKS

import re


class CodedBlock:
    def __init__(self, code: str, df: pd.DataFrame):
        self.code = code
        self.df = df

    def __repr__(self):
        # return f"CodedBlock(code={self.code!r}, rows={len(self.df)})"
        return str(self.df)


class BalanceSheet:
    PARAMETER_NAMES = ["id", "name", "category",
                       "debit_total", "debit_nc", "debit_ic",
                       "credit_total", "credit_nc", "credit_ic",
                       "balance_total", "balance_nc", "balance_ic"]
    LOGICAL_COLUMNS = PARAMETER_NAMES[3:]
    INDICATORS = [LogicalIndicator("Проценти по ОВДП (коди 6120-6122)", IndicatorKind.PRIMARY, "=SUM(():())", 2,
                                   ("6120", "6122+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ("6120", "6120+", "6121", "6121+", "6122", "6122+")),
                  LogicalIndicator("Проценти по ДС (коди 6127-6128)", IndicatorKind.PRIMARY, "=SUM(():())", 2,
                                   ("6127", "6128+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ("6127", "6127+", "6128", "6128+")),
                  LogicalIndicator("Обсяг ДС дебит (середні, поділено на кількість банк.днів) (код 1430, 1440)",
                                   IndicatorKind.PRIMARY, "=()+()", 2, ("1430", "1440"),
                                   ("debit_total", "debit_nc", "debit_ic",),
                                   ("1430", "1430+", "1440", "1440+")),
                  LogicalIndicator("Обсяг ДС кредит (середні, поділено на кількість банк.днів) (код 1430, 1440)",
                                   IndicatorKind.PRIMARY, "=()+()", 2, ("1430", "1440"),
                                   ("credit_total", "credit_nc", "credit_ic",),
                                   ("1430", "1430+", "1440", "1440+")),
                  LogicalIndicator("Кошти в НБУ (кор.рахунок) (кредит) (код 1200)", IndicatorKind.PRIMARY, "=()", 1,
                                   ("1200",),
                                   ("credit_total", "credit_nc", "credit_ic",),
                                   ("1200", "1200+")),
                  LogicalIndicator("Кошти в НБУ (УСЬОГО) (кредит) (розділ 12)", IndicatorKind.PRIMARY, "=()+()", 2,
                                   ("1200", "1208"),
                                   ("credit_total", "credit_nc", "credit_ic",),
                                   ("1200", "1200+", "1208", "1208+")),
                  LogicalIndicator("Процентні доходи (коди р.60,р.61)", IndicatorKind.PRIMARY,
                                   "=()+()+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+()+()",
                                   22,
                                   ("6000", "6000+", "6010", "6015+", "6020", "6027+", "6033", "6035+", "6040", "6041+",
                                    "6050", "6055+", "6060", "6063+", "6090", "6094+", "6110", "6113+", "6120", "6128+",
                                    "6141", "6141+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(6000, 6200)] + [f"{i}+" for i in range(6000, 6200)]),
                  LogicalIndicator("Проценти по ДС у % до процентних доходів", IndicatorKind.SECONDARY, "=()/()*100", 2,
                                   ("Проценти по ДС (коди 6127-6128)", "Процентні доходи (коди р.60,р.61)"),
                                   ("balance_total", "balance_nc", "balance_ic",), []),
                  LogicalIndicator("Проценти по ОВДП у % до процентних доходів", IndicatorKind.SECONDARY, "=()/()*100",
                                   2, ("Проценти по ОВДП (коди 6120-6122)", "Процентні доходи (коди р.60,р.61)"),
                                   ("balance_total", "balance_nc", "balance_ic",), []),
                  LogicalIndicator("Проценти по ДС+ОВДП у % до процентних доходів", IndicatorKind.SECONDARY, "=()+()",
                                   2, ("Проценти по ДС у % до процентних доходів",
                                       "Проценти по ОВДП у % до процентних доходів"),
                                   ("balance_total", "balance_nc", "balance_ic",), []),
                  LogicalIndicator("Проценти за кредитами НК (коди 602, 603, 609)", IndicatorKind.PRIMARY,
                                   "=SUM(():())+SUM(():())+SUM(():())", 6,
                                   ("6020", "6027+", "6033", "6035+", "6090", "6094+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"602{i}" for i in range(0, 10)] + [f"602{i}+" for i in range(0, 10)]
                                   + [f"603{i}" for i in range(0, 10)] + [f"603{i}+" for i in range(0, 10)]
                                   + [f"609{i}" for i in range(0, 10)] + [f"609{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Проценти за кредитами ДГ (коди 605, 606, 611)", IndicatorKind.PRIMARY,
                                   "=SUM(():())+SUM(():())+SUM(():())", 6,
                                   ("6050", "6055+", "6060", "6063+", "6110", "6113+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"605{i}" for i in range(0, 10)] + [f"605{i}+" for i in range(0, 10)]
                                   + [f"606{i}" for i in range(0, 10)] + [f"606{i}+" for i in range(0, 10)]
                                   + [f"611{i}" for i in range(0, 10)] + [f"611{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Проценти за кредитами ЗДУ (код 604)", IndicatorKind.PRIMARY, "=SUM(():())", 2,
                                   ("6040", "6041+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"604{i}" for i in range(0, 10)] + [f"604{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Проценти за поточними вкладами ДГ (код 7040)", IndicatorKind.PRIMARY, "=()+()", 2,
                                   ("7040", "7040+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ["7040", "7040+"]),
                  LogicalIndicator("Проценти за строковими вкладами ДГ (код 7041)", IndicatorKind.PRIMARY, "=()+()", 2,
                                   ("7041", "7041+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ["7041", "7041+"]),
                  LogicalIndicator("Процентні витрати банків ДГ - всього ", IndicatorKind.SECONDARY, "=()+()", 2, (
                  "Проценти за поточними вкладами ДГ (код 7040)", "Проценти за строковими вкладами ДГ (код 7041)"),
                                   ("balance_total", "balance_nc", "balance_ic",), []),
                  LogicalIndicator("Кредити субʼєктам господарювання (р.20, р.23, 260)", IndicatorKind.PRIMARY,
                                   "=SUM(():())+SUM(():())+SUM(():())", 6,
                                   ("2010", "2089+", "2301", "2398+", "2600", "2609+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(2000, 2100)] + [f"{i}+" for i in range(2000, 2100)]
                                   + [f"{i}" for i in range(2300, 2400)] + [f"{i}+" for i in range(2300, 2400)]
                                   + [f"260{i}" for i in range(0, 10)] + [f"260{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Кредити ЗДУ (р.21)", IndicatorKind.PRIMARY, "=SUM(():())", 2, ("2103", "2149+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(2100, 2200)] + [f"{i}+" for i in range(2100, 2200)]),
                  LogicalIndicator("Кредити фізичним особам (р. 22, 24, 262)", IndicatorKind.PRIMARY,
                                   "=SUM(():())+SUM(():())+SUM(():())", 6,
                                   ("2203", "2249+", "2450", "2458+", "2620", "2629+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(2200, 2300)] + [f"{i}+" for i in range(2200, 2300)]
                                   + [f"{i}" for i in range(2400, 2500)] + [f"{i}+" for i in range(2400, 2500)]
                                   + [f"262{i}" for i in range(0, 10)] + [f"262{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Кредити небанківським установам (265)", IndicatorKind.PRIMARY, "=SUM(():())", 2,
                                   ("2650", "2659+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"265{i}+" for i in range(0, 10)] + [f"265{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Прибуток звітного року (5040)", IndicatorKind.PRIMARY, "=()", 2, ("5040",),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ["5040", "5040+"]),
                  LogicalIndicator("Збиток звітного року (5041)", IndicatorKind.PRIMARY, "=()", 2, ("5041",),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ["5041", "5041+"]),
                  LogicalIndicator("Податок на прибуток (7900)", IndicatorKind.PRIMARY, "=()+()", 2, ("7900", "7900+",),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   ["7900", "7900+"]),
                  LogicalIndicator("Інші податки (741)", IndicatorKind.PRIMARY, "=SUM(():())", 2, ("7410", "7419+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"741{i}" for i in range(0, 10)] + [f"741{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Відрахування в резерви (770)", IndicatorKind.PRIMARY, "=SUM(():())", 2,
                                   ("7700", "7707+"),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"770{i}" for i in range(0, 10)] + [f"770{i}+" for i in range(0, 10)]),
                  LogicalIndicator("Загальні доходи (6)", IndicatorKind.PRIMARY,
                                   "=SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+()+()+()+()+()+SUM(():())+SUM(():())+SUM(():())+SUM(():())+()+SUM(():())",
                                   46,
                                   ("6000", "6000+", "6010", "6015+", "6020", "6027+", "6033", "6035+",
                                    "6040", "6041+", "6050", "6055+", "6060", "6063+", "6090", "6094+",
                                    "6110", "6113+", "6120", "6128+", "6141", "6141+", "6201", "6209+",
                                    "6214", "6219+", "6223", "6226+", "6300", "6303+", "6320", "6330",
                                    "6340", "6350", "6360", "6390", "6399+", "6490", "6499+", "6500",
                                    "6509+", "6510", "6519+", "6520", "6711", "6717+",),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(6000, 7000)] + [f"{i}+" for i in range(6000, 7000)]),
                  LogicalIndicator("Загальні витрати (7)", IndicatorKind.PRIMARY,
                                   "=()+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+()+()+()+()+()+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+SUM(():())+()+SUM(():())+()",
                                   42,
                                   ("7003", "7010", "7017+", "7020", "7028+", "7040", "7048+", "7060",
                                    "7060+", "7070", "7071+", "7120", "7122+", "7140", "7142+", "7300",
                                    "7301+", "7320", "7330", "7340", "7350", "7360", "7390", "7399+",
                                    "7400", "7409+", "7410", "7419+", "7420", "7424+", "7430", "7433+",
                                    "7450", "7457+", "7490", "7499+", "7500", "7509+", "7520", "7700",
                                    "7707+", "7900",),
                                   ("balance_total", "balance_nc", "balance_ic",),
                                   [f"{i}" for i in range(7000, 8000)] + [f"{i}+" for i in range(7000, 8000)]),
                  ]

    def __init__(self, parent_file, sheet):
        self.sheet = sheet
        self.parent_file = parent_file
        # self.resolve_column_mergers()
        self.coded_blocks = self.extract_coded_blocks(start_row=self.find_first_occurrence("1001").row,
                                                      code_col=self.find_first_occurrence("1001").column)
        self.start_row = self.sheet.max_row + 3
        self.start_col = column_index_from_string(self.column_parity["category"])
        self.build_indicator_dataframe(BalanceSheet.INDICATORS,
                                       BalanceSheet.LOGICAL_COLUMNS,
                                       self.start_row + 1,
                                       self.start_col)  # self.indicator_frame
        # self.insert_frame(self.parent_file,
        #                   start_row=self.start_row,
        #                   start_col=self.start_col)  # self.indicator_frame_cells
        self.create_pivot_db()  # self.pivot_db

    @property
    def column_parity(self):
        sieved = self.account_row_sieve(self.find_first_occurrence("1001"))
        column_letters = BalanceSheet.column_letters_from_row(sieved)
        return dict(zip(BalanceSheet.PARAMETER_NAMES, column_letters))

    def find_first_occurrence(self, target: str):
        """
        Finds the first occurrence of the target string
        in a given openpyxl-opened sheet; rows take
        precedence over columns
        :param sheet:
        :param target:
        :return:
        """
        for row in self.sheet.iter_rows():
            for cell in row:
                if cell.value == target:
                    return cell
        return None

    @staticmethod
    def contains_id(start_cell):
        """
        Checks if a given cell contains
        an account id, i.e. a 4-digit integer
        :param start_cell:
        :return:
        """
        if not start_cell.value:
            return False
        val = start_cell.value.strip()
        if len(val) == 4:
            try:
                int(val)
            except ValueError:
                return False
        else:
            return False
        return True

    @staticmethod
    def is_data_row(lst):
        try:
            return {"А", "П"} & {cell.value for cell in lst}
        except:
            return False

    @staticmethod
    def column_letters_from_row(cells: list):
        return [cell.column_letter for cell in cells]

    def construct_row_by_template(self, column_letters: list, row_number):
        return [self.sheet[f"{letter}{row_number}"] for letter in column_letters]

    def account_row_sieve(self, start_cell):
        col, row = start_cell.column, start_cell.row
        account_row_cells = [cell for cell in self.sheet[row][col - 1:] if
                             (cell.value is not None and cell.value != "")]
        return account_row_cells

    def extract_coded_blocks(
            self,
            start_row: int,
            end_row: int | None = None,
            code_col: int = 1,  # 1-based (Excel style)
    ) -> dict[CodedBlock]:
        blocks = {}

        current_code = None
        current_rows = []

        for row in self.sheet.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=code_col,
                values_only=False,
        ):
            code_cell = row[0]
            code_value = code_cell.value

            if self.contains_id(code_cell):
                if current_code is not None:
                    df = pd.DataFrame(current_rows, dtype=object, columns=BalanceSheet.PARAMETER_NAMES)
                    if current_code not in blocks.keys():
                        blocks[current_code] = CodedBlock(current_code, df)
                    else:
                        appended = pd.concat([blocks[current_code].df, df], ignore_index=True)
                        blocks[current_code] = CodedBlock(current_code, appended)

                sieved = self.account_row_sieve(code_cell)
                temp = BalanceSheet.column_letters_from_row(sieved)
                current_code = code_value
                current_rows = [sieved]

            else:
                if (current_code is not None) and BalanceSheet.is_data_row(row):
                    row_number = row[0].row
                    current_rows.append(self.construct_row_by_template(temp, row_number))

        return blocks

    def build_indicator_dataframe(
            self,
            indicators,
            logical_columns,
            start_row,
            start_col
    ):
        df = pd.DataFrame(
            index=[f.name for f in indicators],
            columns=logical_columns,
            dtype=object,
        )

        indicator_index = {f.name: i for i, f in enumerate(indicators)}

        for row_idx, indicator in enumerate(indicators):
            for i, col in enumerate(logical_columns):
                if col not in indicator.mask:
                    df.loc[indicator.name, col] = None
                    continue
                ind = Indicator(indicator.skeleton)

                if indicator.kind is IndicatorKind.PRIMARY:
                    inputs = []
                    for code_like in indicator.inputs:
                        code, aux = code_like[:-1], code_like[-1]
                        if aux == "+":
                            row = self.coded_blocks[code].df.iloc[-1, 0].row
                        else:
                            code = code_like
                            row = self.coded_blocks[code].df.iloc[0, 0].row
                        inputs.append(f"{self.column_parity[col]}{row}")

                else:  # SECONDARY
                    # print(start_col, i, get_column_letter(start_col + i))
                    inputs = [
                        f"{get_column_letter(start_col + i + 1)}"
                        f"{start_row + indicator_index[name]}"
                        for name in indicator.inputs
                    ]

                df.loc[indicator.name, col] = ind.render(*inputs)

        self.indicator_frame = df

        return

    def insert_frame(self, parent_file, start_row, start_col):
        """Insert indicator frame directly to the sheet using openpyxl (no pandas ExcelWriter)"""
        from openpyxl.cell.cell import MergedCell
        from openpyxl.styles import Font, PatternFill, Border, Side

        # Create DataFrame with actual cells (handle merged cells)
        def get_actual_cell(cell):
            if isinstance(cell, MergedCell):
                for merged_range in self.sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        min_col, min_row, max_col, max_row = merged_range.bounds
                        return self.sheet.cell(row=min_row, column=min_col)
            return cell

        df_cells = pd.DataFrame([[get_actual_cell(self.sheet.cell(i, j))
                                  for j in range(start_col + 1, start_col + 1 + len(self.indicator_frame.columns))]
                                 for i in range(start_row + 1, start_row + 1 + len(self.indicator_frame.index))],
                                index=self.indicator_frame.index,
                                columns=self.indicator_frame.columns,
                                dtype=object)

        # Unmerge cells in the target range
        end_row = start_row + len(self.indicator_frame.index) + 1
        end_col = start_col + len(self.indicator_frame.columns) + 1

        merged_to_remove = []
        for merged_range in list(self.sheet.merged_cells.ranges):
            min_col, min_row, max_col, max_row = merged_range.bounds
            if (min_row <= end_row and max_row >= start_row and
                    min_col <= end_col and max_col >= start_col):
                merged_to_remove.append(merged_range)

        for merged_range in merged_to_remove:
            self.sheet.unmerge_cells(str(merged_range))

        # Define styles for headers (using openpyxl, not pandas)
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # lightgray
        header_font = Font(bold=True)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Write column headers
        for col_idx, col_name in enumerate(self.indicator_frame.columns):
            cell = self.sheet.cell(row=start_row, column=start_col + col_idx + 1)
            cell.value = col_name
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border

        # Write index header
        index_cell = self.sheet.cell(row=start_row, column=start_col)
        index_cell.value = "Index"
        index_cell.font = header_font
        index_cell.fill = header_fill
        index_cell.border = border

        # Write data rows directly to the sheet
        for row_idx, (index_name, row_data) in enumerate(self.indicator_frame.iterrows()):
            # Alternate background colors
            row_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3",
                                   fill_type="solid") if row_idx % 2 == 0 else None

            # Write index name
            index_cell = self.sheet.cell(row=start_row + row_idx + 1, column=start_col)
            index_cell.value = index_name
            index_cell.border = border
            if row_fill:
                index_cell.fill = row_fill

            # Write row data (values)
            for col_idx, value in enumerate(row_data):
                cell = self.sheet.cell(row=start_row + row_idx + 1, column=start_col + col_idx + 1)
                cell.value = value
                cell.border = border
                if row_fill:
                    cell.fill = row_fill

        self.indicator_frame_cells = df_cells

    def create_pivot_db(self):
        rows = []

        for code, block in self.coded_blocks.items():
            df = block.df
            for i, (_, row) in enumerate(df.iterrows()):

                # convert row to plain values
                row_dict = {
                    k: (v.value if isinstance(v, Cell) else v)
                    for k, v in row.items()
                }

                row_dict["id"] = code
                name = df.iloc[0, 1].value
                row_dict["name"] = name

                base = row_dict.copy()
                base["indicator"] = "(без привʼязки)"
                rows.append(base)

                for ind in BalanceSheet.INDICATORS:
                    if ind.kind == IndicatorKind.PRIMARY:
                        for c in ind.associated_codes:
                            target_code = c.rstrip("+")
                            target_row = 1 if c.endswith("+") else 0

                            if target_code == code and target_row == i:
                                new_row = row_dict.copy()
                                new_row["indicator"] = ind.name
                                rows.append(new_row)

        pivot_db = pd.DataFrame(rows)

        match = re.match(r'^\s*(\d+)', self.sheet.title)
        number_str = match.group(1) if match else ""
        text_part = self.sheet.title[match.end():].strip() if match else self.sheet.title

        def _bank_class(bank_name):
            b_name = bank_name.lower()
            if b_name in STATE_BANKS:
                return "Державний"
            elif b_name in FOREIGN_BANKS:
                return "Іноземний капітал"
            else:
                return "Приватний капітал"

        pivot_db.insert(0, "bank_class", _bank_class(text_part), True)
        pivot_db.insert(0, "bank_name", text_part, True)
        pivot_db.insert(0, "bank_code", number_str, True)

        indicator_col = pivot_db.pop("indicator")
        pivot_db.insert(4, "indicator", indicator_col)  # 4 = desired column position

        self.pivot_db = pivot_db


if __name__ == '__main__':
    wb = xl.load_workbook("Untitled_2.xlsx")
    sheet = BalanceSheet(sheet=wb.worksheets[1], parent_file="Untitled_2.xlsx")
    # start_cell = sheet.find_first_occurrence("1001")
    # print(start_cell)
    # # print(sheet.create_account_row(start_cell, show_values=False))
    # # df = pd.DataFrame(sheet.account_rows)
    # df = sheet.coded_blocks
    # pd.set_option("display.max_columns", None)
    # pd.set_option("display.max_rows", None)
    # print(df)
    # # print(sheet.get_account_row_info("1001"))
    # print(sheet.indicator_frame)
    # # sheet.insert_frame("../Untitled_1.xlsx")
    # print(sheet.indicator_frame_cells)
    # # print(generate_account_rows(sheet))
    pd.set_option("display.max_columns", None)
    # pd.set_option("display.max_rows", None)
    # print(sheet.coded_blocks)
    print(sheet.coded_blocks["1405"])
    # print(sheet.sheet_db)
    print(sheet.pivot_db)

    # df = pd.DataFrame({"A": [4, 4], "B": [9, 9]})
    # new_df = df.apply(lambda x: x**0.5)
    # print(df, new_df)