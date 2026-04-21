from .utils import *
from .utils import _datalabel_text_style, _normalize_name
from .indicator_tables_creator import IndicatorTablesCreator
from ..account_rows import BalanceSheet

class IndexSheetCreator:
    INDICATOR_CORRESPONDENCE = {
        "ПРОЦЕНТИ ЗА ДЕПОЗИТНИМИ СЕРТИФІКАТАМИ (ДС), тис. грн": ("Проценти по ДС (коди 6127-6128)", 7),
        "ПРОЦЕНТНІ ДОХОДИ ВСЬОГО (ПД), тис. грн": ("Процентні доходи (коди р.60,р.61)", 7),
        "ПРОЦЕНТИ ЗА ОВДП (ОВДП), тис.грн": ("Проценти по ОВДП (коди 6120-6122)", 7),
        "ПРОЦЕНТИ ЗА КРЕДИТАМИ БІЗНЕСУ": ("Проценти за кредитами НК (коди 602, 603, 609)", 7),
        "ПРОЦЕНТИ ЗА КРЕДИТАМИ НАСЕЛЕННЮ": ("Проценти за кредитами ДГ (коди 605, 606, 611)", 7),
        "ПРОЦЕНТИ  ЗА КРЕДИТАМИ ЗДУ": ("Проценти за кредитами ЗДУ (код 604)", 7),
        "ПРОЦЕНТНІ ВИТРАТИ БАНКІВ ЗА ДЕПОЗИТАМИ ТА НАСЕЛЕННЯ, тис.грн": ("Процентні витрати банків ДГ - всього ", 7),
        "ДС/ПД, %": ("Проценти по ДС у % до процентних доходів", 7),
        "(ДС+ОВДП)/ПД, %": ("Проценти по ДС+ОВДП у % до процентних доходів", 7),
        "ОВДП/ПД, %": ("Проценти по ОВДП у % до процентних доходів", 7),
        "Обсяг ДС дебит (середні, поділено на кількість банк.днів)": (
            "Обсяг ДС дебит (середні, поділено на кількість банк.днів) (код 1430, 1440)", 1),
        "Обсяг ДС кредит (середні, поділено на кількість банк.днів)": (
            "Обсяг ДС кредит (середні, поділено на кількість банк.днів) (код 1430, 1440)", 5),
        "Кошти в НБУ (кор.рахунок) (кредит) (код 1200)": ("Кошти в НБУ (кор.рахунок) (кредит) (код 1200)", 4),
        "Кошти в НБУ (УСЬОГО) (кредит) (розділ 12)": ("Кошти в НБУ (УСЬОГО) (кредит) (розділ 12)", 4),
        "Проценти за поточними вкладами ДГ (код 7040)": ("Проценти за поточними вкладами ДГ (код 7040)", 7),
        "Проценти за строковими вкладами ДГ (код 7041)": ("Проценти за строковими вкладами ДГ (код 7041)", 7),
        "Кредити субʼєктам господарювання (р.20, р.23, 260)": ("Кредити субʼєктам господарювання (р.20, р.23, 260)", 7),
        "Кредити ЗДУ (р.21)": ("Кредити ЗДУ (р.21)", 7),
        "Кредити фізичним особам (р. 22, 24, 262)": ("Кредити фізичним особам (р. 22, 24, 262)", 7),
        "Кредити небанківським установам (265)": ("Кредити небанківським установам (265)", 7),
        "Прибуток звітного року (5040)": ("Прибуток звітного року (5040)", 7),
        "Збиток звітного року (5041)": ("Збиток звітного року (5041)", 7),
        "Податок на прибуток (7900)": ("Податок на прибуток (7900)", 7),
        "Інші податки (741)": ("Інші податки (741)", 7),
        "Відрахування в резерви (770)": ("Відрахування в резерви (770)", 7),
        "Загальні доходи (6)": ("Загальні доходи (6)", 7),
        "Загальні витрати (7)": ("Загальні витрати (7)", 7)
    }

    def __init__(self, indicator_tables_creator):
        print(f"\n=== STAGE 3: Index Sheet Creation ===")
        print(f"    Initialising an IndexSheetCreator object")
        self.indicator_tables_creator = indicator_tables_creator
        self.workbook = self.indicator_tables_creator.workbook
        self.workbook_data = self.indicator_tables_creator.workbook_data
        self.parent_file = self.indicator_tables_creator.parent_file

        self.index_sheet = self.create_index_sheet() #
        self.matching_sheets = self.find_matching_sheets() #
        self.sheets_dict = self.indicator_tables_creator.sheets_dict

        self.config()
        print(f"    STAGE 3 COMPLETE")

    def config(self):
        print(f"    Performing Index Sheet configuration:")
        self.format_header() #
        self.add_hyperlinks() #
        self.fill_formulas() #
        self.data_last_row = self.index_sheet.max_row
        self.sum_columns_by_header() #
        print(f"    Index Sheet configured")

    def find_matching_sheets(self):
        matching_sheets = []
        for sheet_name in self.workbook.sheetnames:
            if sheet_name == "Index":
                continue
            if is_valid_sheet_name(sheet_name):
                matching_sheets.append(sheet_name)
        matching_sheets.sort(key=lambda x: (int(re.match(r'^\s*(\d+)', x).group(1)), x))
        print(f"    Valid sheet names separated")
        return matching_sheets

    def create_index_sheet(self): #index_sheet_creator
        """
        Create or replace the Index sheet.

        Args:
            self.workbook: openpyxl Workbook object

        Returns:
            Worksheet object for the Index sheet
        """
        # Check if Index sheet already exists
        if "Index" in self.workbook.sheetnames:
            print(f"    'Index' sheet already exists - removing it...")
            index_sheet = self.workbook["Index"]
            self.workbook.remove(index_sheet)
            print(f"     ✓ Existing Index sheet removed")

        # Create new Index sheet at the beginning
        print(f"    Creating new Index sheet...")
        index_sheet = self.workbook.create_sheet("Index", 0)
        print(f"    ✓ Index sheet created")

        return index_sheet

    def format_header(self): #index_sheet_creator
        """
        Format the header rows of the Index sheet.
        """
        print(f"        Formatting Index Sheet headers...")
        # Hide description rows
        for row_id in [1, 2, 3, 4]:
            self.index_sheet.row_dimensions[row_id].hidden = True

        # Description rows
        self.index_sheet['C2'] = "Оборотно-сальдовий баланс банку*"
        self.index_sheet['D2'] = "відповідно до Додатку 1 до постанови Правління Національного банку України"
        self.index_sheet[
            'D3'] = "від 15 березня 2018 року № 11 \"Про встановлення переліку інформації, що підлягає обов'язковому опублікуванню банками України\" (зі змінами)"
        self.index_sheet['D4'] = "за даними статистичної звітності з файлів 02X \"Дані про обороти та залишки на рахунках\"."

        # Row 5: Column headers
        self.index_sheet['A5'] = "NKB"
        self.index_sheet['B5'] = "Порядковий номер"
        self.index_sheet['C5'] = "Назва банку"
        self.index_sheet['D5'] = "МФО"

        header_fill = PatternFill(start_color="0070C0",
                                  end_color="0070C0",
                                  fill_type="solid")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Format A-D columns
        for col in ['A5', 'B5', 'C5', 'D5']:
            self.index_sheet[col].font = Font(name='Arial', size=11, bold=True, color="000000")
            self.index_sheet[col].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            self.index_sheet[col].border = thin_border

        self.index_sheet.column_dimensions['A'].width = 15
        self.index_sheet.column_dimensions['B'].width = 12
        self.index_sheet.column_dimensions['C'].width = 30
        self.index_sheet.column_dimensions['D'].width = 15

        # E-AD headers
        headers_e_ad = {
            5: "ПРОЦЕНТИ ЗА ДЕПОЗИТНИМИ СЕРТИФІКАТАМИ (ДС), тис. грн",
            6: "ПРОЦЕНТИ ЗА ОВДП (ОВДП), тис.грн",
            7: "ПРОЦЕНТИ  ЗА КРЕДИТАМИ ЗДУ",
            8: "ПРОЦЕНТИ ЗА КРЕДИТАМИ БІЗНЕСУ",
            9: "ПРОЦЕНТИ ЗА КРЕДИТАМИ НАСЕЛЕННЮ",
            10: "ПРОЦЕНТНІ ДОХОДИ ВСЬОГО (ПД), тис. грн",
            11: "ДС/ПД, %",
            12: "ДС до підсумку, %",
            13: "(ДС+ОВДП)/ПД, %",
            14: "ОВДП/ПД, %",
            15: "ОВДП до підсумку",
            16: "Обсяг ДС дебит (середні, поділено на кількість банк.днів)",
            17: "Обсяг ДС кредит (середні, поділено на кількість банк.днів)",
            18: "Кошти в НБУ (кор.рахунок) (кредит) (код 1200)",
            19:"Кошти в НБУ (УСЬОГО) (кредит) (розділ 12)",
            20:"Проценти за поточними вкладами ДГ (код 7040)",
            21:"Проценти за строковими вкладами ДГ (код 7041)",
            22:"Кредити субʼєктам господарювання (р.20, р.23, 260)",
            23:"Кредити ЗДУ (р.21)",
            24:"Кредити фізичним особам (р. 22, 24, 262)",
            25:"Кредити небанківським установам (265)",
            26:"Прибуток звітного року (5040)",
            27:"Збиток звітного року (5041)",
            28:"Податок на прибуток (7900)",
            29:"Інші податки (741)",
            30:"Відрахування в резерви (770)",
            31:"Загальні доходи (6)",
            32:"Загальні витрати (7)"
        }

        # Blue headers (columns E-J / 5-10)
        for col_idx in range(5, 11):
            col_letter = get_column_letter(col_idx)
            cell = self.index_sheet[f'{col_letter}5']
            cell.value = headers_e_ad.get(col_idx, "")
            cell.font = Font(name='Arial', size=9, bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            self.index_sheet.column_dimensions[col_letter].width = 18
            self.index_sheet.row_dimensions[5].height = 60

        # Regular headers (columns K-AF / 11-32)
        for col_idx in range(11, 40):
            col_letter = get_column_letter(col_idx)
            cell = self.index_sheet[f'{col_letter}5']
            cell.value = headers_e_ad.get(col_idx, "")
            cell.font = Font(name='Arial', size=9, bold=True, color="000000")
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            self.index_sheet.column_dimensions[col_letter].width = 18
            self.index_sheet.row_dimensions[5].height = 60

        # Freeze first 4 columns and header row 5
        self.index_sheet.freeze_panes = "E6"

        print(f"        ✓ Index Sheet headers formatted")

    def add_hyperlinks(self): #index_sheet_creator
        """
        Add hyperlinks to matching sheets starting from row 6.

        Args:
            self.workbook: openpyxl Workbook object
            self.index_sheet: The Index worksheet
            self.matching_sheets: List of sheet names that match the pattern
        """
        print(f"        Setting up hyperlinks...")

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        current_row = 6
        index_counter = 1

        for sheet_name in self.matching_sheets:
            match = re.match(r'^\s*(\d+)', sheet_name)
            number_str = match.group(1) if match else ""
            text_part = sheet_name[match.end():].strip() if match else sheet_name

            # Column A: Number with hyperlink
            cell_a = self.index_sheet.cell(row=current_row, column=1)
            cell_a.value = number_str
            cell_a.hyperlink = f"#'{sheet_name}'!A1"
            cell_a.font = Font(name='Arial', size=11, color="0000FF", underline="single")
            cell_a.border = thin_border
            cell_a.alignment = Alignment(horizontal='center', vertical='center')

            # Column B: Index (1, 2, 3, ...)
            cell_b = self.index_sheet.cell(row=current_row, column=2)
            cell_b.value = index_counter
            cell_b.font = Font(name='Arial', size=11)
            cell_b.border = thin_border
            cell_b.alignment = Alignment(horizontal='center', vertical='center')

            # Column C: Sheet name without number with hyperlink
            cell_c = self.index_sheet.cell(row=current_row, column=3)
            cell_c.value = text_part
            cell_c.hyperlink = f"#'{sheet_name}'!A1"
            cell_c.font = Font(name='Arial', size=11, color="0000FF", underline="single")
            cell_c.border = thin_border
            cell_c.alignment = Alignment(horizontal='left', vertical='center')

            # Column D: MFO (if found in "Зміст" sheet)
            mfo_value = None
            if "Зміст" in self.workbook.sheetnames and number_str:
                contents = self.workbook["Зміст"]
                for row_idx in range(1, contents.max_row + 1):
                    cell_val = contents.cell(row=row_idx, column=1).value
                    if cell_val is None:
                        continue
                    try:
                        cell_text = str(cell_val).strip()
                        if cell_text == number_str or (cell_text.isdigit() and int(cell_text) == int(number_str)):
                            mfo_value = contents.cell(row=row_idx, column=3).value
                            break
                    except Exception:
                        pass

            cell_d = self.index_sheet.cell(row=current_row, column=4)
            if mfo_value is not None:
                cell_d.value = str(mfo_value).strip()
                cell_d.hyperlink = f"#'{sheet_name}'!A1"
                cell_d.font = Font(name='Arial', size=11, color="0000FF", underline="single")
            cell_d.border = thin_border
            cell_d.alignment = Alignment(horizontal='center', vertical='center')

            current_row += 1
            index_counter += 1

        # Adjust row height for better readability
        for row in range(6, current_row):
            self.index_sheet.row_dimensions[row].height = 22

        print(f"        ✓ Hyperlinks set up for all the {len(self.matching_sheets)} valid bank sheets")

    def fill_formulas(self): #index_sheet_creator
        """
        Fill Index sheet with formulas linking to bank sheets.

        This is the core logic from Danya's_code.py
        """
        print(f"        Filling out Index Sheet formulae:")
        content_sheet = self.workbook["Index"]

        ind_cor = IndexSheetCreator.INDICATOR_CORRESPONDENCE
        print(f"            Found {len(ind_cor)} indicators to match")

        sheet_count = 0
        total_sheets = len(self.sheets_dict)
        print(f"            Processing the {total_sheets} valid bank sheets:")

        for codes in self.sheets_dict:
            sheet_count += 1
            object = self.sheets_dict[codes]
            code = codes[0]
            name_of_sheet = codes[1]

            matrix_base = object.indicator_frame_cells
            indicator_row = 5
            index_column = 1
            started_row = 6
            print(f"                [{sheet_count}/{total_sheets}] Processing: {codes}")
            print(f"                    → Looking for bank code {code} in the Index Sheet...")

            # Find the row in Index sheet that corresponds to this bank
            for row in content_sheet.iter_rows(
                    min_row=started_row,
                    max_row=content_sheet.max_row,
                    min_col=index_column,
                    max_col=index_column):
                cell = row[0]
                if cell.value != None and int(cell.value) == int(code):
                    main_bank_row = cell.row
                    print(f"                    ✓ Found at row {main_bank_row}")
                    break

            print(f"                    → Matching indicators from matrix...")

            # For each indicator in the bank's data
            for row in matrix_base.itertuples():
                indicator = row.Index

                # Find matching column header in Index sheet
                for column in content_sheet.iter_cols(min_row=indicator_row, max_row=indicator_row, min_col=5, max_col=100):
                    cell = column[0]

                    if ind_cor.get(str(cell.value)) != None and ind_cor.get(str(cell.value))[0] == str(indicator):
                        print(f"                        ✓ Match: {str(cell.value)}")
                        hnp = ind_cor[str(cell.value)][1]
                        cel = row[hnp]
                        row_link = cel.row
                        column_link = cel.column_letter
                        cell_link = content_sheet.cell(row=main_bank_row, column=cell.column)
                        cell_link.value = make_formula(name_of_sheet, column_link, row_link)
                        break

            print(f"                ✓ Formulae filled out for sheet {sheet_count}/{total_sheets}")

        print(f"            ✓ All formulae filled out!")

        # create_pivot_sheet(self.sheets_dict, self.workbook)

    def sum_columns_by_header(self): #index_sheet_creator
        """
        Write totals row under data using Excel formulas.
        Group rows are also formula-based and only filled for columns
        where row 5 header is not empty.
        """
        print(f"        Creating the totals rows")
        results = {}
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        def _to_float(value):
            if value is None:
                return 0.0
            if isinstance(value, (int, float)):
                return float(value)
            if isinstance(value, str):
                cleaned = value.replace(' ', '').replace(',', '.')
                try:
                    return float(cleaned)
                except ValueError:
                    return 0.0
            return 0.0

        def _resolve_value(value, depth=0):
            if depth > 2:
                return 0.0

            if isinstance(value, str):
                # Handle links like ='SHEET'!$C$10
                match = re.match(r"^='(.+)'!\$?([A-Z]+)\$?(\d+)$", value)
                if match:
                    src_sheet_name = match.group(1)
                    src_col = match.group(2)
                    src_row = int(match.group(3))

                    if src_sheet_name in self.workbook.sheetnames:
                        src_ws = self.workbook[src_sheet_name]
                        src_value = src_ws[f"{src_col}{src_row}"].value
                        # Prefer cached calculated value (data_only workbook)
                        if src_sheet_name in self.workbook_data.sheetnames:
                            src_ws_values = self.workbook_data[src_sheet_name]
                            cached_value = src_ws_values[f"{src_col}{src_row}"].value
                            if cached_value is not None:
                                return _to_float(cached_value)

                        return _resolve_value(src_value, depth + 1)
                    return 0.0

            return _to_float(value)

        print(f"            Calculating the totals for each column")

        # Write totals row under all bank rows
        total_row = self.data_last_row + 1
        self.index_sheet.cell(row=total_row, column=1).value = "Всього"
        self.index_sheet.cell(row=total_row, column=1).font = Font(name='Arial', size=10, bold=True)
        self.index_sheet.cell(row=total_row, column=1).alignment = Alignment(horizontal='center', vertical='center')

        # Keep list of columns where row-5 header is not empty
        non_empty_header_cols = []

        # Map row-5 headers to column indexes
        header_to_col = {}
        for col_idx in range(5, self.index_sheet.max_column + 1):
            header_value = self.index_sheet.cell(row=5, column=col_idx).value
            if header_value is not None and str(header_value).strip() != "":
                header_to_col[str(header_value)] = col_idx

        ds_col_idx = header_to_col.get("ПРОЦЕНТИ ЗА ДЕПОЗИТНИМИ СЕРТИФІКАТАМИ (ДС), тис. грн")
        ovdp_col_idx = header_to_col.get("ПРОЦЕНТИ ЗА ОВДП (ОВДП), тис.грн")
        pd_col_idx = header_to_col.get("ПРОЦЕНТНІ ДОХОДИ ВСЬОГО (ПД), тис. грн")
        ds_pd_col_idx = header_to_col.get("ДС/ПД, %")
        ds_share_col_idx = header_to_col.get("ДС до підсумку, %")
        ds_ovdp_pd_col_idx = header_to_col.get("(ДС+ОВДП)/ПД, %")
        ovdp_pd_col_idx = header_to_col.get("ОВДП/ПД, %")
        ovdp_share_col_idx = header_to_col.get("ОВДП до підсумку")
        finrez_col_idx = header_to_col.get("ФінРез (ОВДП-ДС)")

        # Fill DS/PD formula for each bank row (same-row calculation)
        if ds_col_idx is not None and pd_col_idx is not None and ds_pd_col_idx is not None:
            ds_col_letter = get_column_letter(ds_col_idx)
            pd_col_letter = get_column_letter(pd_col_idx)
            ds_pd_col_letter = get_column_letter(ds_pd_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                ds_pd_cell = self.index_sheet[f"{ds_pd_col_letter}{row_idx}"]
                ds_pd_cell.value = f"=IFERROR({ds_col_letter}{row_idx}/{pd_col_letter}{row_idx}*100,0)"
                ds_pd_cell.number_format = '0.0'

        # Fill DS share formula for each bank row (row DS / total DS column)
        if ds_col_idx is not None and ds_share_col_idx is not None:
            ds_col_letter = get_column_letter(ds_col_idx)
            ds_share_col_letter = get_column_letter(ds_share_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                ds_share_cell = self.index_sheet[f"{ds_share_col_letter}{row_idx}"]
                ds_share_cell.value = f"=IFERROR({ds_col_letter}{row_idx}/SUM({ds_col_letter}6:{ds_col_letter}{self.data_last_row})*100,0)"
                ds_share_cell.number_format = '0.0'

        # Fill (DS+OVDP)/PD formula for each bank row
        if ds_col_idx is not None and ovdp_col_idx is not None and pd_col_idx is not None and ds_ovdp_pd_col_idx is not None:
            ds_col_letter = get_column_letter(ds_col_idx)
            ovdp_col_letter = get_column_letter(ovdp_col_idx)
            pd_col_letter = get_column_letter(pd_col_idx)
            ds_ovdp_pd_col_letter = get_column_letter(ds_ovdp_pd_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                ds_ovdp_pd_cell = self.index_sheet[f"{ds_ovdp_pd_col_letter}{row_idx}"]
                ds_ovdp_pd_cell.value = f"=IFERROR(({ds_col_letter}{row_idx}+{ovdp_col_letter}{row_idx})/{pd_col_letter}{row_idx}*100,0)"
                ds_ovdp_pd_cell.number_format = '0.0'

        # Fill OVDP/PD formula for each bank row
        if ovdp_col_idx is not None and pd_col_idx is not None and ovdp_pd_col_idx is not None:
            ovdp_col_letter = get_column_letter(ovdp_col_idx)
            pd_col_letter = get_column_letter(pd_col_idx)
            ovdp_pd_col_letter = get_column_letter(ovdp_pd_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                ovdp_pd_cell = self.index_sheet[f"{ovdp_pd_col_letter}{row_idx}"]
                ovdp_pd_cell.value = f"=IFERROR({ovdp_col_letter}{row_idx}/{pd_col_letter}{row_idx}*100,0)"
                ovdp_pd_cell.number_format = '0.0'

        # Fill OVDP share formula for each bank row (row OVDP / total OVDP column)
        if ovdp_col_idx is not None and ovdp_share_col_idx is not None:
            ovdp_col_letter = get_column_letter(ovdp_col_idx)
            ovdp_share_col_letter = get_column_letter(ovdp_share_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                ovdp_share_cell = self.index_sheet[f"{ovdp_share_col_letter}{row_idx}"]
                ovdp_share_cell.value = f"=IFERROR({ovdp_col_letter}{row_idx}/SUM({ovdp_col_letter}6:{ovdp_col_letter}{self.data_last_row})*100,0)"
                ovdp_share_cell.number_format = '0.0'

        # Fill FinRez formula for each bank row: DS + OVDP
        if ds_col_idx is not None and ovdp_col_idx is not None and finrez_col_idx is not None:
            ds_col_letter = get_column_letter(ds_col_idx)
            ovdp_col_letter = get_column_letter(ovdp_col_idx)
            finrez_col_letter = get_column_letter(finrez_col_idx)

            for row_idx in range(6, self.data_last_row + 1):
                finrez_cell = self.index_sheet[f"{finrez_col_letter}{row_idx}"]
                finrez_cell.value = f"=IFERROR({ds_col_letter}{row_idx}+{ovdp_col_letter}{row_idx},0)"
                finrez_cell.number_format = '#,##0.00'

        # Iterate through data columns (E and onward)
        for col_idx in range(5, self.index_sheet.max_column + 1):
            col_letter = get_column_letter(col_idx)
            header_cell = self.index_sheet[f'{col_letter}5']
            header_value = header_cell.value

            if header_value is None or str(header_value).strip() == "":
                continue

            non_empty_header_cols.append(col_idx)
            sum_cell = self.index_sheet[f"{col_letter}{total_row}"]

            # Special handling for percentage columns
            if ds_pd_col_idx is not None and col_idx == ds_pd_col_idx:
                if ds_col_idx is not None and pd_col_idx is not None:
                    ds_col_letter = get_column_letter(ds_col_idx)
                    pd_col_letter = get_column_letter(pd_col_idx)
                    sum_cell.value = f"=IFERROR({ds_col_letter}{total_row}/{pd_col_letter}{total_row}*100,0)"
                else:
                    sum_cell.value = "=0"
            elif ds_share_col_idx is not None and col_idx == ds_share_col_idx:
                if ds_col_idx is not None:
                    ds_col_letter = get_column_letter(ds_col_idx)
                    sum_cell.value = f"=IFERROR({ds_col_letter}{total_row}/SUM({ds_col_letter}6:{ds_col_letter}{self.data_last_row})*100,0)"
                else:
                    sum_cell.value = "=0"
            elif ds_ovdp_pd_col_idx is not None and col_idx == ds_ovdp_pd_col_idx:
                if ds_col_idx is not None and ovdp_col_idx is not None and pd_col_idx is not None:
                    ds_col_letter = get_column_letter(ds_col_idx)
                    ovdp_col_letter = get_column_letter(ovdp_col_idx)
                    pd_col_letter = get_column_letter(pd_col_idx)
                    sum_cell.value = f"=IFERROR(({ds_col_letter}{total_row}+{ovdp_col_letter}{total_row})/{pd_col_letter}{total_row}*100,0)"
                else:
                    sum_cell.value = "=0"
            elif ovdp_pd_col_idx is not None and col_idx == ovdp_pd_col_idx:
                if ovdp_col_idx is not None and pd_col_idx is not None:
                    ovdp_col_letter = get_column_letter(ovdp_col_idx)
                    pd_col_letter = get_column_letter(pd_col_idx)
                    sum_cell.value = f"=IFERROR({ovdp_col_letter}{total_row}/{pd_col_letter}{total_row}*100,0)"
                else:
                    sum_cell.value = "=0"
            elif ovdp_share_col_idx is not None and col_idx == ovdp_share_col_idx:
                if ovdp_col_idx is not None:
                    ovdp_col_letter = get_column_letter(ovdp_col_idx)
                    sum_cell.value = f"=IFERROR({ovdp_col_letter}{total_row}/SUM({ovdp_col_letter}6:{ovdp_col_letter}{self.data_last_row})*100,0)"
                else:
                    sum_cell.value = "=0"
            else:
                sum_cell.value = f"=IFERROR(SUM({col_letter}6:{col_letter}{self.data_last_row}),0)"

            sum_cell.font = Font(name='Arial', size=10, bold=True)
            sum_cell.alignment = Alignment(horizontal='center', vertical='center')
            sum_cell.number_format = '#,##0.00'

            results[str(header_value)] = {
                'column': col_letter,
                'formula': sum_cell.value
            }
            print(f"                ✓ Column {col_letter}: {header_value}")
            print(f"                Formula: {sum_cell.value}")

        if not non_empty_header_cols:
            print(f"            ⚠ No non-empty headers found from row 5")
            return results

        last_formula_col = max(non_empty_header_cols)

        # Apply row styling (borders/fill) across entire Index width
        for col_idx in range(1, last_formula_col + 1):
            cell = self.index_sheet.cell(row=total_row, column=col_idx)
            cell.border = thin_border
            cell.fill = total_fill

        self.index_sheet.row_dimensions[total_row].height = 22

        # Category totals by fixed bank groups
        bank_row_by_name = {}
        for row_idx in range(6, self.data_last_row + 1):
            bank_name = _normalize_name(self.index_sheet.cell(row=row_idx, column=3).value)
            if bank_name:
                bank_row_by_name[bank_name] = row_idx

        state_rows = [row for name, row in bank_row_by_name.items() if name in STATE_BANKS]
        foreign_rows = [row for name, row in bank_row_by_name.items() if name in FOREIGN_BANKS]
        all_rows = list(bank_row_by_name.values())

        state_row = total_row + 1
        foreign_row = total_row + 2
        private_row = total_row + 3

        category_rows = [
            (state_row, "Державні банки", state_rows),
            (foreign_row, "Банки з іноземним капіталом", foreign_rows),
            (private_row, "Банки з українським приватним капіталом", None),
        ]

        def _category_formula(rows_list, col_index):
            if not rows_list:
                return "=0"
            col_letter = get_column_letter(col_index)
            refs = ",".join([f"{col_letter}{r}" for r in sorted(rows_list)])
            return f"=IFERROR(SUM({refs}),0)"

        for out_row, category_name, source_rows in category_rows:
            self.index_sheet.cell(row=out_row, column=1).value = category_name
            self.index_sheet.cell(row=out_row, column=1).font = Font(name='Arial', size=10)
            self.index_sheet.cell(row=out_row, column=1).alignment = Alignment(horizontal='left', vertical='center')

            # Sum across only columns that have non-empty header in row 5
            for col_idx in non_empty_header_cols:
                out_cell = self.index_sheet.cell(row=out_row, column=col_idx)
                if ds_pd_col_idx is not None and col_idx == ds_pd_col_idx:
                    if ds_col_idx is not None and pd_col_idx is not None:
                        ds_col_letter = get_column_letter(ds_col_idx)
                        pd_col_letter = get_column_letter(pd_col_idx)
                        out_cell.value = f"=IFERROR({ds_col_letter}{out_row}/{pd_col_letter}{out_row}*100,0)"
                    else:
                        out_cell.value = "=0"
                elif ds_share_col_idx is not None and col_idx == ds_share_col_idx:
                    if ds_col_idx is not None:
                        ds_col_letter = get_column_letter(ds_col_idx)
                        out_cell.value = f"=IFERROR({ds_col_letter}{out_row}/{ds_col_letter}{total_row}*100,0)"
                    else:
                        out_cell.value = "=0"
                elif ovdp_share_col_idx is not None and col_idx == ovdp_share_col_idx:
                    if ovdp_col_idx is not None:
                        ovdp_col_letter = get_column_letter(ovdp_col_idx)
                        out_cell.value = f"=IFERROR({ovdp_col_letter}{out_row}/{ovdp_col_letter}{total_row}*100,0)"
                    else:
                        out_cell.value = "=0"
                elif source_rows is None:
                    col_letter = get_column_letter(col_idx)
                    out_cell.value = f"=IFERROR({col_letter}{total_row}-{col_letter}{total_row + 1}-{col_letter}{total_row + 2},0)"
                else:
                    out_cell.value = _category_formula(source_rows, col_idx)
                out_cell.number_format = '#,##0.00'
                out_cell.alignment = Alignment(horizontal='center', vertical='center')

            for col_idx in range(1, last_formula_col + 1):
                cell = self.index_sheet.cell(row=out_row, column=col_idx)
                cell.border = thin_border
                cell.fill = total_fill

            self.index_sheet.row_dimensions[out_row].height = 22

        print(f"        Totals rows filled out")
        print(f"        Creating ancillary tables")

        # Additional two formula tables on Index (for middle and right charts)

        # Place both small tables under the main table, starting from column B
        middle_col_name = 2  # B
        middle_col_value = 3  # C
        right_col_name = 8  # H
        right_col_value = 9  # I
        table_header_row = private_row + 2

        categories = [
            ("Державні банки", state_row),
            ("Банки з іноземним капіталом", foreign_row),
            ("Банки з українським приватним капіталом", private_row),
        ]

        # Middle table: "Частка процентів по депсертифікатах у процентних доходах банків, %"
        self.index_sheet.cell(row=table_header_row, column=middle_col_name).value = "Категорія"
        self.index_sheet.cell(row=table_header_row, column=middle_col_value).value = "ДС/ПД, %"

        # Right table: "Проценти за депозитними сертифікатами ... у розрізі банків, %"
        self.index_sheet.cell(row=table_header_row, column=right_col_name).value = "Категорія"
        self.index_sheet.cell(row=table_header_row, column=right_col_value).value = "ДС до підсумку, %"

        for idx, (category_label, category_row) in enumerate(categories, start=1):
            out_row = table_header_row + idx

            # Middle table values
            self.index_sheet.cell(row=out_row, column=middle_col_name).value = f"=A{category_row}"
            if ds_pd_col_idx is not None:
                ds_pd_letter = get_column_letter(ds_pd_col_idx)
                self.index_sheet.cell(row=out_row, column=middle_col_value).value = f"=IFERROR({ds_pd_letter}{category_row},0)"
            else:
                self.index_sheet.cell(row=out_row, column=middle_col_value).value = "=0"

            # Right table values
            self.index_sheet.cell(row=out_row, column=right_col_name).value = f"=A{category_row}"
            if ds_share_col_idx is not None:
                ds_share_letter = get_column_letter(ds_share_col_idx)
                self.index_sheet.cell(row=out_row, column=right_col_value).value = f"=IFERROR({ds_share_letter}{category_row},0)"
            else:
                self.index_sheet.cell(row=out_row, column=right_col_value).value = "=0"

        # Style both tables
        table_cols = [middle_col_name, middle_col_value, right_col_name, right_col_value]
        for r in range(table_header_row, table_header_row + 4):
            for c in table_cols:
                cell = self.index_sheet.cell(row=r, column=c)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                if r == table_header_row:
                    cell.font = Font(name='Arial', size=10, bold=True)
                    cell.fill = total_fill
                else:
                    if c in [middle_col_value, right_col_value]:
                        cell.number_format = '0.0'

        self.index_sheet.column_dimensions[get_column_letter(middle_col_name)].width = 38
        self.index_sheet.column_dimensions[get_column_letter(middle_col_value)].width = 14
        self.index_sheet.column_dimensions[get_column_letter(right_col_name)].width = 38
        self.index_sheet.column_dimensions[get_column_letter(right_col_value)].width = 16

        print(f"        Both ancillary tables created")
        print(f"        Creating Index Sheet charts")

        # Charts under the two small tables (pie on the left, bar on the right)
        middle_cats = Reference(self.index_sheet, min_col=middle_col_name, min_row=table_header_row + 1, max_row=table_header_row + 3)
        middle_vals = Reference(self.index_sheet, min_col=middle_col_value, min_row=table_header_row + 1, max_row=table_header_row + 3)

        bar = BarChart()
        bar.type = "col"
        bar.title = "Частка процентів по депсертифікатах у\nпроцентних доходах банків, %"
        bar.add_data(middle_vals, titles_from_data=False)
        bar.set_categories(middle_cats)
        bar.height = 13
        bar.width = 13
        bar.legend = None
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        bar.x_axis.title = None
        bar.y_axis.title = None
        bar.dataLabels = DataLabelList()
        bar.dataLabels.showSerName = False
        bar.dataLabels.showCatName = False
        bar.dataLabels.showVal = True
        bar.dataLabels.txPr = _datalabel_text_style(size_pt=13, bold=True, color="FF0000")

        right_cats = Reference(self.index_sheet, min_col=right_col_name, min_row=table_header_row + 1, max_row=table_header_row + 3)
        right_vals = Reference(self.index_sheet, min_col=right_col_value, min_row=table_header_row + 1, max_row=table_header_row + 3)

        pie_right = PieChart()
        pie_right.title = "Проценти за депозитними сертифікатами\nу розрізі банків, %"
        pie_right.title.txPr = _datalabel_text_style(size_pt=13, bold=True, color="FF0000")
        pie_right.add_data(right_vals, titles_from_data=False)
        pie_right.set_categories(right_cats)
        pie_right.height = 14
        pie_right.width = 14
        pie_right.legend = None
        pie_right.dataLabels = DataLabelList()
        pie_right.dataLabels.showSerName = False
        pie_right.dataLabels.showCatName = True
        pie_right.dataLabels.showVal = False
        pie_right.dataLabels.showPercent = True
        pie_right.dataLabels.showLeaderLines = True
        pie_right.dataLabels.txPr = _datalabel_text_style(size_pt=11, bold=True, color="000000")

        chart_top_row = table_header_row + 5
        self.index_sheet.add_chart(bar, f"B{chart_top_row}")
        self.index_sheet.add_chart(pie_right, f"H{chart_top_row}")

        print(f"        Index Sheet charts created")

        # print(f"✓ Category rows written: {len(category_rows)}")
        # print("  ✓ Added 2 extra source tables on Index")
        #
        # print(f"  ✓ Totals written to row {total_row}")

        return results