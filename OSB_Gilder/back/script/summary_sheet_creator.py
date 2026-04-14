from .utils import *
from .utils import _datalabel_text_style
from .index_sheet_creator import IndexSheetCreator

class SummarySheetCreator:
    def __init__(self, index_sheet_creator):
        self.index_sheet_creator = index_sheet_creator
        self.workbook = self.index_sheet_creator.workbook
        self.workbook_data = self.index_sheet_creator.workbook_data
        self.indicator_correspondence = IndexSheetCreator.INDICATOR_CORRESPONDENCE
        self.index_sheet = self.index_sheet_creator.index_sheet
        self.data_last_row = self.index_sheet_creator.data_last_row

        self.indicator_summary_sheet = self.create_indicator_summary_sheet()

    #summary_sheet_creator
    def create_indicator_summary_sheet(self):
        """
        Create a formatted summary table similar to the provided example:
        - Column A: category names
        - Column B: sums in billions (SUM from Index / 1_000_000)
        - Column C: share in total, %

        "Інші процентні доходи" = "ВСЬОГО" - (ДС + ОВДП + бізнес + населення)
        """
        print("\n=== STEP 6: Creating indicator summary table ===")

        summary_sheet_name = "Indicator_Summary"

        if summary_sheet_name in self.workbook.sheetnames:
            self.workbook.remove(self.workbook[summary_sheet_name])

        index_position = self.workbook.sheetnames.index("Index") if "Index" in self.workbook.sheetnames else 0
        summary_sheet = self.workbook.create_sheet(summary_sheet_name, index=index_position + 1)

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        title_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        body_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

        # Row 1: title
        summary_sheet.merge_cells("A1:C1")
        summary_sheet["A1"] = "ПРОЦЕНТНІ ДОХОДИ банків"
        summary_sheet["A1"].font = Font(name='Arial', size=11, bold=True)
        summary_sheet["A1"].alignment = Alignment(horizontal='center', vertical='center')
        for col in ["A", "B", "C"]:
            summary_sheet[f"{col}1"].fill = title_fill
            summary_sheet[f"{col}1"].border = thin_border

        # Row 2: headers
        summary_sheet["B2"] = "млрд"
        summary_sheet["C2"] = "%"
        for cell_ref in ["A2", "B2", "C2"]:
            summary_sheet[cell_ref].fill = title_fill
            summary_sheet[cell_ref].font = Font(name='Arial', size=10, bold=True)
            summary_sheet[cell_ref].alignment = Alignment(horizontal='center', vertical='center')
            summary_sheet[cell_ref].border = thin_border

        # Map header text from Index row 5 to its column index
        header_to_col_idx = {}
        for col_idx in range(1, self.index_sheet.max_column + 1):
            header_val = self.index_sheet.cell(row=5, column=col_idx).value
            if header_val is not None:
                header_to_col_idx[str(header_val)] = col_idx

        last_data_row = max(6, self.data_last_row)

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

        def _resolve_index_cell_value(cell_value):
            """Resolve direct numbers and links like ='Sheet'!$C$10 to float."""
            if isinstance(cell_value, str):
                match = re.match(r"^='(.+)'!\$?([A-Z]+)\$?(\d+)$", cell_value)
                if match:
                    src_sheet_name = match.group(1)
                    src_col = match.group(2)
                    src_row = int(match.group(3))
                    print(f"ids: {src_col}, {src_row}")

                    print(src_sheet_name, src_sheet_name in self.workbook.sheetnames)
                    if src_sheet_name in self.workbook.sheetnames:
                        src_sheet = self.workbook[src_sheet_name]
                        src_raw = src_sheet[f"{src_col}{src_row}"].value
                        print(f"raw: {src_raw}")

                        # Prefer cached calculated value (data_only workbook)
                        print(src_sheet_name in self.workbook_data.sheetnames)
                        if src_sheet_name in self.workbook_data.sheetnames:
                            src_values_sheet = self.workbook_data[src_sheet_name]
                            print(src_values_sheet)
                            src_cached = src_values_sheet[f"{src_col}{src_row}"].value
                            print(f"cached: {src_cached}")
                            if src_cached is not None:
                                # print(f"to_float: {_to_float(src_cached)}")
                                return _to_float(src_cached)
                            print(f"to_float: {_to_float(src_cached)}")
                        return _to_float(src_raw)
                    return 0.0

            return _to_float(cell_value)

        def indicator_sum_value(header_name):
            col_idx = header_to_col_idx.get(header_name)
            print(header_name, col_idx)
            if col_idx is None:
                return 0.0
            col = get_column_letter(col_idx)
            total = f"=SUM({self.index_sheet.title}!{col}6:{col}{last_data_row})/1000000"

            return total

        ds_header = "ПРОЦЕНТИ ЗА ДЕПОЗИТНИМИ СЕРТИФІКАТАМИ (ДС), тис. грн"
        ovdp_header = "ПРОЦЕНТИ ЗА ОВДП (ОВДП), тис.грн"
        business_header = "ПРОЦЕНТИ ЗА КРЕДИТАМИ БІЗНЕСУ"
        population_header = "ПРОЦЕНТИ ЗА КРЕДИТАМИ НАСЕЛЕННЮ"
        total_header = "ПРОЦЕНТНІ ДОХОДИ ВСЬОГО (ПД), тис. грн"

        ds_value = indicator_sum_value(ds_header)
        ovdp_value = indicator_sum_value(ovdp_header)
        business_value = indicator_sum_value(business_header)
        population_value = indicator_sum_value(population_header)
        total_value = indicator_sum_value(total_header)
        other_value = (f"={summary_sheet.cell(row=8, column=2).coordinate} - ("
                       f"{summary_sheet.cell(row=3, column=2).coordinate} +"
                       f"{summary_sheet.cell(row=4, column=2).coordinate} +"
                       f"{summary_sheet.cell(row=5, column=2).coordinate} +"
                       f"{summary_sheet.cell(row=6, column=2).coordinate})")

        rows = [
            (3, "Депозитні сертифікати", ds_value),
            (4, "ОВДП", ovdp_value),
            (5, "Кредити бізнесу", business_value),
            (6, "Кредити населенню", population_value),
            (7, "Інші процентні доходи", other_value),
            (8, "ВСЬОГО", total_value),
        ]

        for row_idx, row_name, numeric_value in rows:
            summary_sheet.cell(row=row_idx, column=1).value = row_name
            summary_sheet.cell(row=row_idx, column=2).value = numeric_value

            if row_idx == 8:
                summary_sheet.cell(row=row_idx, column=3).value = 100.0
            else:
                num = summary_sheet.cell(row=row_idx, column=2).coordinate
                tot = summary_sheet.cell(row=8, column=2).coordinate
                summary_sheet.cell(row=row_idx, column=3).value = f"={num}/{tot} * 100"

            for col_idx in range(1, 4):
                cell = summary_sheet.cell(row=row_idx, column=col_idx)
                cell.fill = body_fill
                cell.border = thin_border
                if col_idx == 1:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = '0.0'

        # Optional last row from sample layout
        summary_sheet.cell(row=9, column=1).value = "у % до доходів банків"
        summary_sheet.cell(row=9, column=2).value = ""
        summary_sheet.cell(row=9, column=3).value = ""
        for col_idx in range(1, 4):
            cell = summary_sheet.cell(row=9, column=col_idx)
            cell.fill = body_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='left' if col_idx == 1 else 'center', vertical='center')

        # Emphasize total row
        for col_idx in range(1, 4):
            summary_sheet.cell(row=8, column=col_idx).font = Font(name='Arial', size=10, bold=True)

        summary_sheet.column_dimensions['A'].width = 38
        summary_sheet.column_dimensions['B'].width = 12
        summary_sheet.column_dimensions['C'].width = 10
        summary_sheet.column_dimensions['E'].width = 4
        summary_sheet.column_dimensions['F'].width = 20
        summary_sheet.column_dimensions['G'].width = 20
        summary_sheet.column_dimensions['H'].width = 20
        summary_sheet.row_dimensions[1].height = 22

        # Pie chart based on rows 3-7 from Indicator_Summary table
        labels = Reference(summary_sheet, min_col=1, min_row=3, max_row=7)
        data = Reference(summary_sheet, min_col=2, min_row=2, max_row=7)

        pie = PieChart()
        pie.title = "Структура процентних доходів банків\nза 12 міс. 2025 року, %"
        pie.title.txPr = _datalabel_text_style(size_pt=13, bold=True, color="FF0000")
        pie.title.overlay = False
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.height = 10
        pie.width = 13
        pie.legend = None

        pie.dataLabels = DataLabelList()
        pie.dataLabels.showSerName = False
        pie.dataLabels.showCatName = True
        pie.dataLabels.showVal = False
        pie.dataLabels.showPercent = True
        pie.dataLabels.showLeaderLines = True
        pie.dataLabels.txPr = _datalabel_text_style(size_pt=10, bold=True, color="000000")

        summary_sheet.add_chart(pie, "F1")

        print(f"  ✓ Created '{summary_sheet_name}' in sample layout")
        return summary_sheet