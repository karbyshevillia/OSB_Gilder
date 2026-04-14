from .utils import *

class PivotSheetCreator:
    def __init__(self, index_sheet_creator):
        self.index_sheet_creator = index_sheet_creator
        self.workbook = self.index_sheet_creator.workbook
        self.sheets_dict = self.index_sheet_creator.sheets_dict

        self.pivot_sheet = self.create_pivot_sheet()

    def create_pivot_sheet(self): #pivot_sheet_creator
        pivot_sheet = self.workbook.create_sheet("Pivot_Data")
        df = pd.concat([bs.pivot_db for bs in self.sheets_dict.values()])
        for row in dataframe_to_rows(df, index=False, header=True):
            pivot_sheet.append(row)
        tab = Table(displayName="Pivot_Data", ref=f"A1:{get_column_letter(pivot_sheet.max_column)}{pivot_sheet.max_row}")

        # 3. Add a visual style (optional but recommended)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style

        # 4. Add the table to the worksheet
        pivot_sheet.add_table(tab)
        return pivot_sheet