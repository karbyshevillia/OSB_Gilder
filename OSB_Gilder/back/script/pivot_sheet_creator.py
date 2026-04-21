from .utils import *
from ..account_rows import BalanceSheet

class PivotSheetCreator:
    def __init__(self, index_sheet_creator):
        print(f"\n=== STAGE 5: Pivot Sheet Creation ===")
        print(f"    Initialising a PivotSheetCreator object")
        self.index_sheet_creator = index_sheet_creator
        self.workbook = self.index_sheet_creator.workbook
        self.sheets_dict = self.index_sheet_creator.sheets_dict

        self.pivot_sheet = self.create_pivot_sheet()
        print(f"    STAGE 5 COMPLETE")

    def create_pivot_sheet(self): #pivot_sheet_creator
        index_position = self.workbook.sheetnames.index("Indicator_Summary") if "Indicator_Summary" in self.workbook.sheetnames else 0
        pivot_sheet = self.workbook.create_sheet("Pivot_Data", index=index_position + 1)
        print(f"    Generating database")
        df = pd.concat([bs.pivot_db for bs in self.sheets_dict.values()])
        for row in dataframe_to_rows(df, index=False, header=True):
            pivot_sheet.append(row)
        print(f"    Creating the table in the Pivot_Data Sheet")
        tab = Table(displayName="Pivot_Data", ref=f"A1:{get_column_letter(pivot_sheet.max_column)}{pivot_sheet.max_row}")

        # 3. Add a visual style (optional but recommended)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style

        # 4. Add the table to the worksheet
        pivot_sheet.add_table(tab)
        return pivot_sheet