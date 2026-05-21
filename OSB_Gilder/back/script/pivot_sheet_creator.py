from .utils import *
from ..account_rows import BalanceSheet
import time

class PivotSheetCreator:
    def __init__(self, index_sheet_creator):
        print(f"\n=== STAGE 5: Pivot Sheet Creation ===")
        print(f"    Initialising a PivotSheetCreator object")
        self.index_sheet_creator = index_sheet_creator
        self.workbook = self.index_sheet_creator.workbook
        self.sheets_dict = self.index_sheet_creator.sheets_dict
        self.progress_var = self.index_sheet_creator.progress_var
        self.delay = 0.01
        self.percentage = 5

        self.pivot_sheet = self.create_pivot_sheet()
        print(f"    STAGE 5 COMPLETE")

    def update_progress(self, value):
        current = self.progress_var.get()
        new = current + value / 100.0
        self.progress_var.set(new)
        time.sleep(self.delay)

    def create_pivot_sheet(self): #pivot_sheet_creator
        increment_percent = self.percentage / 2

        index_position = self.workbook.sheetnames.index("Indicator_Summary") if "Indicator_Summary" in self.workbook.sheetnames else 0
        pivot_sheet = self.workbook.create_sheet("Pivot_Data", index=index_position + 1)
        print(f"    Generating database")
        df = pd.concat([bs.create_db for bs in self.sheets_dict.values()])
        for row in dataframe_to_rows(df, index=False, header=True):
            pivot_sheet.append(row)
        self.update_progress(increment_percent)
        print(f"    Creating the table in the Pivot_Data Sheet")
        tab = Table(displayName="Pivot_Data", ref=f"A1:{get_column_letter(pivot_sheet.max_column)}{pivot_sheet.max_row}")

        # 3. Add a visual style (optional but recommended)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style

        # 4. Add the table to the worksheet
        pivot_sheet.add_table(tab)
        self.update_progress(increment_percent)

        return pivot_sheet