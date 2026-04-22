from .utils import *
from ..account_rows import BalanceSheet
import time

class IndicatorTablesCreator:
    def __init__(self, workbook, workbook_data, parent_file, progress_var):
        print(f"\n=== STAGE 2: Indicator Tables Creation ===")
        print(f"    Initialising an IndicatorTablesCreator object")
        self.workbook = workbook
        self.workbook_data = workbook_data
        self.parent_file = parent_file
        self.progress_var = progress_var
        self.delay = 0.01
        self.percentage = 55
        self.sheets_dict = self.create_sheets_dict()
        print(f"    STAGE 2 COMPLETE")

    def update_progress(self, value):
        current = self.progress_var.get()
        new = current + value / 100.0
        self.progress_var.set(new)
        time.sleep(self.delay)

    def create_sheets_dict(self): #indicator_tables_creator
        """Create dictionary of valid bank sheets with BalanceSheet objects."""
        print(f"    Creating BalanceSheet objects:")
        sheets_dict = {}

        sheet_num = 0
        total_sheets = len(self.workbook.worksheets)

        increment_percent = 55 / total_sheets if total_sheets else 0

        for sheet in self.workbook.worksheets:
            sheet_num += 1
            title = code_bank(sheet.title)
            if title != None:
                print(f"        [{sheet_num}/{total_sheets}] Processing sheet: {sheet.title}")
                sheet_value = BalanceSheet(sheet=sheet, parent_file=self.parent_file)
                sheets_dict[title] = sheet_value
                print(f"            ✓ Added to dictionary")
            else:
                print(f"        [{sheet_num}/{total_sheets}] Skipping: {sheet.title} (no bank code)")
            self.update_progress(increment_percent)

        print(f"    ✓ Created dictionary with {len(sheets_dict)} BalanceSheet objects")
        return sheets_dict