from .utils import *
from ..account_rows import BalanceSheet
import time


class IndicatorTablesCreator:
    def __init__(self, workbook, workbook_data, parent_file, progress_var):
        print(f"\n=== STAGE 2: Indicator Tables Creation ===")
        print(f"    Initialising an IndicatorTablesCreator object")
        self.workbook = workbook

        # NOTE: You MUST pass workbook_data into main.py again so we can extract the values!
        self.workbook_data = workbook_data

        self.parent_file = parent_file
        self.progress_var = progress_var
        self.delay = 0.01
        self.percentage = 55
        self.sheets_dict = {}
        self.create_sheets_dict()
        print(f"    STAGE 2 COMPLETE")

    def update_progress(self, value):
        current = self.progress_var.get()
        new = current + value / 100.0
        self.progress_var.set(new)
        time.sleep(self.delay)

    def create_sheets_dict(self):
        print(f"    Creating BalanceSheet objects:")

        valid_sheets = [s.title for s in self.workbook.worksheets if code_bank(s.title) is not None]
        total = len(valid_sheets)

        if total == 0:
            return

        increment_percent = 55 / total

        for i, sheet_name in enumerate(valid_sheets):
            print(f"        [{i + 1}/{total}] Processing sheet: {sheet_name}")

            target_sheet_write = self.workbook[sheet_name]

            # EXTRACT THE GRID USING CALAMINE
            # .to_python() instantly returns the exact 2D list of lists we need
            # It also automatically strips out formulas and gives you the raw values
            try:
                calamine_sheet = self.workbook_data.get_sheet_by_name(sheet_name)
                sheet_values_grid = calamine_sheet.to_python()
            except KeyError:
                print(f"        [!] Sheet {sheet_name} not found in data workbook.")
                continue

            # PASS THE GRID TO BALANCESHEET (Nothing changes here!)
            bs = BalanceSheet(
                parent_file=self.parent_file,
                sheet=target_sheet_write,
                sheet_grid=sheet_values_grid
            )

            bs.insert_frame(
                parent_file=self.parent_file,
                start_row=bs.start_row,
                start_col=bs.start_col
            )

            title = code_bank(sheet_name)
            self.sheets_dict[title] = bs
            self.update_progress(increment_percent)

        print(f"    ✓ Created dictionary and wrote {len(self.sheets_dict)} indicator tables")