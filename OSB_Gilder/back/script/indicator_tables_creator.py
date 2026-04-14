from .utils import *

class IndicatorTablesCreator:
    def __init__(self, workbook, workbook_data, parent_file):
        self.workbook = workbook
        self.workbook_data = workbook_data
        self.parent_file = parent_file
        self.sheets_dict = self.create_sheets_dict()

    def create_sheets_dict(self): #indicator_tables_creator
        """Create dictionary of valid bank sheets with BalanceSheet objects."""
        print("\n=== STEP 2: Creating BalanceSheet objects ===")
        sheets_dict = {}

        sheet_num = 0
        total_sheets = len(self.workbook.worksheets)

        for sheet in self.workbook.worksheets:
            sheet_num += 1
            title = code_bank(sheet.title)
            if title != None:
                print(f"  [{sheet_num}/{total_sheets}] Processing sheet: {sheet.title}")
                sheet_value = BalanceSheet(sheet=sheet, parent_file=self.parent_file)
                sheets_dict[title] = sheet_value
                print(f"      ✓ Added to dictionary")
            else:
                print(f"  [{sheet_num}/{total_sheets}] Skipping: {sheet.title} (no bank code)")

        print(f"\n✓ Created dictionary with {len(sheets_dict)} bank sheets\n")
        return sheets_dict