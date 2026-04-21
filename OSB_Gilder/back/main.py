import OSB_Gilder.back.script.utils as utils
from OSB_Gilder.back.script.indicator_tables_creator import IndicatorTablesCreator
from OSB_Gilder.back.script.index_sheet_creator import IndexSheetCreator
from OSB_Gilder.back.script.pivot_sheet_creator import PivotSheetCreator
from OSB_Gilder.back.script.summary_sheet_creator import SummarySheetCreator

class OSBGilder:
    def __init__(self, parent_file):
        print(f"Initialising OSBGilder object for {parent_file}")
        self.parent_file = parent_file
        print(f"\n=== STAGE 1: Workbook preparation ===")
        print(f"    Loading workbook from {self.parent_file}")
        self.wb = utils.xl.load_workbook(parent_file)
        print(f"    Loading read-only workbook from {self.parent_file}")
        self.wb_data = utils.xl.load_workbook(parent_file, data_only=True)
        print(f"    STAGE 1 COMPLETE")

    def main(self):
        print(f"\nStarting workbook modification")
        self.indicator_tables_creator = IndicatorTablesCreator(self.wb, self.wb_data, self.parent_file)
        self.index_sheet_creator = IndexSheetCreator(self.indicator_tables_creator)
        self.summary_sheet_creator = SummarySheetCreator(self.index_sheet_creator)
        self.pivot_sheet_creator = PivotSheetCreator(self.index_sheet_creator)
        print(f"\nWorkbook modification complete")
        self.wb.save(self.parent_file)
        print(f"Modified workbook saved as {self.parent_file}")

if __name__ == '__main__':
    test = OSBGilder("/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/test_chamber/TEST_singular.xlsx")
    test.main()