import OSB_Gilder.back.script.utils as utils
from OSB_Gilder.back.script.indicator_tables_creator import IndicatorTablesCreator
from OSB_Gilder.back.script.index_sheet_creator import IndexSheetCreator
from OSB_Gilder.back.script.pivot_sheet_creator import PivotSheetCreator
from OSB_Gilder.back.script.summary_sheet_creator import SummarySheetCreator

class OSBGilder:
    def __init__(self, parent_file):
        self.parent_file = parent_file
        self.wb = utils.xl.load_workbook(parent_file)
        self.wb_data = utils.xl.load_workbook(parent_file, data_only=True)

    def main(self):
        self.indicator_tables_creator = IndicatorTablesCreator(self.wb, self.wb_data, self.parent_file)
        self.index_sheet_creator = IndexSheetCreator(self.indicator_tables_creator)
        self.summary_sheet_creator = SummarySheetCreator(self.index_sheet_creator)
        self.pivot_sheet_creator = PivotSheetCreator(self.index_sheet_creator)
        self.wb.save(self.parent_file)

if __name__ == '__main__':
    test = OSBGilder("test_chamber/TEST.xlsx")
    test.main()