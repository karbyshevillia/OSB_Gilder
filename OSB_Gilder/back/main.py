from threading import Thread, Event
import OSB_Gilder.back.script.utils as utils
import time
from OSB_Gilder.back.script.indicator_tables_creator import IndicatorTablesCreator
from OSB_Gilder.back.script.index_sheet_creator import IndexSheetCreator
from OSB_Gilder.back.script.pivot_sheet_creator import PivotSheetCreator
from OSB_Gilder.back.script.summary_sheet_creator import SummarySheetCreator
from python_calamine import CalamineWorkbook

class OSBGilder(Thread):
    def __init__(self, parent_file, progress_var):
        self.start_time = time.perf_counter()
        super().__init__(target=self.main, daemon=True)
        print(f"Initialising OSBGilder object for {parent_file}")
        self.parent_file = parent_file
        self.progress_var = progress_var
        self.delay = 0.01
        # self.stop = Event()

    def update_progress(self, value):
        current = self.progress_var.get()
        new = current + value / 100.0
        self.progress_var.set(new)
        time.sleep(self.delay)

    def main(self):
        print(f"\n=== STAGE 1: Workbook preparation ===")
        print(f"    Loading workbook from {self.parent_file}")

        self.wb = utils.xl.load_workbook(self.parent_file)
        self.update_progress(15)

        print(f"    Loading read-only workbook from {self.parent_file}")

        self.wb_data = CalamineWorkbook.from_path(self.parent_file)
        # self.wb_data = utils.xl.load_workbook(self.parent_file, data_only=True)
        self.update_progress(10)

        print(f"    STAGE 1 COMPLETE")
        print(f"\nStarting workbook modification")

        self.indicator_tables_creator = IndicatorTablesCreator(self.wb,
                                                               self.wb_data,
                                                               self.parent_file,
                                                               self.progress_var)
        # self.update_progress(30)

        self.index_sheet_creator = IndexSheetCreator(self.indicator_tables_creator)
        # self.update_progress(10)

        self.summary_sheet_creator = SummarySheetCreator(self.index_sheet_creator)
        # self.update_progress(10)

        self.pivot_sheet_creator = PivotSheetCreator(self.index_sheet_creator)
        # self.update_progress(10)

        print(f"\nWorkbook modification complete")

        self.wb.save(self.parent_file)
        self.update_progress(5)

        print(f"Modified workbook saved as {self.parent_file}")
        self.finish_time = time.perf_counter()
        print(f"Finished in {(self.finish_time - self.start_time):.2f} second(s)")

if __name__ == '__main__':
    test = OSBGilder("/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/test_chamber/TEST_singular.xlsx")
    test.main()