from threading import Thread, Event
import OSB_Gilder.back.script.utils as utils
import time
from OSB_Gilder.back.script.indicator_tables_creator import IndicatorTablesCreator
from OSB_Gilder.back.script.index_sheet_creator import IndexSheetCreator
from OSB_Gilder.back.script.pivot_sheet_creator import PivotSheetCreator
from OSB_Gilder.back.script.summary_sheet_creator import SummarySheetCreator
from python_calamine import CalamineWorkbook
import customtkinter as ctk

class OSBGilder(Thread):
    def __init__(self, parent_file, ind_file, banks_file, progress_var, cancel_event):
        self.start_time = time.perf_counter()
        super().__init__(target=self.run, daemon=True)
        print(f"Initialising OSBGilder object for {parent_file}")
        self.parent_file = parent_file
        self.ind_file = ind_file
        self.banks_file = banks_file

        self.progress_var = progress_var
        self.delay = 0.01
        self.cancel_event = cancel_event

    def update_progress(self, value):
        if not self.cancel_event.is_set():
            current = self.progress_var.get()
            new = current + value / 100.0
            self.progress_var.set(new)
            time.sleep(self.delay)

    def run(self):
        print(f"\n=== STAGE 1: Workbook preparation ===")
        print(f"    Loading workbook from {self.parent_file}")

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        self.wb = utils.xl.load_workbook(self.parent_file)
        self.update_progress(15)

        print(f"    Loading read-only workbook from {self.parent_file}")

        self.wb_data = CalamineWorkbook.from_path(self.parent_file)
        self.update_progress(10)

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        print(f"    STAGE 1 COMPLETE")
        print(f"\nStarting workbook modification")

        self.indicator_tables_creator = IndicatorTablesCreator(self.wb,
                                                               self.wb_data,
                                                               self.parent_file,
                                                               progress_var=self.progress_var,
                                                               cancel_event=self.cancel_event,
                                                               ind_file=self.ind_file,
                                                               banks_file=self.banks_file)

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        self.index_sheet_creator = IndexSheetCreator(self.indicator_tables_creator)

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        self.summary_sheet_creator = SummarySheetCreator(self.index_sheet_creator)

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        self.pivot_sheet_creator = PivotSheetCreator(self.index_sheet_creator)

        if self.cancel_event.is_set():
            print(f"Aborting modification of {self.parent_file}...")
            return

        print(f"\nWorkbook modification complete")

        self.wb.save(self.parent_file)
        self.update_progress(5)

        print(f"Modified workbook saved as {self.parent_file}")
        self.finish_time = time.perf_counter()
        print(f"Finished in {(self.finish_time - self.start_time):.2f} second(s)")

if __name__ == '__main__':
    rt = ctk.CTk()
    pv = ctk.DoubleVar()
    c = Event()
    test = OSBGilder("/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/test_chamber/OSB_bank_2026-03-01.xlsx",
                     progress_var=pv,
                     cancel_event=c,
                     ind_file="/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/back/testing/ind.csv",
                     banks_file="/Users/illiaknu/Desktop/OSB_Gilder/OSB_Gilder/back/testing/banks.json")
    test.run()