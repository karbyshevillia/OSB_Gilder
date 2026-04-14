from tkinter import ttk, Tk, filedialog, messagebox, PhotoImage
import tkinter as tk
import os
import shutil
import json
# from OSB_Gilder.back.account_rows import BalanceSheet, process_excel
# from OSB_Gilder.back.combined_script import BalanceSheet, process_excel
from OSB_Gilder.back.account_rows import BalanceSheet
from OSB_Gilder.back.main import OSBGilder


class IndicatorSelectorWindow:
    """TTK-based selector for BalanceSheet indicators"""

    DEFAULT_CONFIG_FILE = "default_indicators.json"

    def __init__(self, root, parent_gui):
        self.root = tk.Toplevel(root)
        self.root.title("Select Indicators and Column")
        self.root.geometry("650x500")
        self.root.resizable(True, True)
        self.parent_gui = parent_gui

        # Store selections per column: {column_name: {indicator_name: bool}}
        self.selections_per_column = {}
        self.current_column = None
        self.load_default_indicators()

        # tk.BooleanVar for each indicator
        self.vars = {}

        self.setup_ui()

    def setup_ui(self):
        # Column Selection at Top
        column_frame = ttk.Frame(self.root)
        column_frame.pack(fill="x", padx=12, pady=10)

        ttk.Label(column_frame, text="Select Column for Calculation:", font=("Segoe UI", 11, "bold")).pack(side="left",
                                                                                                           padx=5)

        self.column_var = tk.StringVar(value=self.current_column or BalanceSheet.LOGICAL_COLUMNS[0])
        column_dropdown = ttk.Combobox(
            column_frame,
            textvariable=self.column_var,
            values=BalanceSheet.LOGICAL_COLUMNS,
            state="readonly",
            width=25
        )
        column_dropdown.pack(side="left", padx=5)
        column_dropdown.bind("<<ComboboxSelected>>", self.on_column_changed)

        # Separator
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", pady=5)

        # Header for indicators
        header = ttk.Label(self.root, text="Select Indicators", font=("Segoe UI", 15, "bold"), padding=10)
        header.pack(fill="x", padx=12, pady=(10, 6))

        # Bottom buttons - TESTING: placed ABOVE indicators
        btn_frame = ttk.Frame(self.root, padding=10, relief="solid", borderwidth=2)
        btn_frame.pack(fill="x", padx=12, pady=10)

        ttk.Button(btn_frame, text="Вибрати всі", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Видалити всі", command=self.clear_all).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Зберегти", command=self.save_default_indicators).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Закрити", command=self.on_close).pack(side="right", padx=4)

        # Scrollable frame for indicators
        frame = ttk.Frame(self.root, padding=10, height=300)
        frame.pack(fill="x", expand=True, padx=12, pady=6)
        frame.pack_propagate(False)  # Prevent frame from shrinking to fit content

        canvas = tk.Canvas(frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollable = ttk.Frame(canvas)

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable, anchor="nw")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Create a ttk.Checkbutton for each indicator
        self.checks = {}
        for indicator in BalanceSheet.INDICATORS:
            name = indicator.name
            var = tk.BooleanVar(value=False)
            chk = ttk.Checkbutton(scrollable, text=name, variable=var, padding=8)
            chk.pack(fill="x", padx=8, pady=4, anchor="w")
            self.vars[name] = var
            self.checks[name] = chk

        # Load initial column selections
        self.current_column = self.column_var.get()
        self.load_column_selections()

    def on_column_changed(self, event=None):
        """Called when user changes the column selection"""
        # Save current selections before switching
        self.save_current_column_selections()

        # Update current column
        self.current_column = self.column_var.get()

        # Load selections for new column
        self.load_column_selections()

    def save_current_column_selections(self):
        """Save current indicator selections for the current column"""
        if self.current_column:
            self.selections_per_column[self.current_column] = {
                name: var.get() for name, var in self.vars.items()
            }

    def load_column_selections(self):
        """Load indicator selections for the current column"""
        if self.current_column in self.selections_per_column:
            # Restore saved selections
            for name, var in self.vars.items():
                var.set(self.selections_per_column[self.current_column].get(name, False))
        else:
            # Clear all selections for new column
            for var in self.vars.values():
                var.set(False)

    def select_all(self):
        for name, var in self.vars.items():
            var.set(True)

    def clear_all(self):
        for name, var in self.vars.items():
            var.set(False)

    def save_default_indicators(self):
        # Save current column selections first
        self.save_current_column_selections()

        try:
            with open(self.DEFAULT_CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.selections_per_column, f, ensure_ascii=False, indent=2)

            # Count total selections across all columns
            total_count = sum(
                sum(1 for selected in column_data.values() if selected)
                for column_data in self.selections_per_column.values()
            )
            messagebox.showinfo("Saved",
                                f"Saved selections for {len(self.selections_per_column)} column(s) with {total_count} total indicator(s).")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save defaults: {e}")

    def load_default_indicators(self):
        if os.path.exists(self.DEFAULT_CONFIG_FILE):
            try:
                with open(self.DEFAULT_CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.selections_per_column = json.load(f)
                    # Set current column to first available or first in list
                    if self.selections_per_column:
                        self.current_column = list(self.selections_per_column.keys())[0]
                    else:
                        self.current_column = BalanceSheet.LOGICAL_COLUMNS[0]
            except Exception:
                self.selections_per_column = {}
                self.current_column = BalanceSheet.LOGICAL_COLUMNS[0]
        else:
            self.selections_per_column = {}
            self.current_column = BalanceSheet.LOGICAL_COLUMNS[0]

    def get_selected_indicators(self):
        """Returns dict of {column: [indicator_names]}"""
        return {
            column: [name for name, selected in indicators.items() if selected]
            for column, indicators in self.selections_per_column.items()
        }

    def on_close(self):
        # Save current selections before closing
        self.save_current_column_selections()

        # refresh parent status and close
        if self.parent_gui:
            self.parent_gui.update_indicators_status()
        self.root.destroy()


class GUI:
    def __init__(self, root):
        self.root = root
        icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        self.root.iconphoto(True, PhotoImage(file=icon_path))
        self.root.title("OSB Analyzer")
        self.root.configure(bg="#bdbdbd")
        self.root.resizable(False, False)

        self.style = ttk.Style()
        self.style.theme_use("default")

        FONT = ("Segoe UI", 11)
        TITLE_FONT = ("Segoe UI", 15, "bold")

        self.style.configure("App.TFrame", background="#bdbdbd")
        self.style.configure("Card.TFrame", background="white")

        self.style.configure(
            "Title.TLabel",
            background="#bdbdbd",
            font=TITLE_FONT,
            foreground="#1f2937"
        )

        self.style.configure(
            "CardText.TLabel",
            background="white",
            font=FONT,
            foreground="#374151"
        )

        self.style.configure(
            "Primary.TButton",
            font=FONT,
            padding=10,
            background="#2563eb",
            foreground="white"
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#1d4ed8")]
        )

        self.style.configure(
            "Success.TButton",
            font=FONT,
            padding=10,
            background="#16a34a",
            foreground="white"
        )
        self.style.map(
            "Success.TButton",
            background=[("active", "#15803d")]
        )

        self.style.configure(
            "Secondary.TButton",
            font=FONT,
            padding=10
        )

        self.input_file_path = None  # шлях до вхідного Excel
        self.output_file_path = None  # шлях до результату (створюється в іншому файлі)

    def open_indicator_selector(self):
        """Open the indicator selector window"""
        IndicatorSelectorWindow(self.root, self)

    def update_indicators_status(self):
        """Update the indicators status label"""
        config_file = IndicatorSelectorWindow.DEFAULT_CONFIG_FILE
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    config_data = json.load(f)

                    # Count total selections across all columns
                    total_count = 0
                    column_count = 0
                    for column, indicators in config_data.items():
                        if isinstance(indicators, dict):
                            selected = sum(1 for v in indicators.values() if v)
                            if selected > 0:
                                total_count += selected
                                column_count += 1

                    if total_count > 0:
                        self.indicators_status_label.config(
                            text=f"Вибрано: {column_count} колонок, {total_count} індикаторів"
                        )
                    else:
                        self.indicators_status_label.config(
                            text="Індикатори не вибрано"
                        )
            except Exception as e:
                print(f"Error reading indicators: {e}")
                self.indicators_status_label.config(
                    text="Індикатори не вибрано"
                )
        else:
            self.indicators_status_label.config(
                text="Індикатори не вибрано"
            )

    def center_window(self, width, height):
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()

        x = (screen_w - width) // 2
        y = (screen_h - height) // 2

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def run(self):
        self.center_window(760, 550)

        container = ttk.Frame(self.root, style="App.TFrame")
        container.pack(fill="both", expand=True, padx=30, pady=25)

        ttk.Label(container, text="Файл", style="Title.TLabel").pack(anchor="w")

        input_card = ttk.Frame(container, style="Card.TFrame")
        input_card.pack(fill="x", pady=12)

        input_inner = ttk.Frame(input_card, style="Card.TFrame")
        input_inner.pack(fill="x", padx=24, pady=24)

        ttk.Button(
            input_inner,
            text="Завантажити файл",
            style="Secondary.TButton",
            command=self.load_file,
            width=25
        ).pack(side="left")

        ttk.Button(
            input_inner,
            text="Опрацювати",
            style="Primary.TButton",
            command=self.process_file,
            width=16
        ).pack(side="right")

        # Indicators Section
        ttk.Label(container, text="Індикатори", style="Title.TLabel").pack(anchor="w", pady=(25, 0))

        indicators_card = ttk.Frame(container, style="Card.TFrame")
        indicators_card.pack(fill="x", pady=12)

        indicators_inner = ttk.Frame(indicators_card, style="Card.TFrame")
        indicators_inner.pack(fill="x", padx=24, pady=24)

        ttk.Button(
            indicators_inner,
            text="Вибрати індикатори",
            style="Secondary.TButton",
            command=self.open_indicator_selector,
            width=25
        ).pack(side="left")

        self.indicators_status_label = ttk.Label(
            indicators_inner,
            text="Індикатори не вибрано",
            style="CardText.TLabel"
        )
        self.indicators_status_label.pack(side="right")

        # Update status on startup
        self.update_indicators_status()

        ttk.Label(container, text="Вихід", style="Title.TLabel").pack(anchor="w", pady=(25, 0))

        output_card = ttk.Frame(container, style="Card.TFrame")
        output_card.pack(fill="x", pady=12)

        output_inner = ttk.Frame(output_card, style="Card.TFrame")
        output_inner.pack(fill="x", padx=24, pady=24)

        self.status_label = ttk.Label(
            output_inner,
            text="Файл ще не завантажено",
            style="CardText.TLabel"
        )
        self.status_label.pack(side="left")

        ttk.Button(
            output_inner,
            text="Завантажити",
            style="Success.TButton",
            command=self.save_result,
            width=16
        ).pack(side="right")

        self.root.mainloop()

    def load_file(self):
        """Вибір ВХІДНОГО Excel-файлу"""

        file_path = filedialog.askopenfilename(
            title="Оберіть Excel файл",
            filetypes=[("Excel файли", "*.xlsx")]
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            messagebox.showerror(
                "Помилка",
                "Дозволено лише Excel файли (.xlsx)"
            )
            return

        self.input_file_path = file_path
        self.status_label.config(
            text=f"Завантажено файл: {os.path.basename(file_path)}"
        )

    def process_file(self):
        """

        Цей метод НЕ містить реальної логіки.
        Тут потрібно:
        - викликати функцію з іншого файлу (наприклад processor.py)
        - передати шлях до input_file_path
        - отримати шлях до створеного Excel-файлу
        - записати його в self.output_file_path

        from processor import process_excel
        self.output_file_path = process_excel(self.input_file_path)
        """

        if not self.input_file_path:
            messagebox.showwarning(
                "Файл не вибрано",
                "Спочатку завантажте Excel файл."
            )
            return

        gilder = OSBGilder(self.input_file_path)
        gilder.main()
        self.output_file_path = self.input_file_path

        self.status_label.config(
            text="Файл опрацьовано (очікується створення результату)"
        )

    def save_result(self):
        """Збереження РЕЗУЛЬТАТУ через Save As"""

        if not self.output_file_path or not os.path.exists(self.output_file_path):
            messagebox.showwarning(
                "Немає результату",
                "Результат ще не створено."
            )
            return

        save_path = filedialog.asksaveasfilename(
            title="Зберегти файл як",
            defaultextension=".xlsx",
            initialfile="OSB_Result.xlsx",
            filetypes=[("Excel файли", "*.xlsx")]
        )

        if save_path:
            shutil.copyfile(self.output_file_path, save_path)


if __name__ == "__main__":
    root = Tk()
    app = GUI(root)
    app.run()