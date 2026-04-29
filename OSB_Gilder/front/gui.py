import os, shutil
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, PhotoImage, messagebox
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from OSB_Gilder.back.main import OSBGilder
from threading import Event


class CTkWithDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class GilderUI(CTkWithDnD):
    def __init__(self):
        super().__init__()

        # Window Setup: Increased height slightly to accommodate the progress bar safely
        self.title("OSB Gilder")
        self.geometry("750x750")
        self.configure(fg_color="#F7F9FC")
        self.resizable(False, False)

        # Variables
        self.processing_mode = ctk.StringVar(value="amend")
        self.selected_file_path = None

        # Spinner Variables
        self.spinner_frames = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        self.spinner_idle = '—'
        self.spinner_done = '✓'
        self.spinner_cancelled = '✗'
        self.spinner_idx = 0
        self.is_spinning = False

        # Build UI Components
        self._build_header()
        self._build_dropzone()
        self._build_options()
        self._build_footer()

        # Miscellany
        # icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        # self.iconphoto(True, PhotoImage(file=icon_path))

        # # Processor variable
        self.worker_thread = None

        # Cancel event
        self.cancel_event = Event()

    def _build_header(self):
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(pady=(40, 20), fill="x")

        logo_label = ctk.CTkLabel(header_frame, text="📗", font=("Arial", 40), text_color="#2A9D6A")
        logo_label.pack()

        title_label = ctk.CTkLabel(header_frame, text="OSB Gilder",
                                   font=("Segoe UI", 28, "bold"), text_color="#1E293B")
        title_label.pack(pady=(5, 0))

        subtitle_label = ctk.CTkLabel(header_frame, text="Обробка та модифікація ОСЗ НБУ",
                                      font=("Segoe UI", 14), text_color="#64748B")
        subtitle_label.pack()

    def _build_dropzone(self):
        self.drop_frame = ctk.CTkFrame(self, fg_color="white",
                                       border_color="#D2D6DC", border_width=2, corner_radius=15,
                                       cursor="hand2")
        self.drop_frame.pack(pady=20, padx=80, fill="both", expand=True)

        self.inner_frame = ctk.CTkFrame(self.drop_frame, fg_color="transparent", cursor="hand2")
        self.inner_frame.place(relx=0.5, rely=0.5, anchor="center")

        self.cloud_icon = ctk.CTkLabel(self.inner_frame, text="☁️", font=("Arial", 50),
                                       text_color="#94A3B8", cursor="hand2")
        self.cloud_icon.pack(pady=(0, 10))

        self.drag_text = ctk.CTkLabel(self.inner_frame, text="Перетягніть сюди .xlsx файл для обробки",
                                      font=("Segoe UI", 18), text_color="#1E293B", cursor="hand2")
        self.drag_text.pack()

        self.browse_text = ctk.CTkLabel(self.inner_frame, text="або натисніть для пошуку",
                                        font=("Segoe UI", 14), text_color="#2563EB", cursor="hand2")
        self.browse_text.pack(pady=(5, 0))

        clickable_elements = [self.drop_frame, self.inner_frame, self.cloud_icon,
                              self.drag_text, self.browse_text]

        for element in clickable_elements:
            element.bind("<Button-1>", self._browse_files)
            element.bind("<Enter>", self._highlight_dropzone)
            element.bind("<Leave>", self._unhighlight_dropzone)
            element.drop_target_register(DND_FILES)
            element.dnd_bind('<<Drop>>', self._handle_file_drop)
            element.dnd_bind('<<DragEnter>>', self._highlight_dropzone)
            element.dnd_bind('<<DragLeave>>', self._unhighlight_dropzone)

    def _build_options(self):
        options_container = ctk.CTkFrame(self, fg_color="transparent")
        options_container.pack(pady=(10, 20), padx=80, fill="x")

        choose_label = ctk.CTkLabel(options_container, text="Оберіть спосіб обробки:",
                                    font=("Segoe UI", 14, "bold"), text_color="#1E293B")
        choose_label.pack(pady=(0, 15))

        cards_frame = ctk.CTkFrame(options_container, fg_color="transparent")
        cards_frame.pack(fill="x")

        cards_frame.grid_columnconfigure((0, 1), weight=1, uniform="row")

        self.amend_card = self._create_option_card(
            cards_frame, "amend", "✏️", "Оригінал", "Робити зміни в оригінальному файлі", 0, pad_x=(0, 10)
        )
        self.copy_card = self._create_option_card(
            cards_frame, "copy", "📄", "Копія", "Робити зміни в окремому файлі-копії", 1, pad_x=(10, 0)
        )

        self._update_card_styles()

    def _create_option_card(self, parent, value, icon, title, subtitle, col, pad_x):
        card = ctk.CTkFrame(parent, fg_color="white", border_color="#E2E8F0",
                            border_width=1, corner_radius=8, cursor="hand2")
        card.grid(row=0, column=col, padx=pad_x, sticky="ew")
        card.bind("<Button-1>", lambda e, v=value: self._select_option(v))
        card.grid_columnconfigure(2, weight=1)

        rb = ctk.CTkRadioButton(card, text="", variable=self.processing_mode, value=value,
                                width=20, fg_color="#2A9D6A", border_color="#94A3B8",
                                command=self._update_card_styles)
        rb.grid(row=0, column=0, padx=(15, 10), pady=20)
        rb.bind("<Button-1>", lambda e, v=value: self._select_option(v))

        icon_label = ctk.CTkLabel(card, text=icon, font=("Arial", 20))
        icon_label.grid(row=0, column=1, padx=(0, 10))
        icon_label.bind("<Button-1>", lambda e, v=value: self._select_option(v))

        text_frame = ctk.CTkFrame(card, fg_color="transparent")
        text_frame.grid(row=0, column=2, sticky="w", pady=15)
        text_frame.bind("<Button-1>", lambda e, v=value: self._select_option(v))

        title_lbl = ctk.CTkLabel(text_frame, text=title, font=("Segoe UI", 14, "bold"), text_color="#1E293B")
        title_lbl.pack(anchor="w")
        title_lbl.bind("<Button-1>", lambda e, v=value: self._select_option(v))

        sub_lbl = ctk.CTkLabel(text_frame, text=subtitle, font=("Segoe UI", 12), text_color="#64748B",
                               wraplength=180, justify="left")
        sub_lbl.pack(anchor="w")
        sub_lbl.bind("<Button-1>", lambda e, v=value: self._select_option(v))

        return card

    def _select_option(self, value):
        self.processing_mode.set(value)
        self._update_card_styles()

    def _update_card_styles(self):
        self.amend_card.configure(border_color="#E2E8F0", border_width=1)
        self.copy_card.configure(border_color="#E2E8F0", border_width=1)
        if self.processing_mode.get() == "amend":
            self.amend_card.configure(border_color="#2A9D6A", border_width=1)
        else:
            self.copy_card.configure(border_color="#2A9D6A", border_width=1)

    def _on_progress_update(self, *args):
        if not self.cancel_event.is_set():
            # 1. Get the current progress
            current_progress = self.progress_var.get()

            # 2. Update your status label (replacing your lambda)
            self.status_var.set(f"{(current_progress * 100):.1f}%")

            # 3. Stop the spinner if we hit 100% (1.0)
            if current_progress >= 1.0:
                self.stop_spinner()
                self.reset_process_button()
        else:
            self.status_var.set("0.0%")

    def _build_footer(self):
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        # Reduced the bottom padding from 40 to 20 to preserve layout balance
        footer_frame.pack(pady=(10, 20), fill="x")

        # --- Cosmetic Progress Bar ---
        self.progress_var = ctk.DoubleVar(value=0.0)

        # NEW: Container to hold the bar and the spinner side-by-side
        progress_container = ctk.CTkFrame(footer_frame, fg_color="transparent")
        progress_container.pack(pady=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(progress_container, variable=self.progress_var, width=300, height=10,
                                               progress_color="#2A9D6A", fg_color="#E2E8F0")
        self.progress_bar.set(0)
        self.progress_bar.pack(side="left")

        # NEW: Spinner Label (Fixed width of 25 so it doesn't push elements around while spinning)
        self.spinner_label = ctk.CTkLabel(progress_container, text=self.spinner_idle, font=("Segoe UI", 16, "bold"),
                                          text_color="#2A9D6A", width=25)
        self.spinner_label.pack(side="left", padx=(10, 0))

        self.status_var = ctk.StringVar(value="0.0%")
        self.status_label = ctk.CTkLabel(footer_frame,
                                         # text="0%",
                                         font=("Segoe UI", 12, "bold"),
                                         text_color="#64748B",
                                         textvariable=self.status_var)
        self.status_label.pack(pady=(0, 10))

        # self.progress_var.trace_add("write",
        #                             lambda *args: self.status_var.set(f"{(self.progress_var.get() * 100):.1f}%"))
        self.progress_var.trace_add("write", self._on_progress_update)

        # Set an unmistakably large height of 75 so you can see it working
        self.process_btn = ctk.CTkButton(footer_frame, text="⚙️ Обробити",
                                    font=("Segoe UI", 20, "bold"),
                                    fg_color="#2A9D6A", hover_color="#218056",
                                    width=300, height=50, corner_radius=8,
                                    command=self._process_file)

        self.process_btn.pack(pady=10)

    def reset_process_button(self):
        """Reverts the button back to the green Process state."""
        self.process_btn.configure(
            text="⚙️ Обробити",
            fg_color="#2A9D6A",
            hover_color="#218056",
            command=self._process_file
        )

    def cancel_processing(self):
        """Instantly resets the UI and flags the ghost thread to die."""
        if self.worker_thread and self.worker_thread.is_alive():
            print("Cancelling process...")
            self.cancel_event.set()  # Trip the flag!

            # Cut ties with the thread
            self.worker_thread = None

            # Instantly reset the UI to make it look like it stopped
            self.stop_spinner()
            # self.progress_var.set(0.0)
            self.drag_text.configure(text=f"Обробку файлу {os.path.basename(self.selected_file_path)} скасовано.", text_color="#EAB308")

            # Revert the button back to green
            self.reset_process_button()

    def start_spinner(self):
        self.is_spinning = True
        self._animate_spinner()

    def stop_spinner(self):
        self.is_spinning = False
        if not self.cancel_event.is_set():
            spinner_label = self.spinner_done
            spinner_label_colour = "#2A9D6A"
        else:
            spinner_label = self.spinner_cancelled
            spinner_label_colour = "#EF4444"
        self.spinner_label.configure(text=spinner_label, text_color=spinner_label_colour)

    def _animate_spinner(self):
        if self.is_spinning:
            self.spinner_label.configure(text=self.spinner_frames[self.spinner_idx])
            self.spinner_idx = (self.spinner_idx + 1) % len(self.spinner_frames)
            self.after(50, self._animate_spinner)

    def _highlight_dropzone(self, event=None):
        self.drop_frame.configure(fg_color="#F1F5F9")

    def _unhighlight_dropzone(self, event=None):
        try:
            x, y = self.winfo_pointerx(), self.winfo_pointery()
            x1 = self.drop_frame.winfo_rootx()
            y1 = self.drop_frame.winfo_rooty()
            x2 = x1 + self.drop_frame.winfo_width()
            y2 = y1 + self.drop_frame.winfo_height()

            if not (x1 <= x <= x2 and y1 <= y <= y2):
                self.drop_frame.configure(fg_color="white")
        except:
            self.drop_frame.configure(fg_color="white")

    def _browse_files(self, event=None):
        file_path = filedialog.askopenfilename(
            title="Оберіть файл Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self._update_ui_with_file(file_path)

    def _handle_file_drop(self, event):
        self.drop_frame.configure(fg_color="white")
        file_path = event.data.strip('{}')

        if file_path.lower().endswith(('.xlsx', '.xls')):
            self._update_ui_with_file(file_path)
        else:
            self.drag_text.configure(text="Цей тип файлу не підтримується. Будь ласка оберіть файл Excel.",
                                     text_color="red")

    def _update_ui_with_file(self, file_path):
        self.selected_file_path = file_path
        filename = os.path.basename(file_path)

        self.cloud_icon.configure(text="✅", text_color="#2A9D6A")
        self.drag_text.configure(text=f"Завантажено: {filename}", text_color="#1E293B", font=("Segoe UI", 16, "bold"))
        self.browse_text.configure(text="Натисніть, щоб обрати інший файл.")
        self.drop_frame.configure(border_color="#2A9D6A")

    # def _on_process_press(self):
    #     if self.processor is not None and self.processor.is_alive():
    #         file_name = self.processor.parent_file
    #         confirmation = messagebox.askokcancel(title="Підтвердити операцію", message=f"Триває обробка файлу {file_name}."
    #                                                                                     f"\nПочаток обробки нового файлу означатиме"
    #                                                                                     f"\nпередчасне переривання обробки вищевказаного"
    #                                                                                     f"\nфайлу. Чи бажаєте продовжити?")
    #         if confirmation:

    def _process_file(self):
        if not self.selected_file_path:
            self.drag_text.configure(text="⚠️ Будь ласка, спершу оберіть файл!", text_color="#EAB308")
            return

        self.progress_var.set(0.0)
        self.start_spinner()

        # CLEAR THE CANCEL EVENT
        self.cancel_event.clear()

        # TURN THE BUTTON RED AND CHANGE IT TO CANCEL
        self.process_btn.configure(
            text="❌ Скасувати",
            fg_color="#EF4444",  # Red color
            hover_color="#DC2626",  # Darker red on hover
            command=self.cancel_processing
        )

        mode = self.processing_mode.get()
        if mode == "amend":
            self.worker_thread = OSBGilder(self.selected_file_path, self.progress_var, self.cancel_event)
            self.worker_thread.daemon = True
            self.worker_thread.start()

        elif mode == "copy":
            file_path = filedialog.asksaveasfilename(
                title="Оберіть назву обробленого файлу",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx *.xls")],
                initialfile=Path(self.selected_file_path).stem + "_modified"
            )
            # If the user cancels the file dialog, revert the button and return early
            if not file_path:
                self.reset_process_button()
                self.stop_spinner()
                return

            shutil.copyfile(self.selected_file_path, file_path)
            self.worker_thread = OSBGilder(file_path, self.progress_var, self.cancel_event)
            self.worker_thread.daemon = True
            self.worker_thread.start()

        print(f"Виконується обробка файлу '{self.selected_file_path}' у режимі '{mode}'.")


if __name__ == "__main__":
    app = GilderUI()
    app.mainloop()