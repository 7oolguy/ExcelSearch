import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from components.excel_manager import ExcelManager
from components.file_handler import load_user_data, save_user_data


class MainWin:
    def __init__(self) -> None:
        self._root = tk.Tk()
        self._root.title("Search Excel")
        self._root.geometry("1000x1000")

        self._root.grid_columnconfigure(5, weight=1)
        self._root.grid_rowconfigure(5, weight=1)

        self.excel = ExcelManager()
        self.settings = load_user_data() or {}

        self.file_path = None
        self.sheet_name = None

        self.create_widgets()
        self.load_user_data()

    def create_widgets(self):
        px = 10
        py = 0
        # GROUP 1
        self.excel_label = ttk.Label(
            self._root,
            text="Select Excel"
        )
        self.excel_label.grid(
            row=0,
            column=0,
            padx=px,
            pady=py,
            sticky="s")

        self.excel_btn = tk.Button(
            self._root,
            text="Select",
            command=self.select_excel
        )
        self.excel_btn.grid(
            row=1,
            column=0,
            padx=px,
            pady=py,
            sticky="n")

        # GROUP 2
        self.sheet_label = ttk.Label(
            self._root,
            text="Select Sheet"
        )
        self.sheet_label.grid(
            row=0,
            column=1,
            padx=px,
            pady=py,
            sticky="s")

        self.sheet_combo = ttk.Combobox(
            self._root,
            state="readonly"
        )
        self.sheet_combo.grid(
            row=1,
            column=1,
            padx=px,
            pady=py,
            sticky="n")

        # GROUP 3
        self.table_label = ttk.Label(
            self._root,
            text="Select Table"
        )
        self.table_label.grid(
            row=0,
            column=2,
            padx=px,
            pady=py,
            sticky="s")

        self.table_combo = ttk.Combobox(
            self._root,
            state="readonly"
        )
        self.table_combo.grid(
            row=1,
            column=2,
            padx=px,
            pady=py,
            sticky="n")

        # DIVISOR
        self.div = tk.Frame(
            self._root,
            bg="#2c2c2c",
            width=1,
            height=100
        )
        self.div.grid(
            row=0,
            rowspan=2,
            column=3,
            padx=px,
            pady=py,
            sticky="nsew")

        # GROUP 4 -> SEARCH
        self.search_label = ttk.Label(
            self._root,
            text="Search:"
        )
        self.search_label.grid(
            row=0,
            column=4,
            padx=px,
            pady=py,
            sticky="s")

        self.search_entry = tk.Entry(self._root)
        self.search_entry.grid(
            row=1,
            column=4,
            padx=px,
            pady=py,
            sticky="n")

        self.search_btn = tk.Button(
            self._root,
            text="Search",
            command=self.search_key_word,
            width=20
        )
        self.search_btn.grid(
            row=0,
            rowspan=2,
            column=5,
            columnspan=2,
            padx=px,
            pady=10,
            sticky="w")

    def select_excel(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
            )
        if self.file_path:
            self.excel.load(self.file_path)
            self.update_sheet_dropdown()
            self.save_user_data()

    def update_sheet_dropdown(self):
        self.sheet_combo['values'] = self.excel.get_sheet_names()
        if self.sheet_name:
            self.sheet_combo.set(self.sheet_name)
        else:
            self.sheet_combo.current(0)
        self.update_table_dropdown()

    def update_table_dropdown(self):
        tables = self.excel.get_table_names()
        if tables:
            self.table_combo['values'] = tables
            self.table_combo.current(0)
        else:
            self.table_combo.set("No Table Found")

    def save_user_data(self):
        data = {
            "file_path": self.file_path,
            "sheet_name": self.sheet_name
        }
        save_user_data(data)

    def load_user_data(self):
        if self.settings:
            self.file_path = self.settings.get("file_path", None)
            self.sheet_name = self.settings.get("sheet_name", None)
            if self.file_path:
                self.excel.load(self.file_path)
                self.update_sheet_dropdown

    def search_key_word(self):
        value = self.search_entry.get()
        if not value:
            pass
        else:
            rows = self.excel.get_value_row(value)
            print(rows)

    def run(self):
        self._root.mainloop()
