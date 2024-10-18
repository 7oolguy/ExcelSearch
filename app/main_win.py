import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from components.excel_manager import ExcelManager
from components.file_handler import load_user_data, save_user_data


class MainWin:
    def __init__(self) -> None:
        self._root = tk.Tk()
        self._root.title("Search Excel")
        self._root.geometry("700x1000")

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
            text="Select T"
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
            command= self.search_key_word,
            width=20
        )
        self.search_btn.grid(
            row=0,
            rowspan=2,
            column=5,
            columnspan=2,
            padx=px,
            pady=10,
            sticky="w"
            )

        self.search_result_count_label = ttk.Label(
            self._root,
            text="Row Count:"
            )
        self.search_result_count_label.grid(
            row=0,
            rowspan=2,
            column=7,
            columnspan=2,
            padx=px,
            pady=10,
            sticky='w')

        # VIEW
        self.table = TableView(self._root, self.excel.get_rows())
        self.table.create_table_widget()

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
        # Tables in pandas are not handled like in openpyxl; this is optional or can be removed
        self.table_combo.set("N/A")

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
            return
        else:
            # Use the pandas-based search function
            rows = self.excel.get_value_row(value)
            if rows:
                headers = self.excel.get_header()
                self.search_result_count_label.config(text=f'Count Row: {len(rows)}')
                self.table.set_table_data(headers, rows)
            else:
                print("No matching rows found.")

    def run(self):
        self._root.mainloop()

class TableView:
    def __init__(self, parent, original_data=None) -> None:
        self.parent = parent
        self.tree = None
        self.headers = []
        self.data = []
        self.original_data = original_data or []  # Store full unfiltered data here
        self.undo_stack = []
        self.redo_stack = []
        self.create_table_widget()

    def create_table_widget(self):
        self.tree = ttk.Treeview(self.parent, columns=self.headers, show='headings', selectmode='browse')
        self.tree.grid(row=3, column=0, columnspan=7, sticky='ns')

        # Vertical scrollbar
        scrollbar_y = ttk.Scrollbar(self.parent, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=3, column=7, sticky='ns')
        self.tree.configure(yscroll=scrollbar_y.set)

        # Horizontal Scrollbar
        scrollbar_x = ttk.Scrollbar(self.parent, orient="horizontal", command=self.tree.xview)
        scrollbar_x.grid(row=4, column=0, columnspan=7, sticky='ew')
        self.tree.configure(xscroll=scrollbar_x.set)

        self.tree.bind("<ButtonRelease-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)  # Double click

        # Add Go Back and Go Forth buttons
        self.back_button = tk.Button(self.parent, text="Go Back", command=self.go_back)
        self.back_button.grid(row=5, column=0, pady=10, padx=10, sticky="ew")

        self.forth_button = tk.Button(self.parent, text="Go Forth", command=self.go_forth)
        self.forth_button.grid(row=5, column=1, pady=10, padx=10, sticky="ew")

    def set_table_data(self, headers, data):
        if not headers or not data:
            print("Error: Headers or data are empty.")
            return

        if not self.original_data:
            self.original_data = data.copy()

        self.headers = headers
        self.data = data

        self.tree.delete(*self.tree.get_children())

        self.tree["columns"] = self.headers

        for col in self.headers:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)

        for row_data in self.data:
            self.tree.insert("", "end", values=row_data)

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.tree.identify_row(event.y)  # Get the clicked row's ID
            col_id = self.tree.identify_column(event.x)  # Get the clicked column's ID
            col_index = int(col_id.replace("#", "")) - 1  # Convert the column ID to an index
            row_index = self.tree.index(row_id)  # Get the index of the clicked row

            # Get the item from the tree at the clicked row
            item = self.tree.item(row_id, 'values')

            # Fetch the value in the specific column (cell)
            cell_value = item[col_index]  # Get the value at the clicked column

            print(f"Cell clicked at Row: {row_index}, Column: {col_index}, Value: {cell_value}")

            # Optionally, return or use the value in some way (for example, display it in a label or entry)
            return cell_value

    def on_double_click(self, event):
        """Double-click event handler to filter table data."""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.tree.identify_row(event.y)  # Get the clicked row's ID
            col_id = self.tree.identify_column(event.x)  # Get the clicked column's ID
            col_index = int(col_id.replace("#", "")) - 1  # Convert the column ID to an index

            # Get the item from the tree at the clicked row
            item = self.tree.item(row_id, 'values')
            cell_value = item[col_index]  # Get the value at the clicked column

            print(f"Double-clicked value: {cell_value}")

            # Call filter_by_value to filter the table based on this value
            self.filter_by_value(cell_value)

    def filter_by_value(self, value):
        """Filter table data to show only rows where the clicked value matches."""
        filtered_data = [row for row in self.original_data if value in row]  # Adjust the condition as needed for matching

        # Set the filtered data into the table
        self.set_table_data(self.headers, filtered_data)

    def go_back(self):
        """Go back to the previous table state."""
        if self.undo_stack:
            # Save current state to redo stack before undoing
            self.redo_stack.append(self.data.copy())

            # Pop the last state from the undo stack
            last_state = self.undo_stack.pop()

            # Apply the last state
            self.set_table_data(self.headers, last_state)

    def go_forth(self):
        """Go forth to the next table state."""
        if self.redo_stack:
            # Save current state to undo stack before redoing
            self.undo_stack.append(self.data.copy())

            # Pop the last state from the redo stack
            next_state = self.redo_stack.pop()

            # Apply the next state
            self.set_table_data(self.headers, next_state)
