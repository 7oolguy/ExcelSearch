import openpyxl
import openpyxl.utils
import openpyxl.utils.exceptions
import openpyxl.worksheet
import openpyxl.worksheet.worksheet


class ExcelManager:
    def __init__(self, file_path=None, sheet_name=None):
        self._file_path = file_path
        self._sheet_name = sheet_name
        self._workbook = None
        self._sheet = None

    def __str__(self) -> str:
        return self.file_path

    def __eq__(self, value: object) -> bool:
        return self.file_path == value or self.sheet_name == value

    @property
    def file_path(self):
        return self._file_path

    @file_path.setter
    def file_path(self, file_path):
        if isinstance(file_path, str):
            self._file_path = file_path
        else:
            raise ValueError("File Path needs to be a string.")

    @property
    def sheet_name(self):
        return self._sheet_name

    @sheet_name.setter
    def sheet_name(self, sheet_name):
        if isinstance(sheet_name, str):
            self._sheet_name = sheet_name
        else:
            raise ValueError("Sheet Name needs to be a string.")

    @property
    def workbook(self):
        return self._workbook

    @workbook.setter
    def workbook(self, workbook):
        if isinstance(workbook, openpyxl.Workbook):
            self._workbook = workbook
        else:
            raise ValueError("The value must be an instance of openpyxl Workbook.")

    @property
    def sheet(self):
        return self._sheet

    @sheet.setter
    def sheet(self, sheet):
        if isinstance(sheet, openpyxl.worksheet.worksheet.Worksheet):
            self._sheet = sheet
        else:
            raise ValueError("The value must be an instance of openpyxl Worksheet")

    def load(self, file_path: str, sheet_name: str = None) -> None:
        try:
            self.file_path = file_path
            self.workbook = openpyxl.load_workbook(file_path)
            if sheet_name:
                if sheet_name in self.workbook.sheetnames:
                    self.sheet_name = sheet_name
                else:
                    raise ValueError("Sheet does not exists in the workbook.")
            else:
                self.sheet_name = self.workbook.sheetnames[0]
            self.sheet = self.workbook[self.sheet_name]
        except FileNotFoundError:
            raise FileNotFoundError(f"The file '{file_path}' was not found.")
        except openpyxl.utils.exceptions.InvalidFileException:
            raise ValueError("The file provided is not a valid Excel file.")

    def save(self):
        self.workbook.save(self.file_path)

    def get_sheet_names(self) -> list:
        if self.workbook:
            return self.workbook.sheetnames
        else:
            raise ValueError("No workbook is currently loaded.")

    def get_table_names(self) -> list:
        if self.sheet:
            return list(self.sheet.tables.keys())
        else:
            raise ValueError("No sheet is currently loaded.")

    def change_cell_value(self, row: int, col: int, value: str):
        if not self.sheet:
            raise ValueError("No sheet is currently loaded")

        if row and col:
            self.sheet.cell(row=row, column=col).value = value
        else:
            raise ValueError("You must provide the row number and the column number.")

        self.save()

    def get_header(self, table_name: str = None) -> list:
        if not self.sheet:
            raise ValueError("No sheet is currently loaded.")

        if table_name:
            if table_name not in self.sheet.tables:
                raise ValueError(f"Table '{table_name}' does not exists in the sheet.")

            table = self.sheet.tables[table_name]
            table_range = self.sheet[table.ref]
            header = [cell.value for cell in next(iter(table_range))]
        else:
            header = [cell for cell in next(self.sheet.iter_rows(values_only=True))]

        return header

    def get_rows(self, table_name: str = None, include_header: bool = True) -> list:
        if not self.sheet:
            raise ValueError("No sheet is currently loaded")

        rows = []

        if table_name:
            if table_name not in self.sheet.tables:
                raise ValueError(f"Table '{table_name}' does not exist in the sheet.")

            table = self.sheet.tables[table_name]
            table_range = self.sheet[table.ref]

            for i, row in enumerate(table_range):
                if not include_header and i == 0:
                    continue  # Skip header row if not included
                rows.append([cell.value for cell in row])

        else:
            for i, row in enumerate(self.sheet.iter_rows(values_only=True)):
                if not include_header and i == 0:
                    continue  # Skip header row if not included
                rows.append(list(row))

        return rows

    def get_value_row(self, value: str, table_name: str = None) -> list[list]:
        rows = self.get_rows(table_name=table_name, include_header=False)
        value_low = value.lower()
        matched_rows = []
        if rows:
            for row in rows:
                if any(value_low in str(item).lower() for item in row):
                    matched_rows.append(row)

        return matched_rows
