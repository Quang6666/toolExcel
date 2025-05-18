import xlwings as xw

class ExcelManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.app = xw.App(visible=False, add_book=False)
        self.wb = self.app.books.open(file_path)
        self.sheet_cache = {}
        self.last_empty_row_cache = {}  # {sheet_name: last_empty_row}

    def get_sheet(self, sheet_name):
        if sheet_name not in self.sheet_cache:
            self.sheet_cache[sheet_name] = self.wb.sheets[sheet_name]
        return self.sheet_cache[sheet_name]

    def get_last_empty_row(self, sheet_name, start_row, start_col, num_fields):
        cache_key = (sheet_name, start_row, start_col, num_fields)
        if cache_key in self.last_empty_row_cache:
            return self.last_empty_row_cache[cache_key]
        ws = self.get_sheet(sheet_name)
        row = start_row
        while ws.range((row, start_col)).value not in (None, ''):
            row += 1
        self.last_empty_row_cache[cache_key] = row
        return row

    def clear_last_empty_row_cache(self, sheet_name=None):
        if sheet_name:
            self.last_empty_row_cache = {k: v for k, v in self.last_empty_row_cache.items() if k[0] != sheet_name}
        else:
            self.last_empty_row_cache = {}

    def write_row(self, sheet_name, start_row, start_col, data):
        ws = self.get_sheet(sheet_name)
        row = self.get_last_empty_row(sheet_name, start_row, start_col, len(data))
        for i in range(len(data)):
            if ws.range((row, start_col + i)).value not in (None, ''):
                raise Exception(f'Ô tại dòng {row}, cột {start_col + i} đã có dữ liệu!')
        for i, value in enumerate(data):
            ws.range((row, start_col + i)).value = value
        self.clear_last_empty_row_cache(sheet_name)
        return row

    def undo_row(self, sheet_name, row, start_col, num_fields):
        ws = self.get_sheet(sheet_name)
        for i in range(num_fields):
            ws.range((row, start_col + i)).value = None
        self.clear_last_empty_row_cache(sheet_name)

    def preview_rows(self, sheet_name, start_row, start_col, num_fields, preview_range=2):
        ws = self.get_sheet(sheet_name)
        row = self.get_last_empty_row(sheet_name, start_row, start_col, num_fields)
        min_row = max(1, row - preview_range)
        max_row = row + preview_range
        cell_cache = {}
        for c in range(start_col, start_col + num_fields):
            for r in range(min_row, max_row + 1):
                cell_cache[(r, c)] = ws.range((r, c)).value
        return cell_cache, row, min_row, max_row

    def save(self):
        self.wb.save()

    def close(self):
        self.wb.close()
        self.app.quit()
