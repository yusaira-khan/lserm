import time

from gsheets_adapter import GsheetsAdapter, GsheetsColor
from output_record import OutputRecord
from lserm_row import LsermRow


class GsheetConverter:
    @staticmethod
    def write(records: list[OutputRecord]):
        adapter = GsheetsAdapter()
        for r in records:
            o = OutputGsheet.convert(r, adapter)
            o.write()


class OutputGsheet(OutputRecord):
    gsheet: GsheetsAdapter = None
    sheet_id = ""
    sheet_title = ""
    SYMBOL_WIDTH = 120
    DESCRIPTION_WIDTH = 360
    HIGHLIGHT_COLOR = "fff2cc"
    SLEEP_TIME = 15 #seconds

    @classmethod
    def convert(cls, sup: OutputRecord, gsheet):
        sup.__class__ = cls
        sup.gsheet = gsheet
        return sup

    def write(self):
        self.sheet_title = self.title()
        self.sheet_id = self.gsheet.find_or_create_sheet(title=self.sheet_title)
        metadata_start = 1
        metadata_end = self.write_metadata(metadata_start)

        table_start = metadata_end + 2
        table_1_col = 1
        table_2_col = table_1_col + 3

        data_start = self.write_header(header_label="Eligible ETFs (Additions Highlighted)",
                                       start_row=table_start, start_col=table_1_col)
        data_end = self.write_list(rows=self._etfs_eligible, with_matching=self._etfs_added,
                                   start_row=data_start, start_col=table_1_col, add_filter=True)
        print(
            f"wrote table from ({table_start},{table_1_col})"
            f" to ({data_end},{table_1_col + 1}) in {self.sheet_title}")

        data_start = self.write_header("Deletions",
                                       start_row=table_start, start_col=table_2_col)
        data_end = self.write_list(rows=self._etfs_deleted,
                                   start_row=data_start, start_col=table_2_col, cross_out=True)
        print(
            f"wrote table from ({table_start},{table_2_col})"
            f" to ({data_end},{table_2_col + 1}) in {self.sheet_title}")

        print(f"sleeping for {self.SLEEP_TIME} seconds")
        time.sleep(self.SLEEP_TIME)

    def write_metadata(self, start_row: int) -> int:
        data = []
        for key in ["file_name", "quarter_end_date", "release_date", "effective_date"]:
            title = key.replace("_", " ").title()
            data.append([title, str(self._metadata[key])])

        return self.upload_values(data, start_row, 1)

    def write_header(self, header_label: str, start_row: int,
                     start_col: int) -> int:
        headers = [[header_label], ["Description", "SYMBOL"]]
        bold_grid_range_list = [
            self.to_grid_range(start_row=start_row, start_col=start_col), #header
            self.to_grid_range(start_row=start_row+1, start_col=start_col), #description
            self.to_grid_range(start_row=start_row + 1, start_col=start_col+1) #symbol
        ]

        border_grid_range = self.to_grid_range(start_row + 1, start_col, num_cols=2)
        self.gsheet.make_bold(bold_grid_range_list)
        self.gsheet.add_border(border_grid_range)
        self.gsheet.change_column_widths(self.sheet_id, [start_col, start_col + 1],
                                         [self.DESCRIPTION_WIDTH, self.SYMBOL_WIDTH])
        return self.upload_values(headers, start_row, start_col)

    def write_list(self, rows: list[LsermRow],
                   start_row: int, start_col: int,
                   with_matching: list[LsermRow] = (),
                   cross_out: bool = False, add_filter: bool = False) -> int:
        table = []
        highlighted_idx = []
        for idx, data in enumerate(rows):
            table_row = [data.col_a, data.col_b]
            for w in with_matching:
                if data == w:
                    highlighted_idx.append(idx)
            table.append(table_row)
        if highlighted_idx:
            grid_range_list = self.table_indexes_to_grid_range_list(start_row, start_col, highlighted_idx)
            color = GsheetsColor(hex=self.HIGHLIGHT_COLOR)
            self.gsheet.set_background(grid_range_list, color)
        if cross_out:
            grid_range_list = self.table_indexes_to_grid_range_list(start_row, start_col, range(len(table)))
            self.gsheet.cross_out(grid_range_list)
        if add_filter:
            grid_range = self.to_grid_range(start_row-1, start_col, num_rows=len(table)+1, num_cols=2)
            self.gsheet.add_filter(grid_range)

        return self.upload_values(table, start_row, start_col)

    def upload_values(self, data: list[list[str]], start_row: int, start_col: int):
        range_address = self.address(start_row=start_row, start_col=start_col, dimensions=data)
        self.gsheet.set_values(data, range_address)
        return start_row + len(data)

    def address(self, start_row: int, start_col: int, dimensions: list[list[str]]) -> str:
        num_cols = max([len(d) for d in dimensions])
        num_rows = len(dimensions)
        end_row = start_row + num_rows - 1
        end_col = start_col + num_cols - 1
        return self.gsheet.range_address(
            sheet_title=self.sheet_title,
            start_row=start_row, end_row=end_row,
            start_col=start_col, end_col=end_col)

    def to_grid_range(self, start_row: int, start_col: int, num_rows=1, num_cols=1):
        start_row_idx = start_row - 1
        start_col_idx = start_col - 1
        end_row_idx = start_row_idx + num_rows
        end_col_idx = start_col_idx + num_cols
        return {
            "startRowIndex": start_row_idx,
            "startColumnIndex": start_col_idx,
            "endRowIndex": end_row_idx,
            "endColumnIndex": end_col_idx,
            "sheetId": self.sheet_id
        }

    def table_indexes_to_grid_range_list(self, start_row, start_col, indexes):
        grid_range_list = []
        for idx in indexes:
            row = start_row+idx
            grid_range_list.append(self.to_grid_range(start_row=row, start_col=start_col))
            grid_range_list.append( self.to_grid_range(start_row=row, start_col=start_col+1))
        return grid_range_list
