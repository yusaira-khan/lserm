import datetime

from gsheets_adapter import GsheetsAdapter
from output_record import OutputRecord
from lserm_row import LsermRow

import openpyxl.worksheet.worksheet
import openpyxl.styles
import openpyxl.worksheet.dimensions


class OutputGsheet(OutputRecord):
    gsheet: GsheetsAdapter = None
    sheet_id = ""
    sheet_title = ""

    @classmethod
    def convert(cls, sup: OutputRecord, wb: openpyxl.Workbook):
        sup.__class__ = cls
        sup.gsheet = GsheetsAdapter()
        return sup

    def write(self):
        self.sheet_title = str(self._metadata["output_title"])
        self.gsheet.create_sheet(title=self.sheet_title)
        metadata_start = 1
        metadata_end = self.write_metadata(metadata_start)

        table_start = metadata_end + 2
        table_1_col = 1
        table_2_col = table_1_col + 4

        data_start = self.write_header(header_label="Eligible ETFs",
                                       start_row=table_start, start_col=table_1_col, with_addition=True)
        data_end = self.write_list(rows=self._etfs_eligible, with_matching=self._etfs_added,
                                   start_row=data_start, start_col=table_1_col)
        print(
            f"wrote table from ({table_start},{table_1_col}) to ({data_end},{table_1_col + 2}) in {self.sheet_title}")

        data_start = self.write_header("Deletions",
                                       start_row=table_start, start_col=table_2_col)
        data_end = self.write_list(rows=self._etfs_deleted,
                                   start_row=data_start, start_col=table_2_col)
        print(
            f"wrote table from ({table_start},{table_2_col}) to ({data_end},{table_2_col + 1}) in {self.sheet_title}")

    def write_metadata(self, start_row: int) -> int:
        data = []
        for key in ["file_name", "quarter_end_date", "release_date", "effective_date"]:
            title = key.replace("_", " ").title()
            data.append([title, str(self._metadata[key])])

        return self.upload_values(data, start_row, 1)

    def write_header(self, header_label: str, start_row: int,
                     start_col: int, with_addition:bool = False) -> int:
        headers = [[header_label], ["Description", "SYMBOL"]]
        if with_addition:
            headers[-1].append("Addition?")
        return self.upload_values(headers, start_row, start_col)

    def write_list(self, rows: list[LsermRow],
                   start_row: int, start_col: int,
                   with_matching: list[LsermRow] = ()) -> int:
        table = []
        for idx, data in enumerate(rows):
            table_row = [data.col_a, data.col_b]
            for w in with_matching:
                if data.col_b == w.col_b:
                    table_row.append("Addition")
            table.append(table_row)

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
