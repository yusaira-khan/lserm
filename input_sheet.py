import datetime
import os.path
import re
import openpyxl
import openpyxl.worksheet.worksheet

from output_record import OutputRecord
from lserm_row import LsermRow


class InputSheet:
    EFFECTIVE_DATE_ADDRESS = "B4"
    RELEASE_DATE_ADDRESS = "B3"
    SECTION_ELIGIBLE = "all"
    SECTION_ADD = "additions"
    SECTION_DEL = "deletions"
    SECTION_INTRO = "introduction"
    SECTION_EMPTY = "middle"

    def __init__(self, path=""):
        self._path = path
        self._sheet = None

        self._metadata = {}
        self.output = OutputRecord(self._metadata)

        self.rows_read = 0
        self._section = self.SECTION_INTRO

    def load(self):
        self._open()
        self._parse()
        self._load_metadata()
        self._close()
        return self

    def _open(self):
        if self._sheet is None:
            self._sheet = openpyxl.load_workbook(self._path, read_only=True).active

    def _close(self):
        if self._sheet is not None:
            self._sheet.parent.close()
            self._sheet = None

    def _parse(self):
        for r_ in self._sheet.rows:
            self.rows_read += 1
            row = LsermRow(r_)
            self._assign_section(row)
            if row.is_table_header() or row.is_empty() or row.is_partially_empty():
                continue
            else:
                self._record_row(row)

    def _assign_section(self, row: LsermRow):
        if self._section == self.SECTION_INTRO:
            if row.is_table_header():
                self._section = self.SECTION_ELIGIBLE
                return
        if self._section != self.SECTION_INTRO and self._section != self.SECTION_EMPTY:
            if row.is_empty():
                self._section = self.SECTION_EMPTY
                return
        if self._section == self.SECTION_EMPTY:
            if row.is_addition_start():
                self._section = self.SECTION_ADD
                return
            if row.is_deletion_start():
                self._section = self.SECTION_DEL
                return

    def _record_row(self, row: LsermRow):
        match self._section:
            case self.SECTION_ELIGIBLE:
                self.output.record_current(row)
            case self.SECTION_ADD:
                self.output.record_addition(row)
            case self.SECTION_DEL:
                self.output.record_deletion(row)

    def _load_metadata(self):
        f = os.path.basename(self._path).replace(".xlsx", "")
        self._metadata["file_name"] = f
        self._metadata["quarter_end_date"] = self.date_from_quarter_str(f)
        self._metadata["effective_date"] = self.get_date_from_value(
            self._sheet[self.EFFECTIVE_DATE_ADDRESS].value)
        self._metadata["release_date"] = self.get_date_from_value(
            self._sheet[self.RELEASE_DATE_ADDRESS].value)

    @staticmethod
    def get_date_from_value(value: str | datetime.datetime):
        if type(value) == str:
            return datetime.date.fromisoformat(value)
        else:
            return value.date()

    @classmethod
    def date_from_quarter_str(cls,f: str):
        s = re.match("Q(?P<quarter>\\d)-(?P<year>\\d+)", f)
        y = int(s["year"])
        q = int(s["quarter"])
        m, d = cls.QUARTERS[q]
        return datetime.date(y, m, d)

    QUARTERS = {
        1: (3, 31),
        2: (6, 30),
        3: (9, 30),
        4: (12, 31)
    }
