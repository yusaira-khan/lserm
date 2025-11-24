import openpyxl.cell.cell


class LsermRow:
    def __init__(self, row):
        self._input_row = row
        self.col_a = row[0].value
        self.col_b = row[1].value

    def matches_etf_prefix(self):
        for p in self.PREFIX_FOR_ETF:
            if self.col_a.lower().startswith(p):
                return True
        return False

    _PREFIX_FOR_ETF_ = [
        "BMO",
        "PIMCO",
        "iShr",
        "iShare"
        "CI"
        "TD",
        "Fidelity",
        "Vanguard",
        "Vangrd",
        "Harvest",
        "Ham",
        "GlobalX",
        "Global X",
        "GblX"
        "Glbl X"
        "GlblX"
        "Desjardn",
        "Desjrdn"
        "Des"
        "Desjardin"
        "Dynamc"
        "Dynamic"
        "Dyna "
        "Sprott",
        "CIBC"
        "RBC"
    ]
    PREFIX_FOR_ETF = [x.lower() for x in _PREFIX_FOR_ETF_]

    def matches_stock_prefix(self):
        for p in self.PREFIX_FOR_STOCK:
            if self.col_a.startswith(p):
                return True
        return False

    PREFIX_FOR_STOCK = ["CI Financial", "Sprott Inc", "Descartes"]

    def is_etf(self):
        return (not self.matches_stock_prefix()) and self.matches_etf_prefix()

    def is_addition_start(self):
        return self.col_a == "ADDITIONS"

    def is_deletion_start(self):
        return self.col_a == "DELETIONS"

    def is_table_header(self):
        return self.col_a == "Description" and self.col_b == "SYMBOL"

    def is_empty(self):
        return self.col_a is None or self.col_a == ""

    def is_partially_empty(self):
        return self.col_b is None or self.col_b == ""

    def line(self):
        return self._input_row[0].row

    def __repr__(self):
        return f"{self.line():04}, {self.col_a}, {self.col_b}"

    def write_excel(self, cell_a: openpyxl.cell.cell.Cell, cell_b: openpyxl.cell.cell.Cell):
        cell_a.value = self.col_a
        cell_b.value = self.col_b
