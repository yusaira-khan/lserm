class LsermRow:
    def __init__(self, row):
        self._input_row = row
        self.col_a = row[0].value
        self.col_a_lowercase = self.col_a.lower() if type(self.col_a) == str else self.col_a
        self.col_b = row[1].value

    def is_etf(self):
        return (
                (self.has_etf_infix() and not self.has_incorrect_infix())
                or
                (self.has_etf_prefix() and not self.has_incorrect_prefix())
        )

    def has_etf_infix(self):
        for i in self.INFIX_FOR_ETF:
            if i in self.col_a_lowercase:
                return True
        return False
    _INFIX_FOR_ETF_ = ["Asset", "Bond", "ETF", "Index"]
    INFIX_FOR_ETF = [x.lower() for x in _INFIX_FOR_ETF_]

    def has_incorrect_infix(self):
        return self.col_b == "NFLX"

    def has_etf_prefix(self):
        for p in self.PREFIX_FOR_ETF:
            if self.col_a_lowercase.startswith(p):
                return True
        return False
    _PREFIX_FOR_ETF_ = [
        "BMO",
        "PIMCO",
        "iShr", "iShare", "RBC",
        "CI", #matches CI Funds and CIBC
        "TD",
        "Fidelity", "Fid "
        "Vanguard", "Vangrd",
        "Harv", #Harvest
        "Ham", #Hamilton
        "GlobalX", "Global X", "GblX", "Glbl X", "GlblX", "BetaPro"
        "Desjardin", "Desjardn", "Desjrdn", "Des",
        "Dynamic", "Dynamc", "Dyna ",
        "Sprott", #Not an etf, but similar to Royal Mint's ETRs
        "Mack", "Wealthsimp",
        "NBI"
        "IA Clarington",
        "Purpose",
        "Ninepoint",
        "Invesco",
        "Evolve", "Ev ", "High Interest Savings Account Fund",
        "Scotia",
        "Manulife",
        "Brompton", "Bromptn"
        "Franklin",
        "AGF",
    ]
    PREFIX_FOR_ETF = [x.lower() for x in _PREFIX_FOR_ETF_]

    def has_incorrect_prefix(self):
        for p in self.PREFIX_FOR_FALSE_POSITIVES:
            if self.col_a.startswith(p):
                return True
        return False
    PREFIX_FOR_FALSE_POSITIVES = [
        "CI Financial", "Cineplex", "Cipher", #matches CI Funds
        "Hammond", #matches Ham for Hamilton
        "Descartes",  #matches Des for Desjardins
        "Sprott Inc",
        "Manulife Fin",
        "AGF Management",
    ]

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

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        if self.is_empty() or other.is_empty():
            return False
        if self.is_partially_empty() and other.is_partially_empty():
            return self.col_a == other.col_a
        return self.col_b == other.col_b


