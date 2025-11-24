import unittest
import unittest.mock

from lserm_row import LsermRow

class LsermRowTest(unittest.TestCase):
    def test_is_empty(self):
        self.assertTrue(lserm_row(col_a="",col_b="").is_empty())
        self.assertFalse(lserm_row(col_a="data", col_b="").is_empty())

    def test_is_partially_empty(self):
        self.assertTrue(lserm_row(col_a="data",col_b="").is_partially_empty())
        self.assertFalse(lserm_row(col_a="data",col_b="data").is_partially_empty())

    def test_is_addition_start(self):
        self.assertTrue(lserm_row(col_a="ADDITIONS",col_b="").is_addition_start())
        self.assertFalse(lserm_row(col_a="blah",col_b="").is_addition_start())

    def test_is_deletion_start(self):
        self.assertTrue(lserm_row(col_a="DELETIONS",col_b="").is_deletion_start())
        self.assertFalse(lserm_row(col_a="blah",col_b="").is_deletion_start())

    def test_is_table_header(self):
        self.assertTrue(lserm_row(col_a="Description",col_b="SYMBOL").is_table_header())
        self.assertFalse(lserm_row(col_a="data",col_b="data").is_table_header())

    def test_is_etf(self):
        self.assertTrue(lserm_row(col_a="CI Global fund",col_b="---").is_etf())


def lserm_row(col_a, col_b, line_number=1):
    base_row = (mock_cell(col_a, line_number), mock_cell(col_b, line_number))
    return LsermRow(base_row)

def mock_cell(value, row_number):
    cell = unittest.mock.Mock()
    cell.row = row_number
    cell.value = value
    return cell