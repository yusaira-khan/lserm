import unittest

from gsheets_adapter import GsheetsAdapter


class GsheetsAdapterTest(unittest.TestCase):
    COLUMN_LETTERS = [
        ("A", 1),
        ("B", 2),
        ("Z", 26),
        ("AA", 27),
        ("AB", 28),
        ("AZ", 52),
        ("BA", 53),
        ("BB", 54),
        ("ZZ", 702),
        ("AAA", 703),
        ("WTF", 16074),
        ("ZZZ", 18278),
    ]

    COLORS = [
        ("000000", {"red": 0.0, "green": 0.0, "blue": 0.0}),
        ("ff0000", {"red": 1.0, "green": 0.0, "blue": 0.0}),
        ("00ff00", {"red": 0.0, "green": 1.0, "blue": 0.0}),
        ("0000ff", {"red": 0.0, "green": 0.0, "blue": 1.0}),
        ("ffffff", {"red": 1.0, "green": 1.0, "blue": 1.0}),
    ]

    def test_correct_column_number(self):
        for letter, number in self.COLUMN_LETTERS:
            with self.subTest(letter=letter, number=number):
                self.assertEqual(GsheetsAdapter.column_number(letter), number)

    def test_correct_column_letter(self):
        for letter, number in self.COLUMN_LETTERS:
            with self.subTest(number=number, letter=letter, ):
                self.assertEqual(GsheetsAdapter.column_letter(number), letter)

    def test_column_letter_raises_error_for_too_small(self):
        with self.assertRaises(ValueError):
            GsheetsAdapter.column_letter(0)

    def test_column_letter_raises_error_for_too_big(self):
        with self.assertRaises(ValueError):
            GsheetsAdapter.column_letter(20_000)

    def test_column_number_raises_error_for_too_small(self):
        with self.assertRaises(ValueError):
            GsheetsAdapter.column_number("")

    def test_column_number_raises_error_for_too_big(self):
        with self.assertRaises(ValueError):
            GsheetsAdapter.column_number("AAAA")

    def test_correct_color_hex_to_dict(self):
        for hex_color, dict_color in self.COLORS:
            with self.subTest(hex_color=hex_color, dict_color=dict_color):
                self.assertEqual(GsheetsAdapter.color_hex_to_dict(hex_color), dict_color)

    def test_correct_color_dict_to_hex(self):
        for hex_color, dict_color in self.COLORS:
            with self.subTest(dict_color=dict_color, hex_color=hex_color):
                self.assertEqual(GsheetsAdapter.color_dict_to_hex(dict_color), hex_color)


