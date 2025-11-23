# These tests are auto-generated with test data from:
# https://github.com/exercism/problem-specifications/tree/main/exercises/rna-transcription/canonical-data.json
# File last updated on 2023-07-19

import unittest
from unittest.mock import Mock

from input_sheet import InputSheet


def create_mock_row(is_table_header_return=False, is_empty_return=False,
                    is_addition_start_return=False,
                    is_deletion_start_return=False):
    m = Mock()
    attrs = {"is_addition_start.return_value": is_addition_start_return,
             "is_deletion_start.return_value": is_deletion_start_return,
             "is_table_header.return_value": is_table_header_return,
             "is_empty.return_value": is_empty_return
             }
    m.configure_mock(**attrs)
    return m


class InputSheetAssignSectionTest(unittest.TestCase):
    def test_default_section(self):
        subject = InputSheet()
        self.assertEqual(subject._section, InputSheet.SECTION_INTRO)

    def test_change_intro_to_current_with_tableheader(self):
        subject = InputSheet()
        subject._assign_section(create_mock_row(is_table_header_return=True))
        self.assertEqual(subject._section, InputSheet.SECTION_ELIGIBLE)

    def test_change_to_middle_with_empty(self):
        subject = InputSheet()
        cases = [InputSheet.SECTION_ELIGIBLE, InputSheet.SECTION_ADD, InputSheet.SECTION_DEL]
        for test_section in cases:
            with self.subTest(test_section=test_section):
                subject._section = test_section
                subject._assign_section(create_mock_row(is_empty_return=True))
                self.assertEqual(subject._section, InputSheet.SECTION_EMPTY)

    def test_not_change_to_middle_with_empty_when_intro(self):
        subject = InputSheet()
        subject._assign_section(create_mock_row(is_empty_return=True))
        self.assertEqual(subject._section, InputSheet.SECTION_INTRO)

    def test_change_empty_to_addition(self):
        subject = InputSheet()
        subject._section = InputSheet.SECTION_EMPTY
        subject._assign_section(create_mock_row(is_addition_start_return=True))
        self.assertEqual(subject._section, InputSheet.SECTION_ADD)
