import math
import unittest

import xlrd

from automation_app.report_parser import _optional_decimal_or_error


class OptionalBillDiscountTests(unittest.TestCase):
    def test_nan_in_optional_discount_is_treated_as_blank(self):
        cell = xlrd.sheet.Cell(xlrd.XL_CELL_NUMBER, math.nan)
        self.assertEqual(
            _optional_decimal_or_error(cell, field_name="Chiết khấu tổng bill (%)"),
            (None, None),
        )


if __name__ == "__main__":
    unittest.main()
