from . import testing
from xlcalculator.utils import resolve_table_ranges
from xlcalculator.xltypes import XLTable

# Dummy column class for XLTable columns
class DummyColumn:
    def __init__(self, name):
        self.name = name

class TableRangeTest(testing.XlCalculatorTestCase):
    def setUp(self):
        # Create a dummy table with columns and a range
        self.table = XLTable(
            name="MyTable",
            sheet="Sheet1",
            cell_range="A1:C5",
            columns=[DummyColumn("Col1"), DummyColumn("Col2"), DummyColumn("Col3")],
            header_row_count=1
        )
        self.tables = {"MyTable": self.table}

    def test_one_column_range(self):
        # Test resolving a single column reference
        result = resolve_table_ranges("MyTable[Col2]", self.tables)
        self.assertEqual(result, "Sheet1!B2:B5")

    def test_empty_bracket_range(self):
        result = resolve_table_ranges("MyTable[]", self.tables)
        self.assertEqual(result, "Sheet1!A2:C5")

    def test_two_column_range(self):
        result = resolve_table_ranges("MyTable[[Col2]:[Col3]]", self.tables)
        self.assertEqual(result, "Sheet1!B2:C5")
