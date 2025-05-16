from . import testing
from xlcalculator.utils import resolve_table_ranges
from xlcalculator.xltypes import XLTable

# Dummy column class for our ExcelTableColumn
class DummyColumn:
    def __init__(self, name):
        self.name = name

class TableRangeTest(testing.XlCalculatorTestCase):
    def setUp(self):
        simple_table = XLTable(
            name="MyTable",
            sheet="Sheet1",
            cell_range="A1:C5",
            columns=[DummyColumn("Col1"), DummyColumn("Col2"), DummyColumn("Col3")],
            header_row_count=1
        )
        edge_case_table = XLTable(
            name="EdgeCaseTable",
            sheet="Sheet1",
            cell_range="A1:C5",
            columns=[DummyColumn("!Sales"), DummyColumn("[Brackets]"), DummyColumn("colon : bracket ][")],
            header_row_count=1,
        )
        totals_table = XLTable(
            name="TotalsTable",
            sheet="Sheet1",
            cell_range="A1:C5",
            columns=[DummyColumn("Col1"), DummyColumn("Col2"), DummyColumn("Col3")],
            header_row_count=2,
            has_totals_row=True
        )
        self.tables = {"MyTable": simple_table, "EdgeCaseTable": edge_case_table, "TotalsTable": totals_table}

    # Column ranges
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

    # Item specifiers
    def test_one_item_specifier(self):
        result = resolve_table_ranges("MyTable[#Headers]", self.tables)
        self.assertEqual(result, "Sheet1!A1:C1")

    def test_all_item_specifier(self):
        result = resolve_table_ranges("MyTable[#All]", self.tables)
        self.assertEqual(result, "Sheet1!A1:C5")

    def test_totals_item_specifier(self):
        result = resolve_table_ranges("TotalsTable[#Totals]", self.tables)
        self.assertEqual(result, "Sheet1!A5:C5")

    def test_data_item_specifier(self):
        result = resolve_table_ranges("TotalsTable[#Data]", self.tables, "Sheet1!A5")
        self.assertEqual(result, "Sheet1!A3:C4")

    def test_two_item_specifiers(self):
        result = resolve_table_ranges("MyTable[[#Headers],[#Data]]", self.tables)
        self.assertEqual(result, "Sheet1!A1:C5")

    def test_two_item_specifiers_with_totals(self):
        result = resolve_table_ranges("TotalsTable[[#Data],[#Totals]]", self.tables)
        self.assertEqual(result, "Sheet1!A3:C5")

    def test_this_row_item_specifier(self):
        result = resolve_table_ranges("MyTable[[#This Row],[Col1]]", self.tables, "Sheet!E2")
        self.assertEqual(result, "Sheet1!A2:A2")

    def test_item_specifiers_with_column_range(self):
        result = resolve_table_ranges("MyTable[[#Headers],[Col1]:[Col3]]", self.tables)
        self.assertEqual(result, "Sheet1!A1:C1")

    # Edge cases
    # Big challenge is to parse the items correctly when it has special characters such as spaces, brackets, colons, etc.
    def test_with_spaces(self):
        result = resolve_table_ranges("MyTable[ [Col2] ]", self.tables)
        self.assertEqual(result, "Sheet1!B2:B5")

    def test_column_name_with_special_characters(self):
        result = resolve_table_ranges("EdgeCaseTable['[Brackets']]", self.tables)
        self.assertEqual(result, "Sheet1!B2:B5")

        result = resolve_table_ranges("EdgeCaseTable[colon : bracket ']'[]", self.tables)
        self.assertEqual(result, "Sheet1!C2:C5")

    def test_column_range_with_special_characters(self):
        result = resolve_table_ranges("EdgeCaseTable[['[Brackets']]:[colon : bracket ']'[]]", self.tables)
        self.assertEqual(result, "Sheet1!B2:C5")

    # Error cases
    def test_not_a_table_range(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("SheetRefWithBracket[]!A2", self.tables)
        
        # External table reference
        with self.assertRaises(Exception):
            resolve_table_ranges("[external.xlsx]Sheet1!A1", self.tables)

    def test_table_not_found(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("NonExistentTable[Col1]", self.tables)

    def test_column_not_found(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("EdgeCaseTable[Col4]", self.tables)

    def test_invalid_item_specifier(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("TotalsTable[#Total]", self.tables)

    def test_invalid_item_specifier_combination(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("TotalsTable[[#Headers],[#Totals]]", self.tables)

    def test_invalid_no_totals_row(self):
        with self.assertRaises(Exception):
            resolve_table_ranges("MyTable[#Totals]", self.tables)

