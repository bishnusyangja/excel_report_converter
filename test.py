import unittest

from main import read_data_from_sheet, is_last_row, format_row_item


class TestMain(unittest.TestCase):

    def test_is_last_row_with_invalid(self):
        site_id = "site abc"
        result = is_last_row(site_id)
        self.assertEqual(result, None)

    def test_is_last_row_with_valid(self):
        site_id = "site 14"
        result = is_last_row(site_id)
        self.assertEqual(result, site_id)

    def test_format_row_item(self):
        row_item = [['Page Views', 'Page Views', 'Total Time Spent', 'Total Time spent', 'Visitors'],
                    ['2022-01-01', '2022-01-02', '2022-01-01', '2022-01-02', '2022-01-04'],
                    [4, 8, 9, 12, 14]
                    ]
        row_id = 'site 100'
        result = format_row_item(row_item, row_id,)
        expected = {'2022-01-01': [['Total Time Spent', 9]], '2022-01-02': [['Page Views', 8],
                            ['Total Time spent', 12]], '2022-01-04': [['Visitors', 14]]}
        self.assertEqual(result, expected)

    def test_read_data_from_sheet(self):
        sheet = ''
        read_data_from_sheet(sheet)
        self.assertEqual(2, 3)


if __name__ == "__main__":
    unittest.main()