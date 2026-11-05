"""表の先頭に空行があるウェバートン帳票の回帰テスト。"""
from pathlib import Path
import unittest
from unittest.mock import Mock, patch

import weberton


class WebertonBlankHeaderTest(unittest.TestCase):
    def extract_table(self, rows):
        page = Mock()
        page.find_tables.return_value.tables = [Mock(extract=lambda: rows)]
        document = Mock()
        document.__len__ = Mock(return_value=2)
        document.__getitem__ = Mock(return_value=page)
        with patch.object(weberton.fitz, 'open') as open_pdf:
            open_pdf.return_value.__enter__.return_value = document
            return weberton.extract_weberton_pdf('unused.pdf')

    def test_header_with_and_without_leading_blank_rows(self):
        header = ['', '品名', 'メーカー', '型番', '数量', '単位', '単価', '金額', '定価']
        detail = ['', 'LAN工事費', '', '', '1', '式', '90,000', '90,000', '']
        for prefix in ([], [[None] * 9], [['　'] * 9, [None] * 9]):
            with self.subTest(prefix=prefix):
                rows = self.extract_table(prefix + [header, detail])
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]['品名'], 'LAN工事費')
                self.assertEqual(rows[0]['金額'], '90000')

    def test_empty_and_unrecognized_tables_are_rejected(self):
        for rows in ([], [[None] * 9], [['別形式'] * 9]):
            with self.subTest(rows=rows):
                with self.assertRaisesRegex(ValueError, '対応する表がありません'):
                    self.extract_table(rows)

    def test_reported_pdf_amounts_and_transfer(self):
        source = Path(__file__).with_name('ウェバートン.pdf')
        if not source.exists():
            self.skipTest('実帳票がありません')
        rows = weberton.extract_weberton_pdf(source)
        self.assertEqual(weberton.validate_amounts(rows), 21)
        total = next(r for r in rows if weberton.compact(r['品名']) == '合計')
        self.assertEqual(total['金額'], '1200000')
        transfer = weberton.build_transfer_df(weberton.prepare_csv_rows(rows))
        self.assertEqual(len(transfer), 25)
        self.assertEqual(list(transfer.columns), ['B', 'C', 'E', 'F', 'I', 'K'])
        self.assertFalse(any(weberton.is_summary(weberton.compact(n)) for n in transfer['B']))
        note = next(r for r in rows if r['金額'] == 'ご支給品')
        self.assertEqual(weberton.build_transfer_df([note]).iloc[0]['K'], '')


if __name__ == '__main__':
    unittest.main()
