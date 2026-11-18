import csv
import os
from pathlib import Path
import tempfile
import unittest
from decimal import Decimal

import fitz
import shibusawa_tokyo as s


ROOT = Path(__file__).resolve().parent
PDF_DIR = Path(os.environ.get('SHIBUSAWA_TOKYO_PDF_DIR',
                             str(Path.home() / 'Desktop' / '見積フォルダ' / '澁澤')))


@unittest.skipUnless((PDF_DIR / 'shibusawa_01.pdf').exists(), '東京の実PDFが必要です')
class TokyoPdfTests(unittest.TestCase):
    def test_matches_previous_csv_values_in_all_five_pdfs(self):
        for n in range(1, 6):
            with self.subTest(pdf=n):
                stem = f'shibusawa_{n:02}'
                suffix = '_移転作業_数量補完.csv' if n == 1 else '_移転作業.csv'
                with (ROOT / 'output' / 'csv' / stem / (stem + suffix)).open(encoding='utf-8-sig', newline='') as file:
                    csv_rows = list(csv.reader(file))[1:]
                expected = [tuple(Decimal(row[i].replace(',', '')) for i in (4, 6, 7))
                            for row in csv_rows if row[4] and row[6] and row[7]]
                df = s.extract_shibusawa_detail_rows(str(PDF_DIR / (stem + '.pdf')))
                actual = [tuple(Decimal(row[key]) for key in ('数量', '単価', '金額'))
                          for row in df.to_dict('records') if row['数量']]
                self.assertEqual(actual, expected)
                self.assertFalse(df['備考'].any())
                for index in df.index[df['項目'] == '諸経費']:
                    self.assertTrue(df.iloc[index + 1].eq('').all())
                transfer = s.build_transfer_df(df.to_dict('records'))
                for index in transfer.index[transfer['B'] == '諸経費']:
                    self.assertTrue(transfer.iloc[index + 1].eq('').all())
                self.assertEqual([v for v in transfer['K'] if v != ''], [int(v[1]) for v in expected])

    def test_three_detail_pages_with_construction_between_them(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'multi.pdf'
            with fitz.open(PDF_DIR / 'shibusawa_01.pdf') as source, fitz.open() as doc:
                doc.insert_pdf(source, from_page=0, to_page=1)
                doc.insert_pdf(source, from_page=2, to_page=2)
                doc.insert_pdf(source, from_page=1, to_page=1)
                doc.insert_pdf(source, from_page=1, to_page=1)
                doc.save(output)
            single = s.extract_shibusawa_detail_rows(str(PDF_DIR / 'shibusawa_01.pdf'))
            result = s.extract_shibusawa_detail_rows(str(output))
            self.assertEqual(len(result), len(single) * 3)
            self.assertEqual(result.loc[result['ページ'] != '', 'ページ'].unique().tolist(), ['1', '3', '4'])
            self.assertEqual(result['項目'].tolist(), single['項目'].tolist() * 3)

    def test_image_pdf_does_not_fall_back_to_wrong_ocr(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'image.pdf'
            with fitz.open() as doc:
                doc.new_page()
                doc.new_page()
                doc.save(output)
            with self.assertRaisesRegex(ValueError, '文字データがありません'):
                s.extract_shibusawa_detail_rows(str(output))


if __name__ == '__main__':
    unittest.main()
