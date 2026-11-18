import unittest
from unittest.mock import MagicMock, patch
from PIL import Image, ImageDraw

import shibusawa as s


class ShibusawaRulesTests(unittest.TestCase):
    def test_all_detail_pages_are_read_in_order(self):
        for scanned in (False, True):
            for detail_count in (2, 3):
                with self.subTest(scanned=scanned, detail_count=detail_count):
                    texts = ["表紙"] + [
                        f"見 積 明 細 内 訳 表（{i}/{detail_count}）\n施工費（詳細別紙）"
                        for i in range(1, detail_count + 1)
                    ] + ["NO4\n【施工費】", "NO5\n前ページより"]
                    document = MagicMock()
                    document.page_count = len(texts)
                    document.__enter__.return_value = document
                    pages = [MagicMock() for _ in texts]
                    for page, text in zip(pages, texts):
                        page.get_text.return_value = "" if scanned else text
                    document.__getitem__.side_effect = pages.__getitem__
                    with patch.object(s.fitz, "open", return_value=document), patch.object(
                        s, "_ocr_page_text", side_effect=lambda doc, i: texts[i]
                    ), patch.object(
                        s, "_extract_coordinate_rows_for_page",
                        side_effect=lambda doc, i: [s._row(str(i), f"明細{i}", "2", "台", "100", "200")],
                    ) as extract:
                        result = s.extract_shibusawa_detail_rows("multiple.pdf")
                    self.assertEqual([call.args[1] for call in extract.call_args_list],
                                     list(range(1, detail_count + 1)))
                    self.assertEqual(result["項目"].tolist(),
                                     [f"明細{i}" for i in range(1, detail_count + 1)])

    def test_construction_pages_and_continuation_are_not_extracted(self):
        document = MagicMock()
        document.page_count = 5
        document.__enter__.return_value = document
        pages = [MagicMock() for _ in range(5)]
        for page, text in zip(pages, ["表紙", "見 積 明 細 内 訳 表\n施工費（詳細別紙）",
                                      "NO2\n【　施工費　】", "NO3\n233,550 前ページより",
                                      "見積明細内訳表（2/2）"]):
            page.get_text.return_value = text
        document.__getitem__.side_effect = pages.__getitem__
        with patch.object(s.fitz, "open", return_value=document), patch.object(
            s, "_extract_coordinate_rows_for_page", return_value=[s._row("1", "施工費（詳細別紙）")]
        ) as extract:
            result = s.extract_shibusawa_detail_rows("sample.pdf")
        self.assertEqual([call.args[1] for call in extract.call_args_list], [1, 4])
        self.assertEqual(result["項目"].tolist(), ["施工費（詳細別紙）"] * 2)

    def test_scanned_construction_page_is_skipped(self):
        document = MagicMock()
        document.page_count = 2
        document.__enter__.return_value = document
        document.__getitem__.return_value.get_text.return_value = ""
        with patch.object(s.fitz, "open", return_value=document), patch.object(
            s, "_ocr_page_text", return_value="NO2\n【 施工費 】"
        ), patch.object(s, "_extract_coordinate_rows_for_page") as extract:
            result = s.extract_shibusawa_detail_rows("scan.pdf")
        self.assertTrue(result.empty)
        extract.assert_not_called()

    def test_completion_messages_stay_out_of_csv(self):
        document = MagicMock()
        document.page_count = 2
        document.__enter__.return_value = document
        document.__getitem__.return_value.get_text.return_value = "見積明細内訳表"
        rows = [
            s._row("1", "ラベル", unit_price="15", amount="18,000", remark="原本の備考"),
            s._row("1", "施工費", "2", unit_price="100"),
            s._row("1", "器材費", "2", amount="600"),
            s._row("1", "時間外費", unit_price="100", amount="200"),
        ]
        with patch.object(s.fitz, "open", return_value=document), patch.object(
            s, "_extract_coordinate_rows_for_page", return_value=rows
        ):
            result = s.extract_shibusawa_detail_rows("sample.pdf")
        self.assertEqual(result.iloc[0]["数量"], "1,200")
        self.assertEqual(result.iloc[0]["備考"], "原本の備考")
        self.assertEqual(len(result.attrs["completion_messages"]), 3)
        self.assertIn("ラベル", result.attrs["completion_messages"][0])
        self.assertIn("金額÷単価", result.attrs["completion_messages"][0])
        self.assertNotIn("補完", result.to_csv(index=False))
        self.assertEqual(list(result.columns), s.COLUMNS)

    def test_overtime_without_quantity_is_removed_before_inference(self):
        rows = [
            s._row("1", "時間外費", unit_price="2,800"),
            s._row("1", "時間外費", unit_price="2,800", amount="5,600"),
            s._row("1", "深夜時間外費", unit_price="3,300"),
            s._row("1", "一般作業員", "1", unit_price="16,000", amount="16,000"),
            s._row("1", "時間外費", "2", unit_price="2,500", amount="5,000"),
        ]
        self.assertEqual(s._prepare_shibusawa_rows(rows), rows[3:])

    def test_uneven_rule_does_not_merge_overtime_and_worker(self):
        image = Image.new("RGB", (1000, 1000), "white")
        draw = ImageDraw.Draw(image)
        for y in range(200, 901, 100):
            if y == 400:
                draw.line((20, y - 1, 500, y - 1), fill="black")
                draw.line((500, y + 1, 980, y + 1), fill="black")
            else:
                draw.line((20, y, 980, y), fill="black", width=2)
        for x in (20, 280, 400, 520, 640, 980):
            draw.line((x, 200, x, 900), fill="black", width=2)
        horizontal, _ = s._detect_table_grid(image)
        self.assertTrue(any(abs(y - 400) <= 1 for y in horizontal))

    def test_general_worker_ocr_aliases_preserve_numbers(self):
        for item in ("一般", "一舩", "一舷", "ー般"):
            with self.subTest(item=item):
                row = s._parse_shibusawa_line(f"{item} 10 20,000 200,000", "1")
                self.assertEqual(row, s._row("1", "一般", "10", unit_price="20,000", amount="200,000"))
                result = s._fill_shibusawa_units(s._prepare_shibusawa_rows([row]))
                self.assertEqual(result, [s._row("1", "一般", "10", "名", "20,000", "200,000")])
        for item in ("一般作業員", "一般廃棄物", "一舩ビル"):
            self.assertEqual(s._canonical_item(item), item)

    def test_work_heading_does_not_receive_person_unit(self):
        rows = [s._row("1", "現場名", remark="搬出作業 (日中作業)")]
        self.assertEqual(s._fill_shibusawa_units(rows)[0]["単位"], "")

    def test_partial_remark_rule_does_not_split_item_row(self):
        image = Image.new("RGB", (1000, 1000), "white")
        draw = ImageDraw.Draw(image)
        for y in range(200, 901, 100):
            draw.line((20, y, 980, y), fill="black", width=2)
        for x in (20, 280, 400, 520, 640, 980):
            draw.line((x, 200, x, 900), fill="black", width=2)
        draw.line((640, 350, 980, 350), fill="black", width=2)
        horizontal, vertical = s._detect_table_grid(image)
        self.assertEqual(len(vertical), 6)
        self.assertEqual(len(horizontal), 8)
        self.assertFalse(any(340 < y < 360 for y in horizontal))

    def test_remark_subrows_are_read_in_the_same_cell(self):
        image = Image.new("RGB", (500, 100), "white")
        draw = ImageDraw.Draw(image)
        draw.line((0, 50, 499, 50), fill="black", width=2)
        draw.rectangle((30, 15, 50, 35), fill="black")
        draw.rectangle((30, 65, 50, 85), fill="black")
        with patch.object(s, "_ocr_image_crop", side_effect=["搬出作業", "(日中作業)"]) as ocr:
            result = s._read_table_cell(image, (0, 0, 500, 100))
        self.assertEqual(ocr.call_count, 2)
        self.assertEqual(result, "搬出作業 (日中作業)")

    def test_filename_does_not_change_extraction(self):
        expected = s._row("1", "新しい項目", "3", unit_price="123", amount="369")
        for filename in ("shibusawa_03.pdf", "shibusawa_05.pdf", "別案件.pdf"):
            with self.subTest(filename=filename):
                document = MagicMock()
                document.page_count = 3
                document.__enter__.return_value = document
                document.__getitem__.return_value.get_text.return_value = "見積明細内訳表"
                with patch.object(s.fitz, "open", return_value=document), patch.object(
                    s, "_extract_coordinate_rows_for_page", return_value=[expected]
                ) as extract:
                    result = s.extract_shibusawa_detail_rows(filename)
                self.assertEqual(extract.call_count, 2)
                self.assertEqual(result.to_dict("records"), [expected, expected])

    def test_no_fixed_item_date_or_reference_rows(self):
        self.assertEqual(s._canonical_item("", unit_price="220", amount="253,000"), "")
        self.assertEqual(s._parse_shibusawa_line("7/12(月) 作業", "1")["項目"], "7/12(月)")
        with patch.object(s, "_ocr_page_text", return_value="NO. 2\n新規品 3 123 369"):
            rows = s._extract_ocr_rows_for_page(None, 1)
        self.assertEqual([row["項目"] for row in rows], ["新規品"])

    def test_shared_corrections_use_actual_numbers(self):
        rows = [
            s._row("1", "ダンボール", unit_price="300", amount="90,000"),
            s._row("1", "エアキャップ", "0", unit_price="5,000", amount="15,000"),
            s._row("1", "諾経費", "71", unit_price="12,345", amount="24,690"),
            s._row("1", "一再作業員", "2", unit_price="17,000"),
            s._row("1", "人員輸送費", "2", amount="22,000"),
            s._row("1", "養生器材費", "1", amount="9,000"),
            s._row("1", "諸経費"),
            s._row("1", "時間外費", unit_price="3,000"),
        ]
        result = s._fill_shibusawa_units(s._prepare_shibusawa_rows(rows))
        self.assertEqual(len(result), 7)
        self.assertEqual([r["数量"] for r in result], ["300", "3", "2", "2", "2", "1", ""])
        self.assertEqual([r["単位"] for r in result], ["枚", "本", "式", "名", "式", "式", ""])
        self.assertEqual(result[2]["項目"], "諸経費")
        self.assertEqual(result[3]["金額"], "34,000")
        self.assertEqual(result[4]["単価"], "11,000")


if __name__ == "__main__":
    unittest.main()
