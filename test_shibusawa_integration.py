"""本体のPDF読込・転記ブロックとCOMへの列指定を検証する。"""
import ast
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock
import os
import re

import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter
import shibusawa_tokyo as s


TREE = ast.parse(Path(__file__).with_name('main.py').read_text(encoding='utf-8'))


def execute(nodes, ns):
    exec(compile(ast.Module(body=nodes, type_ignores=[]), 'main.py', 'exec'), ns)


class ShibusawaIntegrationTests(unittest.TestCase):
    def test_preview_cache_messages_and_k_column(self):
        rows = pd.DataFrame([
            s._row('1', '移転作業'),
            s._row('1', 'ラベル', '1,200', '枚', '15', '18,000', '原本備考'),
        ], columns=s.COLUMNS)
        rows.attrs['completion_messages'] = ['ラベルの数量を補完しました。']
        st = Mock()
        st.session_state = {}
        st.spinner.side_effect = lambda *a: nullcontext()
        extract = Mock(return_value=rows)
        ns = dict(os=os, re=re, pd=pd, st=st, company='澁澤（東京）PDF',
                  supplier_file=SimpleNamespace(name='a.pdf', getvalue=lambda: b'first'),
                  save_uploaded_bytes=Mock(return_value='a.pdf'), ITOKI_COMPANIES=(),
                  SHIBUSAWA_COLUMNS=s.COLUMNS, SHIBUSAWA_COL_MAP=s.TRANSFER_COL_MAP,
                  extract_shibusawa_detail_rows=extract,
                  build_shibusawa_transfer_df=s.build_transfer_df,
                  column_index_from_string=column_index_from_string,
                  get_column_letter=get_column_letter)
        for name in ('OFFICE_COL_MAP', 'COMANY_COL_MAP', 'TEISOH_COL_MAP', 'WEBERTON_COL_MAP',
                     'ITOKI_LIST_PRICE_COL_MAP', 'ITOKI_ESTIMATE_PRICE_COL_MAP'):
            ns[name] = {}
        execute([n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name in
                 {'get_col_map', 'get_row_from_cell', 'validate_dest_cell', 'write_df_to_template_com'}], ns)
        upload = [n for n in TREE.body if isinstance(n, ast.If) and
                  isinstance(n.test, ast.Name) and n.test.id == 'supplier_file']
        execute(upload, ns)
        execute(upload, ns)
        self.assertEqual(extract.call_count, 1)
        self.assertEqual(ns['df_converted'].loc[1, 'E'], '単価')
        st.info.assert_not_called()
        ns['supplier_file'] = SimpleNamespace(name='b.pdf', getvalue=lambda: b'second')
        execute(upload, ns)
        self.assertEqual(extract.call_count, 2)
        transfer = next(n for n in ast.walk(TREE) if isinstance(n, ast.Try) and n.body
                        and isinstance(n.body[0], ast.Assign) and
                        any(isinstance(t, ast.Name) and t.id == 'src_row' for t in n.body[0].targets))
        nodes = []
        for node in transfer.body:
            if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'output_name' for t in node.targets):
                break
            nodes.append(node)
        ns['src'] = 'B1'
        execute(nodes, ns)
        data = ns['df_to_write']
        self.assertEqual(data.iloc[1].to_dict(), {'B': 'ラベル', 'C': '原本備考', 'E': 1200, 'F': '枚', 'K': 15})
        self.assertEqual(data.iloc[0]['B'], '移転作業')
        self.assertNotIn('補完', data.to_csv(index=False))
        written = {}
        ws = Mock()
        ws.Cells.side_effect = lambda row, col: (row, col)
        def target(start, end):
            written[start[1]] = SimpleNamespace(Value=None)
            return written[start[1]]
        ws.Range.side_effect = target
        excel = Mock()
        excel.Workbooks.Open.return_value.Worksheets.return_value = ws
        ns.update(win32=SimpleNamespace(DispatchEx=lambda _: excel), pythoncom=Mock())
        ns['write_df_to_template_com']('same.xlsx', 'same.xlsx', '見積', 'B3', data,
                                      ns['get_col_map']('澁澤（東京）PDF'))
        self.assertEqual(set(written), {2, 3, 5, 6, 11})
        self.assertEqual(written[11].Value, [[''], [15]])
        ns['src'] = 'B3'
        execute(nodes, ns)
        self.assertEqual(len(ns['df_to_write']), 1)

    def test_invalid_number_is_not_written(self):
        with self.assertRaises(ValueError):
            s.build_transfer_df([s._row('1', 'ラベル', '###', '枚', '15')])


if __name__ == '__main__':
    unittest.main()
