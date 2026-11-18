"""東京の横長明細を、CSV作成時と同じPDF文字座標から日程別に抽出する。"""
import re
import unicodedata
from decimal import Decimal

import fitz
import pandas as pd
from shibusawa import COLUMNS, TRANSFER_COL_MAP, build_transfer_df as _build_transfer_df, _row


def build_transfer_df(records):
    """諸経費後の空白行を転記先にも残す。"""
    result = []
    for row in records:
        if not any(str(row.get(key, '')).strip() for key in COLUMNS):
            result.append(dict.fromkeys(TRANSFER_COL_MAP, ''))
        else:
            result.extend(_build_transfer_df([row]).to_dict('records'))
    if not result or not any(row['B'] for row in result):
        raise ValueError('転記対象の行がありません。')
    return pd.DataFrame(result, columns=list(TRANSFER_COL_MAP), dtype=object)


def compact(value):
    return re.sub(r"\s+", "", value or "")


def number(value):
    text = unicodedata.normalize("NFKC", compact(value)).replace(",", "")
    if not text:
        return None
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text):
        raise ValueError(f"PDFの数値を読み取れませんでした: {value}")
    return Decimal(text)


def formatted(value):
    if value is None:
        return ""
    return format(value, "f").rstrip("0").rstrip(".") if value % 1 else str(int(value))


def extract_page(page, page_number):
    tables = page.find_tables().tables
    selected = []
    for table in tables:
        data = table.extract()
        for hi, row in enumerate(data):
            if len(row) >= 7 and compact(row[0]) == "項目" and compact(row[2]) == "単価":
                selected.append((table, data, hi))
                break
    if len(selected) != 1:
        raise ValueError(f"PDF {page_number}ページ目の見積明細内訳表を判別できませんでした。")
    table, data, hi = selected[0]
    bounds = table.rows[hi].cells
    # 単語が列境界をまたぐPDFもあるため、同じ座標抽出を文字単位で行う。
    words = [(*char['bbox'], char['c'])
             for block in page.get_text('rawdict')['blocks'] if 'lines' in block
             for line in block['lines'] for span in line['spans'] for char in span['chars']]

    def region(x0, y0, x1, y1):
        return "".join(w[4] for w in sorted(words, key=lambda w: w[0])
                        if x0 < (w[0] + w[2]) / 2 < x1 and y0 < (w[1] + w[3]) / 2 < y1)

    def at(col, y0, y1):
        x0, _, x1, _ = bounds[col]
        return region(x0, y0, x1, y1)

    stages = [(c, data[hi - 1][c]) for c in range(3, len(bounds) - 2, 2)
              if compact(data[hi][c]) == "数量" and data[hi - 1][c]]
    if not stages:
        raise ValueError(f"PDF {page_number}ページ目の日程欄を読み取れませんでした。")
    units = {'人員輸送車': '車', '資材配達・回収費': '式', '養生費': '式',
             'ワンタッチケース': '箱', 'ラベル': '枚', 'ガムテープ': '個',
             'エアーキャップ': '巻', 'OAケース': '箱', '平台車': '台', '移動用資機材一式': '式'}
    items = []
    context, category = '', '人件費'
    for i in range(hi + 1, len(data)):
        cell = table.rows[i].cells[1]
        if not cell:
            continue
        x0, y0, x1, y1 = cell
        item = compact(region(x0, y0, min(x1, bounds[2][0]), y1))
        if not item or item.startswith('小計'):
            continue
        price_text = at(2, y0, y1)
        if compact(price_text).startswith('('):
            continue
        price = number(price_text)
        name, unit = item, units.get(item, '')
        if item in ('班長', '一般作業員'):
            context = '作業員 ' + item
            name, unit = context, '名'
        elif item in ('4㌧車', '２㌧車', '2㌧車'):
            category = '車両費'
            context = '車両費　' + ('4ｔ' if item.startswith('4') else '2ｔ')
            name, unit = context, '車'
        elif item in ('時間外', '夜間'):
            name = context + ' ' + item
            unit = '時間' if item == '時間外' else ('名' if category == '人件費' else '車')
        elif item.startswith('繁忙期割増'):
            name, unit = category + ' ' + item, '名' if category == '人件費' else '車'
        elif item.startswith('施工費'):
            unit = '式'
        elif item == '間仕切（kg）':
            name, unit = '廃棄処理代行費 間仕切', 'kg'
        values = []
        for col, _ in stages:
            amount = number(at(col + 1, y0, y1))
            qty_text = at(col, y0, y1)
            if amount is None:
                values.append(None)
                continue
            quantity = amount / price if '#' in qty_text and price else number(qty_text)
            if quantity is None or price is None or quantity * price != amount:
                raise ValueError(f"PDF {page_number}ページ目「{name}」の数量×単価と金額が一致しません。")
            values.append((quantity, amount))
        items.append((name, unit, price, values))

    expense_row = next((r for r in data if compact(r[0]) == '諸経費'), None)
    total_row = next((r for r in data if compact(r[0]) == '合計'), None)
    if expense_row is None or total_row is None:
        raise ValueError(f"PDF {page_number}ページ目の諸経費・合計欄を確認できませんでした。")
    output, totals = [], []
    previous_group = None
    group = ''
    for j, (col, label) in enumerate(stages):
        if hi >= 2 and data[hi - 2][col]:
            group = data[hi - 2][col]
        if group and group != previous_group:
            output.append(_row(str(page_number - 1), '【' + group.replace('\n', ' ') + '】'))
            previous_group = group
        output.append(_row(str(page_number - 1), '【' + label.replace('\n', ' ') + '】'))
        total = Decimal(0)
        for name, unit, price, values in items:
            if values[j] is None:
                continue
            qty, amount = values[j]
            output.append(_row(str(page_number - 1), name, formatted(qty), unit, formatted(price), formatted(amount)))
            total += amount
        expense = number(expense_row[col + 1])
        if expense is not None:
            output.append(_row(str(page_number - 1), '諸経費', '1', '式', formatted(expense), formatted(expense)))
            output.append(_row('', ''))
            total += expense
        if total != number(total_row[col + 1]):
            raise ValueError(f"PDF {page_number}ページ目「{label}」の明細合計が原本と一致しません。")
        totals.append(total)
    if sum(totals) != number(total_row[-1]):
        raise ValueError(f"PDF {page_number}ページ目の総合計が原本と一致しません。")
    return output


def extract_shibusawa_detail_rows(pdf_path):
    rows = []
    with fitz.open(pdf_path) as doc:
        for i in range(1, len(doc)):
            text = doc[i].get_text()
            if not text.strip():
                raise ValueError(f"PDF {i + 1}ページ目は文字データがありません。CSV作成時と同じ文字を選択できるPDFを使用してください。")
            if '見積明細内訳表' in compact(text):
                rows.extend(extract_page(doc[i], i + 1))
    if not rows:
        raise ValueError('見積明細内訳表が見つかりませんでした。')
    return pd.DataFrame(rows, columns=COLUMNS)
