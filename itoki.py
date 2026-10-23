"""イトーキ見積CSVの読み込みと列の並べ替え。

使用例: python itoki.py イトーキ_pdf_01.csv
    python itoki.py イトーキ_pdf_01.csv -o output/イトーキ_並べ替え.csv
元の値と行順を保持し、指定の11列をUTF-8 BOM付きCSVに出力する。
"""
from __future__ import annotations

import argparse
import csv
import io
from pathlib import Path
from decimal import Decimal, InvalidOperation
import unicodedata


REQUIRED_COLUMNS = (
    '商品記号', '商品名', '数量', '見積単価', '見積金額', '仕様寸法', '単位', '仕入価',
)
OUTPUT_COLUMNS = (
    '記載頁', '商品名', '商品記号', '仕様寸法', '数量',
    '見積単価', '見積金額', '定価', '備考', '仕入価', '単位',
)
TRANSFER_COL_MAP = {column: column for column in ('A', 'B', 'C', 'D', 'E', 'F', 'K', 'L')}
# 別形式のイトーキCSVで使われている列名。
COLUMN_ALIASES = {'記載頁': 'カタログ頁', '定価': '本体価格'}


def prepare_csv_rows(records):
    """指定の11列に並べ替える。"""
    result = []
    for row in records:
        prepared = {
            name: row[name] if name in row else row[COLUMN_ALIASES[name]]
            for name in OUTPUT_COLUMNS
        }
        dimension = unicodedata.normalize('NFKC', prepared['仕様寸法']).strip()
        if dimension.upper() == '0W*0D*0H':
            prepared['仕様寸法'] = ''
        result.append(prepared)
    return result


def _number(value, label):
    text = str(value).strip().replace(',', '') if value is not None else ''
    if not text:
        return None
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f'{label}を数値として読み取れません: {value}') from exc
    if not number.is_finite():
        raise ValueError(f'{label}に不正な数値があります: {value}')
    return number


def build_transfer_df(records):
    """テンプレートのA～F・K・L列用データ。割り切れない単価は丸めない。

    数量が空欄またはゼロならL列は空欄にする。見出しや行順は保持する。
    """
    import pandas as pd

    result = []
    for row_number, row in enumerate(records, start=1):
        qty = _number(row.get('数量'), f'{row_number}行目の数量')
        cost = _number(row.get('仕入価'), f'{row_number}行目の仕入価')
        price = _number(row.get('定価', row.get('本体価格')), f'{row_number}行目の定価')
        unit_cost = cost / qty if cost is not None and qty not in (None, 0) else None

        def excel_number(value):
            return '' if value is None else int(value) if value == value.to_integral_value() else float(value)

        result.append({
            'A': row.get('記載頁', row.get('カタログ頁', '')),
            'B': row.get('商品名', ''), 'C': row.get('商品記号', ''),
            'D': row.get('仕様寸法', ''), 'E': excel_number(qty),
            'F': row.get('単位', ''), 'K': excel_number(price),
            'L': excel_number(unit_cost),
        })
    return pd.DataFrame(result, columns=list(TRANSFER_COL_MAP), dtype=object)


def read_itoki_csv(csv_path: str | Path) -> list[dict[str, str]]:
    """CP932 / UTF-8のCSVを読み込む。空欄、ゼロ、先頭ゼロを保持する。

    見出し行や空の明細行も削除しない。列名の差（制御/LCなど）も保持する。
    """
    raw = Path(csv_path).read_bytes()
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = raw.decode('cp932')
    reader = csv.DictReader(io.StringIO(text, newline=''))
    headers = reader.fieldnames or []
    missing = [name for name in REQUIRED_COLUMNS if name not in headers]
    if missing:
        raise ValueError(f'イトーキCSVの必須列がありません: {", ".join(missing)}')
    if len(headers) != len(set(headers)):
        raise ValueError('CSVの列名が重複しています。')
    missing = [
        name for name in OUTPUT_COLUMNS
        if name not in headers and COLUMN_ALIASES.get(name) not in headers
    ]
    if missing:
        raise ValueError(f'並べ替えに必要な列がありません: {", ".join(missing)}')
    records = []
    for row in reader:
        if None in row or any(value is None for value in row.values()):
            raise ValueError(f'CSV {reader.line_num}行付近の列数が一致しません。')
        records.append(row)
    return records


def itoki(csv_path: str | Path, output_path: str | Path | None = None) -> Path:
    """指定の11列を別CSVへ出力する。数値の計算・行の削除は行わない。"""
    source = Path(csv_path)
    output = Path(output_path) if output_path else source.with_name(
        f'{source.stem}_itoki_converted.csv'
    )
    if output.suffix.lower() != '.csv':
        raise ValueError('出力ファイルの拡張子は.csvを指定してください。')
    if source.resolve() == output.resolve():
        raise ValueError('入力と出力には別のファイルを指定してください。')
    rows = prepare_csv_rows(read_itoki_csv(source))
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.writer(stream)
        writer.writerow(OUTPUT_COLUMNS)
        for row in rows:
            writer.writerow([row[name] for name in OUTPUT_COLUMNS])
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description='イトーキCSVを指定の11列に並べ替え')
    parser.add_argument('csv_path', type=Path)
    parser.add_argument('-o', '--output', type=Path, help='出力CSVのパス')
    args = parser.parse_args()
    output = itoki(args.csv_path, args.output)
    print(f'出力完了: {output}')


if __name__ == '__main__':
    main()
