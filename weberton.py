r"""ウェバートン見積PDFの2ページ目以降をCSVに抽出する独立CLI。

実行例:
    .venv\Scripts\python.exe weberton.py ウェバートン_PDF.pdf
    .venv\Scripts\python.exe weberton.py 入力.pdf -o 出力.csv

main.pyやExcelへの書き込みは行わない。文字情報と罫線を持つ見積PDFが対象。
"""
from __future__ import annotations

import argparse
import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata

import fitz


COLUMNS = ['ページ', '区分', '品名', 'メーカー', '型番', '数量', '単位', '単価', '金額', '定価']
SECTION_NAMES = {'材料機器費内訳': '材料費', '工事費内訳': '工事費', '諸経費内訳': '諸経費'}
# main.pyのwrite_df_to_template_comへ渡せる列対応（転記先: DataFrame列）。
TRANSFER_COL_MAP = {column: column for column in ('B', 'C', 'E', 'F', 'I', 'K')}


def compact(text: str) -> str:
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', text))


def numeric(text: str) -> str:
    value = compact(text).replace(',', '').replace('−', '-').replace('￥', '').replace('¥', '')
    if not value:
        return ''
    try:
        number = Decimal(value)
    except InvalidOperation as exc:
        raise ValueError(f'数値を読み取れません: {text!r}') from exc
    if not number.is_finite():
        raise ValueError(f'不正な数値です: {text!r}')
    return format(number, 'f')


def extract_weberton_pdf(pdf_path: str | Path) -> list[dict[str, str]]:
    """表を原稿順に返す。行種別列は作らず、小計・値引・合計は保持する。"""
    records = []
    section = ''
    with fitz.open(pdf_path) as document:
        if len(document) < 2:
            raise ValueError('2ページ目以降がありません。')
        for index in range(1, len(document)):
            tables = document[index].find_tables().tables
            matched = False
            for table in tables:
                raw_rows = table.extract()
                header = [compact(cell or '') for cell in raw_rows[0]]
                if len(header) != 9 or header[4:9] != ['数量', '単位', '単価', '金額', '定価']:
                    continue
                matched = True
                for raw in raw_rows[1:]:
                    if len(raw) != 9:
                        raise ValueError(f'{index + 1}ページ: 表の列数が一致しません。')
                    cells = [(cell or '').strip() for cell in raw]
                    if not any(cells):
                        continue
                    if cells[0] and not any(cells[4:]):
                        section = cells[1]
                        continue
                    name = cells[1]
                    if not name:
                        raise ValueError(f'{index + 1}ページ: 品名が空の行があります。')
                    if compact(name) in ('小計', 'お値引', '値引', '合計'):
                        section = '全体'
                    values = [str(index + 1), section, name, cells[2], cells[3],
                              numeric(cells[4]), cells[5], numeric(cells[6]),
                              numeric(cells[7]), numeric(cells[8])]
                    records.append(dict(zip(COLUMNS, values)))
            if not matched:
                raise ValueError(f'{index + 1}ページ: 対応する表がありません。画像PDFや別形式は未対応です。')
    if not records:
        raise ValueError('明細を抽出できませんでした。')
    return records


def validate_amounts(records: list[dict[str, str]]) -> int:
    """原稿の数値を補正せず照合する。不一致はCSV保存前に報告する。"""
    section_sum = Decimal(0)
    detail_sum = Decimal(0)
    discount = Decimal(0)
    detail_count = 0
    totals = []
    for row in records:
        name = compact(row['品名'])
        if not row['金額']:
            raise ValueError(f"金額がありません: {row['品名']}")
        amount = Decimal(row['金額'])
        if name == '小計':
            expected = detail_sum
        elif name == '合計':
            totals.append(amount)
            continue
        elif name in ('お値引', '値引'):
            discount += amount
            continue
        elif name.endswith('小計'):
            expected = section_sum
            section_sum = Decimal(0)
        else:
            if not row['数量'] or not row['単価']:
                raise ValueError(f"数量または単価がありません: {row['品名']}")
            expected = Decimal(row['数量']) * Decimal(row['単価'])
            section_sum += amount
            detail_sum += amount
            detail_count += 1
        if expected != amount:
            raise ValueError(f"金額不一致: {row['品名']}（計算値 {expected}、PDF {amount}）")
    if len(totals) != 1 or totals[0] != detail_sum + discount:
        raise ValueError('合計が見つからない、複数存在する、または明細・値引と一致しません。')
    return detail_count


def prepare_csv_rows(records: list[dict[str, str]]) -> list[dict[str, str]]:
    """照合済みの原稿から小計・合計を除き、区分先頭に見出し行を追加する。"""
    output_rows = []
    previous_section = None
    for record in records:
        if '小計' in compact(record['品名']) or compact(record['品名']) == '合計':
            continue
        row = dict(record)
        section = SECTION_NAMES.get(compact(row['区分']), row['区分'])
        row['区分'] = section
        if section != previous_section and section in SECTION_NAMES.values():
            heading = dict.fromkeys(COLUMNS, '')
            heading.update({'ページ': row['ページ'], '区分': section, '品名': f'【{section}】'})
            output_rows.append(heading)
        output_rows.append(row)
        previous_section = section
    return output_rows


def write_csv(records: list[dict[str, str]], output: str | Path) -> None:
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(records)


def build_transfer_df(records: list[dict[str, str]]):
    """既存のExcel転記関数向けに列を整える。main.pyはインポートしない。

    見出し・値引は保持し、小計と合計は空白を除去して判定し除外する。
    金額列は指定されていないため転記しない。
    """
    import pandas as pd

    required = {'品名', 'メーカー', '型番', '数量', '単位', '単価', '定価'}
    result = []

    def excel_number(value):
        text = numeric(str(value or ''))
        if not text:
            return ''
        number = Decimal(text)
        return int(number) if number == number.to_integral_value() else float(number)

    for row in records:
        missing = required - row.keys()
        if missing:
            raise ValueError(f'転記に必要な列がありません: {", ".join(sorted(missing))}')
        name = row['品名']
        if not name.strip() or '小計' in compact(name) or compact(name) == '合計':
            continue
        result.append({
            'B': name,
            'C': ' '.join(value.strip() for value in (row['メーカー'], row['型番']) if value.strip()),
            'E': excel_number(row['数量']),
            'F': row['単位'],
            'I': excel_number(row['定価']),
            'K': excel_number(row['単価']),
        })
    if not result:
        raise ValueError('転記対象の行がありません。')
    return pd.DataFrame(result, columns=list(TRANSFER_COL_MAP), dtype=object)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('pdf', type=Path, nargs='?', default=Path(__file__).with_name('ウェバートン_PDF.pdf'))
    parser.add_argument('-o', '--output', type=Path, help='出力CSV（省略時は入力PDFと同じフォルダ）')
    parser.add_argument('--transfer-preview', type=Path, help='転記先の列記号を列名とする確認用CSV')
    args = parser.parse_args()
    output = args.output or args.pdf.with_name(args.pdf.stem + '_2ページ以降.csv')
    if output.resolve() == args.pdf.resolve():
        parser.error('出力先に入力PDFと同じパスは指定できません。')
    if args.transfer_preview and args.transfer_preview.resolve() in (args.pdf.resolve(), output.resolve()):
        parser.error('転記確認CSVは入力PDF・抽出CSVと別のパスにしてください。')
    try:
        records = extract_weberton_pdf(args.pdf)
        count = validate_amounts(records)
        output_rows = prepare_csv_rows(records)
        write_csv(output_rows, output)
        if args.transfer_preview:
            transfer_df = build_transfer_df(output_rows)
            args.transfer_preview.parent.mkdir(parents=True, exist_ok=True)
            transfer_df.to_csv(args.transfer_preview, index=False, encoding='utf-8-sig')
    except (OSError, ValueError, RuntimeError) as exc:
        parser.exit(1, f'抽出失敗: {exc}\n')
    print(f'抽出完了: 明細{count}件 / 見出し・値引を含む{len(output_rows)}行（小計・合計除外）。金額照合OK。')
    print(f'保存先: {output.resolve()}')
    if args.transfer_preview:
        print(f'転記列確認: {args.transfer_preview.resolve()}')


if __name__ == '__main__':
    main()
