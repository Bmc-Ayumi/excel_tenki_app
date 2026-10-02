"""オフィス見積の加工（1シート・複数シート）と転記列対応。"""
import re
import math
import openpyxl
import mojimoji



TRANSFER_COL_MAP = {'B': 'C', 'C': 'D', 'D': 'E', 'E': 'F', 'F': 'G', 'K': 'H'}


def get_data_until_blank(ws, start_col, start_row):
    data = []
    row = start_row
    while row <= ws.max_row:
        cell_val = ws.cell(row=row, column=start_col).value
        data.append((row, cell_val))
        row += 1
    return data


def normalize_sheet_name(name: str) -> str:
    # 全角数字を半角に寄せる（例：３→3）
    try:
        return mojimoji.zen_to_han(name)
    except Exception:
        return name


def consolidate_selected_sheets(wb, target_sheet_names, add_blank_row=True):
    """
    target_sheet_names の先頭シートに、2枚目以降のデータを追記して1枚にまとめる
    """
    if not target_sheet_names or len(target_sheet_names) <= 1:
        return

    base_ws = wb[target_sheet_names[0]]

    def is_row_blank(ws, r, max_col):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                return False
        return True

    base_max_col = base_ws.max_column
    base_last = base_ws.max_row
    while base_last > 1 and is_row_blank(base_ws, base_last, base_max_col):
        base_last -= 1
    write_row = base_last + 1

    for src_name in target_sheet_names[1:]:
        if src_name not in wb.sheetnames:
            continue

        src_ws = wb[src_name]
        max_col = max(base_max_col, src_ws.max_column)

        for r in range(1, src_ws.max_row + 1):
            row_vals = []
            all_blank = True

            for c in range(1, max_col + 1):
                v = src_ws.cell(row=r, column=c).value
                row_vals.append(v)
                if v is not None and str(v).strip() != "":
                    all_blank = False

            if all_blank:
                continue

            for c, v in enumerate(row_vals, start=1):
                base_ws.cell(row=write_row, column=c).value = v

            write_row += 1

        if add_blank_row:
            write_row += 1

        wb.remove(src_ws)


def office_int_multi(xlsx_path: str, start_sheet_name: str):
    """
    オフィスインテリア（複数シート版）
    ・プルダウンで選んだシート以降を処理
    ・不要行削除、文字整形、単価補完、列削除
    ・最後に対象シートを1枚に統合
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    all_sheet_names = wb.sheetnames

    if start_sheet_name not in all_sheet_names:
        wb.close()
        raise ValueError(f"指定した開始シートが見つかりません: {start_sheet_name}")

    start_idx = all_sheet_names.index(start_sheet_name)

    # ★ 選択したシート以降をすべて対象にする
    target_sheet_names = all_sheet_names[start_idx:]

    if not target_sheet_names:
        wb.close()
        raise ValueError("処理対象シートがありません")

    target_sheets = [wb[name] for name in target_sheet_names]

    for ws in target_sheets:
        # 0) 行削除
        for r in range(ws.max_row, 0, -1):
            b_val = get_cell_value_with_merge(ws, r, 2)
            b_norm = normalize_text(b_val)

            c_val = get_cell_value_with_merge(ws, r, 3)
            c_norm = normalize_text(c_val)

            cm_all_empty = True
            for col in range(3, 14):  # C～M
                v = get_cell_value_with_merge(ws, r, col)
                if normalize_text(v) != "":
                    cm_all_empty = False
                    break

            hi_text = ""
            for col in (8, 9, 10, 11):  # H, I, J, K
                hi_text += normalize_text(get_cell_value_with_merge(ws, r, col))

            if (
                ("件名" in b_norm)
                or ("名称" in b_norm)
                or ("合計" in c_norm)
                or ("内訳書" in hi_text or "内訳明細書" in hi_text)
                or cm_all_empty
            ):
                ws.delete_rows(r)

        # ① C列・I列の先頭スペース削除
        for row in range(1, ws.max_row + 1):
            v = get_cell_value_with_merge(ws, row, 3)  # C
            if isinstance(v, str):
                ws.cell(row=row, column=3).value = v.lstrip(" 　")

            v = get_cell_value_with_merge(ws, row, 9)  # I
            if isinstance(v, str):
                ws.cell(row=row, column=9).value = v.lstrip(" 　")

        # ② 単価を金額から補完。「式」は金額÷数量（小計は除外）。
        for row in range(1, ws.max_row + 1):
            c_val = get_cell_value_with_merge(ws, row, 3)   # C
            l_val = get_cell_value_with_merge(ws, row, 12)  # L
            m_val = get_cell_value_with_merge(ws, row, 13)  # M

            c_norm2 = normalize_text(c_val)
            c_norm2 = re.sub(r"[（）\(\)【】\[\]{}]", "", c_norm2)

            if "小計" in c_norm2:
                continue

            if normalize_text(l_val) == "" and m_val is not None:
                ws.cell(row=row, column=12).value = supplement_unit_price(
                    m_val,
                    get_cell_value_with_merge(ws, row, 10),  # J: 数量
                    get_cell_value_with_merge(ws, row, 11),  # K: 単位
                )

        # ③ EFG列を削除
        ws.delete_cols(5, 3)

        # ④ さらにE列を削除
        ws.delete_cols(5)

    # ★ 選択したシート以降を、先頭シートに統合
    consolidate_selected_sheets(wb, target_sheet_names, add_blank_row=True)

    output_path = xlsx_path.replace(".xlsx", "_office_multi_converted.xlsx")
    wb.save(output_path)
    wb.close()
    return output_path


def get_cell_value_with_merge(ws, row, col):
    """
    結合セル対応：
    指定セルが結合範囲内なら、左上セルの値を返す
    """
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col
            ).value
    return None


def normalize_text(v):
    """半角・全角スペースを除去（Noneは空扱い）"""
    if v is None:
        return ""
    if not isinstance(v, str):
        v = str(v)
    return re.sub(r"[\s\u3000]+", "", v)


def supplement_unit_price(amount, quantity, unit):
    """式単価は金額を数量で割る。それ以外の単位は従来どおり金額を返す。"""
    if normalize_text(unit) != "式":
        return amount
    try:
        qty = float(mojimoji.zen_to_han(normalize_text(quantity)).replace(",", ""))
        value = float(mojimoji.zen_to_han(normalize_text(amount)).replace(",", ""))
    except (ValueError, TypeError):
        return None
    if not math.isfinite(qty) or not math.isfinite(value) or qty <= 0:
        return None
    return value / qty


def office_int(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows_to_delete = [idx for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2)
                      if row[1] is None or not isinstance(row[1], (int, float))]
    for idx in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(idx)

    ws.delete_cols(17)
    ws.delete_cols(15)
    ws.delete_cols(13)
    ws.delete_cols(8, 2)
    ws.delete_cols(4, 3)
    ws.insert_cols(4)

    pattern_total = re.compile(r'\b小\s*計\b')
    pattern_mid_total = re.compile(r'\b中\s*計\b')  # ★ 追加
    pattern_discount = re.compile(r'\b値\s*引\b')


    for row in ws.iter_rows(min_row=2):
        h_cell = row[7]
        c_cell = row[2]

        # ★ C列が完全に空白の行は処理しないで次へ進む
        if c_cell.value is None:
            continue

        c_text = str(c_cell.value)

        # ★ 中計 も 小計 と同様に除外する
        if (
            (normalize_text(h_cell.value) == "")
            and (not pattern_total.search(c_text))
            and (not pattern_mid_total.search(c_text))   # ← ここ追加
            and (not pattern_discount.search(c_text))
        ):
            h_cell.value = supplement_unit_price(row[8].value, row[5].value, row[6].value)
    for row in ws.iter_rows(min_row=2):
        for cell in row[2:5]:
            if cell.value is not None and isinstance(cell.value, str):
                cell.value = cell.value.strip()

    output_path = xlsx_path.replace(".xlsx", "_converted.xlsx")
    wb.save(output_path)
    wb.close()
    return output_path
