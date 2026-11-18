from __future__ import annotations

import argparse
import csv
import io
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
import warnings
from decimal import Decimal
from pathlib import Path

import fitz
import pandas as pd
from PIL import Image, ImageOps


COLUMNS = ["ページ", "項目", "数量", "単位", "単価", "金額", "備考"]
TRANSFER_COL_MAP = {column: column for column in ("B", "C", "E", "F", "K")}


def build_transfer_df(records: list[dict[str, str]]) -> pd.DataFrame:
    """数量・単価を数値化し、単価はテンプレートのK列へ渡す。"""
    def number(value):
        text = unicodedata.normalize("NFKC", str(value)).replace(",", "").strip()
        if not text:
            return ""
        if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text):
            raise ValueError(f"転記する数値を確認してください: {value}")
        parsed = Decimal(text)
        return int(parsed) if parsed == parsed.to_integral_value() else float(parsed)

    result = []
    for row in records:
        name = str(row["項目"])
        if not name.strip():
            continue
        result.append({
            "B": name, "C": row["備考"],
            "E": number(row["数量"]), "F": row["単位"],
            "K": number(row["単価"]),
        })
    if not result:
        raise ValueError("転記対象の行がありません。")
    return pd.DataFrame(result, columns=list(TRANSFER_COL_MAP), dtype=object)


def _row(page: str, item: str, quantity: str = "", unit: str = "", unit_price: str = "", amount: str = "", remark: str = "") -> dict[str, str]:
    return {
        "ページ": page,
        "項目": item,
        "数量": quantity,
        "単位": unit,
        "単価": unit_price,
        "金額": amount,
        "備考": remark,
    }




def _rate(condition: str, unit: str, unit_price: str) -> dict[str, str]:
    return {"condition": condition, "unit": unit, "unit_price": unit_price}


SHIBUSAWA_RATE_MASTER: dict[str, list[dict[str, str]]] = {
    "平台車": [_rate("", "台", "")],
    "施工費": [_rate("", "式", "")],
    "器材費": [_rate("", "式", "")],
    "養生費": [_rate("", "式", "")],
    "諸経費": [_rate("", "式", "")],
    "班長": [
        _rate("日中", "名", "18,000"),
        _rate("時間外", "名", "2,800"),
        _rate("深夜時間外", "名", "3,300"),
        _rate("夜間", "名", "28,000"),
    ],
    "一般作業員": [
        _rate("日中", "名", "16,000"),
        _rate("時間外", "h", "2,500"),
        _rate("深夜時間外", "h", "3,000"),
        _rate("夜間", "名", "25,000"),
    ],
    "人員輸送車": [
        _rate("", "車", "10,000"),
    ],
    "4t車": [
        _rate("日中", "車", "40,000"),
        _rate("時間外", "h", "3,500"),
        _rate("深夜時間外", "h", "4,000"),
        _rate("夜間", "車", "58,000"),
    ],
    "2t車": [
        _rate("日中", "車", "35,000"),
        _rate("時間外", "h", "3,500"),
        _rate("深夜時間外", "h", "4,000"),
        _rate("夜間", "車", "50,000"),
    ],
    "ワンタッチケース": [_rate("", "枚", "250")],
    "アルバムケース": [_rate("", "枚", "220")],
    "ノートPCケース": [_rate("", "枚", "150")],
    "デスクトップPCケース": [_rate("", "枚", "400")],
    "布テープ": [_rate("", "個", "350")],
    "養生テープ": [_rate("", "", "")],
    "ラベル": [_rate("", "枚", "15")],
    "エアキャップ": [_rate("", "本", "4,500")],
    "ブルーシートロール": [_rate("", "本", "5,000")],
    "資材配達・回収費": [_rate("", "式", "20,000")],
}


SHIBUSAWA_RATE_INDEX: dict[tuple[str, str, str], str] = {
    (item, rate["condition"], rate["unit"]): rate["unit_price"]
    for item, rates in SHIBUSAWA_RATE_MASTER.items()
    for rate in rates
}


def lookup_shibusawa_rate(item: str, condition: str, unit: str) -> str | None:
    return SHIBUSAWA_RATE_INDEX.get((item, condition, unit))


def _lookup_shibusawa_unit(item: str, condition: str = "") -> str:
    """Return a condition-specific unit, falling back to the item's first unit."""
    rates = SHIBUSAWA_RATE_MASTER.get(item, [])
    if condition:
        for rate in rates:
            if rate["condition"] == condition and rate["unit"]:
                return rate["unit"]
    for rate in rates:
        if rate["unit"]:
            return rate["unit"]
    return ""


def _fill_shibusawa_units(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    for row in rows:
        item = row.get(COLUMNS[1], "")
        normalized_item = re.sub(r"\s+", "", unicodedata.normalize("NFKC", item)).rstrip("*")
        specified_unit = ""
        if "班長" in normalized_item:
            specified_unit = "名"
        elif normalized_item == "一般":
            specified_unit = "名"
        elif "エアキャップ" in normalized_item:
            specified_unit = "本"
        elif normalized_item in {"人員輸送費", "資材配達費"}:
            specified_unit = "式"
        elif "養生" in normalized_item or "器材" in normalized_item:
            specified_unit = "式"
        elif normalized_item in {"エア袋", "ダンボール"}:
            specified_unit = "枚"
        if specified_unit:
            row[COLUMNS[3]] = specified_unit
            continue
        if (
            re.sub(r"\s+", "", item).rstrip("*") == "諸経費"
            and not row.get("数量", "").strip()
            and not row.get("金額", "").strip()
        ):
            # 数量・金額のない諸経費は見出しなので単位を付けない。
            row[COLUMNS[3]] = ""
            continue
        if not row.get(COLUMNS[3]):
            text = f"{item} {row.get(COLUMNS[6], '')}"
            conditions = {
                rate["condition"]
                for rate in SHIBUSAWA_RATE_MASTER.get(item, [])
                if rate["condition"] and rate["condition"] in text
            }
            condition = next(iter(conditions), "")
            row[COLUMNS[3]] = _lookup_shibusawa_unit(item, condition)
            if not row[COLUMNS[3]]:
                if not any(row.get(key, "").strip() for key in ("数量", "単価", "金額")):
                    continue
                # Fee rows such as "overtime expense" are not master items,
                # but their condition name still determines the unit.
                condition_rates = [
                    (rate["condition"], rate["unit"])
                    for rates in SHIBUSAWA_RATE_MASTER.values()
                    for rate in rates
                    if rate["condition"] and rate["unit"]
                ]
                for condition_name, unit in sorted(
                    condition_rates,
                    key=lambda value: len(value[0]),
                    reverse=True,
                ):
                    if condition_name in text:
                        row[COLUMNS[3]] = unit
                        break
    return rows


def _ensure_tesseract_env() -> None:
    local_tessdata = Path(__file__).with_name("tessdata")
    if local_tessdata.is_dir():
        os.environ["TESSDATA_PREFIX"] = str(local_tessdata)

    candidates = [
        r"C:\Program Files\Tesseract-OCR",
        r"C:\Program Files (x86)\Tesseract-OCR",
    ]
    for tesseract_dir in candidates:
        if not os.path.isdir(tesseract_dir):
            continue
        tesseract_exe = os.path.join(tesseract_dir, "tesseract.exe")
        if os.path.exists(tesseract_exe):
            if tesseract_dir not in os.environ.get("PATH", ""):
                os.environ["PATH"] = tesseract_dir + os.pathsep + os.environ.get("PATH", "")
            tessdata_dir = os.path.join(tesseract_dir, "tessdata")
            if os.path.isdir(tessdata_dir):
                os.environ.setdefault("TESSDATA_PREFIX", tessdata_dir)
            break


def _tesseract_languages() -> set[str]:
    try:
        result = subprocess.run(
            ["tesseract", "--list-langs"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except FileNotFoundError:
        return set()

    langs: set[str] = set()
    for line in result.stdout.splitlines():
        line = line.strip()
        if not line or line.startswith("List of available languages"):
            continue
        langs.add(line)
    return langs


def _tesseract_lang_arg() -> str:
    _ensure_tesseract_env()
    langs = _tesseract_languages()
    if "jpn" in langs:
        # Japanese is the primary language for these forms. Using only the
        # Japanese model avoids English recognition taking precedence in text
        # cells such as item names.
        return "jpn"
    if "eng" in langs:
        warnings.warn(
            "Tesseract の日本語言語データ (jpn.traineddata) が見つかりません。"
            " OCR は英語のみになり、日本語帳票は文字化けしやすくなります。",
            RuntimeWarning,
            stacklevel=2,
        )
        return "eng"
    return "osd"


def _group_projection_runs(values: list[int], threshold: int) -> list[int]:
    candidates = [index for index, value in enumerate(values) if value >= threshold]
    runs: list[list[int]] = []
    for index in candidates:
        if not runs or index > runs[-1][-1] + 1:
            runs.append([index])
        else:
            runs[-1].append(index)
    return [round(sum(run) / len(run)) for run in runs if len(run) <= 12]


def _render_page_image(page: fitz.Page) -> Image.Image:
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def _detect_table_grid(image: Image.Image) -> tuple[list[int], list[int]]:
    gray = ImageOps.grayscale(image)
    dark = gray.point(lambda pixel: 1 if pixel < 170 else 0)
    width, height = gray.size
    pixels = dark.load()

    horizontal_counts = [sum(pixels[x, y] for x in range(width)) for y in range(height)]
    horizontal_candidates = _group_projection_runs(
        horizontal_counts,
        max(80, int(width * 0.28)),
    )
    horizontal_candidates = [y for y in horizontal_candidates if y >= int(height * 0.18)]
    if len(horizontal_candidates) < 6:
        raise ValueError("表の横罫線を検出できませんでした")

    table_top = horizontal_candidates[0]
    table_bottom = horizontal_candidates[-1]
    vertical_counts = [
        sum(pixels[x, y] for y in range(table_top, table_bottom + 1))
        for x in range(width)
    ]
    vertical_candidates = _group_projection_runs(
        vertical_counts,
        max(80, int((table_bottom - table_top) * 0.35)),
    )
    vertical_candidates = [x for x in vertical_candidates if 0 < x < width - 1]
    if len(vertical_candidates) < 6:
        raise ValueError("表の縦罫線を検出できませんでした")

    # A rule inside the remarks column does not divide the other cells.
    # Allow small vertical deviations in scanned rules. Testing only their
    # centre pixel row can drop a real boundary and merge adjacent amounts.
    left, right = vertical_candidates[0], vertical_candidates[-1]
    horizontal_candidates = [
        y for y in horizontal_candidates
        if sum(
            any(pixels[x, near_y] for near_y in range(max(0, y - 3), min(height, y + 4)))
            for x in range(left, right)
        ) >= (right - left) * 0.75
    ]
    if len(horizontal_candidates) < 6:
        raise ValueError("表全体の横罫線を検出できませんでした")
    return horizontal_candidates, vertical_candidates


def _ocr_image_crop(image: Image.Image, numeric: bool = False) -> str:
    _ensure_tesseract_env()
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        image.save(tmp_path)
        result = subprocess.run(
            [
                "tesseract",
                tmp_path,
                "stdout",
                "-l",
                "eng" if numeric else _tesseract_lang_arg(),
                "--psm",
                "7",
                *(["-c", "tessedit_char_whitelist=0123456789,."] if numeric else []),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        return " ".join(result.stdout.split())
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def _crop_has_ink(image: Image.Image) -> bool:
    gray = ImageOps.grayscale(image)
    dark_pixels = sum(gray.histogram()[:150])
    return dark_pixels >= max(12, int(image.width * image.height * 0.002))


def _read_table_cell(image: Image.Image, box: tuple[int, int, int, int], numeric: bool = False) -> str:
    """Read cell-local subdivisions without cutting text in adjacent columns."""
    x0, y0, x1, y1 = box
    cell = ImageOps.grayscale(image.crop((x0 + 8, y0, x1 - 8, y1)))
    pixels = cell.load()
    rules = _group_projection_runs(
        [sum(pixels[x, y] < 170 for x in range(cell.width)) for y in range(cell.height)],
        int(cell.width * 0.75),
    )
    boundaries = [0] + [y for y in rules if 5 < y < cell.height - 5] + [cell.height]
    values = []
    for top, bottom in zip(boundaries, boundaries[1:]):
        if bottom - top < 12:
            continue
        crop = cell.crop((0, top + 3, cell.width, bottom - 3))
        if not _crop_has_ink(crop):
            continue
        crop = ImageOps.expand(crop, border=12, fill="white")
        value = _ocr_image_crop(crop)
        if numeric and not _cell_number(value):
            value = _ocr_image_crop(crop, numeric=True)
        values.append(_clean_coordinate_cell(value))
    return " ".join(value for value in values if value)


def _clean_coordinate_cell(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    return re.sub(r"\s+", "", value).strip()


def _cell_number(value: str) -> str:
    value = re.sub(r"\s+", "", unicodedata.normalize("NFKC", value))
    value = value.lstrip("¥￥\\").removesuffix("円")
    # Do not accept only the first digit of an OCR result such as 2ー乞,300.
    if not re.fullmatch(r"\d[\d,.]*", value):
        return ""
    return _normalize_numeric_token(value)


def _extract_coordinate_rows_for_page(doc: fitz.Document, page_index: int) -> list[dict[str, str]]:
    image = _render_page_image(doc[page_index])
    horizontal_lines, vertical_lines = _detect_table_grid(image)
    page_label = str(page_index)
    rows: list[dict[str, str]] = []

    for y0, y1 in zip(horizontal_lines, horizontal_lines[1:]):
        if y1 - y0 < 18:
            continue
        cells: list[str] = []
        for x0, x1 in zip(vertical_lines, vertical_lines[1:]):
            if x1 - x0 < 20:
                continue
            cells.append(_read_table_cell(image, (x0, y0, x1, y1), numeric=len(cells) in {1, 2, 3}))
        if len(cells) < 5:
            continue

        item, quantity, unit_price, amount, remark = cells[:5]
        compact_item = item.replace(" ", "")
        if compact_item in {"項目", "小計"}:
            if compact_item == "小計":
                rows.append(_row(page_label, "小計", amount=_cell_number(amount)))
            continue
        if not item and not any((quantity, unit_price, amount, remark)):
            continue

        quantity = _cell_number(quantity)
        unit_price = _cell_number(unit_price)
        amount = _cell_number(amount)
        item = _canonical_item(item, quantity, unit_price, amount, remark)
        rows.append(_row(page_label, item, quantity, "", unit_price, amount, remark))

    return rows


def _ocr_page_text(doc: fitz.Document, page_index: int) -> str:
    _ensure_tesseract_env()
    page = doc[page_index]
    scale = 3
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        pix.save(tmp_path)
        psm = "3"
        lang = _tesseract_lang_arg()
        result = subprocess.run(
            ["tesseract", tmp_path, "stdout", "-l", lang, "--psm", psm],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        return result.stdout
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def _normalize_ocr_line(line: str) -> str:
    line = line.translate(str.maketrans({
        "①": "1",
        "②": "2",
        "③": "3",
        "④": "4",
        "⑤": "5",
        "⑥": "6",
        "⑦": "7",
        "⑧": "8",
        "⑨": "9",
        "⑩": "10",
        "⑪": "11",
        "⑫": "12",
        "⑬": "13",
        "⑭": "14",
        "⑮": "15",
        "⑯": "16",
        "⑰": "17",
        "⑱": "18",
        "⑲": "19",
        "⑳": "20",
    }))
    line = line.replace("：", ":").replace("，", ",").replace("．", ".")
    line = line.replace("－", "-").replace("〜", "〜").replace("～", "〜")
    line = line.replace("—xt", "一式").replace("—st", "一式")
    line = line.replace("一 思", "一式").replace("ー 式", "一式").replace("—=", "一式")
    line = line.replace("«10%", "※10%").replace("Â«10%", "※10%")
    line = re.sub(r"\s*:\s*", ":", line)
    line = re.sub(r"\s*/\s*", "/", line)
    line = re.sub(r"(?<=\d)[,\.]\s+(?=\d{3}\b)", ",", line)
    line = re.sub(r"(?<=\d)\.\s+(?=\d{3}\b)", ",", line)
    line = re.sub(r"(?<=\d)\s*,\s*(?=\d{3}\b)", ",", line)
    line = re.sub(r"\s+", " ", line)
    return line.strip()


def _is_number_like(token: str) -> bool:
    token = token.strip().lstrip("¥")
    token = token.translate(str.maketrans({
        "①": "1",
        "②": "2",
        "③": "3",
        "④": "4",
        "⑤": "5",
        "⑥": "6",
        "⑦": "7",
        "⑧": "8",
        "⑨": "9",
        "⑩": "10",
        "⑪": "11",
        "⑫": "12",
        "⑬": "13",
        "⑭": "14",
        "⑮": "15",
        "⑯": "16",
        "⑰": "17",
        "⑱": "18",
        "⑲": "19",
        "⑳": "20",
    }))
    return bool(re.fullmatch(r"\d[\d,]*", token))


def _normalize_numeric_token(token: str) -> str:
    token = token.strip().lstrip("¥￥")
    token = re.sub(r"^[\-\u2212\u2012\u2013\u2014\u2015\uFF0D]+", "", token)
    token = token.translate(str.maketrans({
        "①": "1",
        "②": "2",
        "③": "3",
        "④": "4",
        "⑤": "5",
        "⑥": "6",
        "⑦": "7",
        "⑧": "8",
        "⑨": "9",
        "⑩": "10",
        "⑪": "11",
        "⑫": "12",
        "⑬": "13",
        "⑭": "14",
        "⑮": "15",
        "⑯": "16",
        "⑰": "17",
        "⑱": "18",
        "⑲": "19",
        "⑳": "20",
        "０": "0",
        "１": "1",
        "２": "2",
        "３": "3",
        "４": "4",
        "５": "5",
        "６": "6",
        "７": "7",
        "８": "8",
        "９": "9",
        "，": ",",
        "．": ".",
    }))
    token = token.replace(" ", "")
    if re.fullmatch(r"\d{1,3}(?:,\d{3})+", token):
        return token
    if re.fullmatch(r"\d+(?:,\d+)?", token):
        return f"{int(token.replace(',', '')):,}"
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", token):
        return token.replace(".", ",")
    if re.fullmatch(r"\d+(?:\.\d+)?", token):
        return f"{int(token.replace('.', '')):,}"
    return ""


def _parse_page_label(ocr_text: str, fallback: str) -> str:
    text = _normalize_ocr_line(ocr_text)
    m = re.search(r"\bNO\.?\s*(\d+)\b", text, re.IGNORECASE)
    return m.group(1) if m else fallback


def _canonical_item(item: str, quantity: str = "", unit_price: str = "", amount: str = "", remark: str = "") -> str:
    item = item.strip().replace(" ", "")

    aliases = {
        "4t": "4t車",
        "4Ｔ": "4t車",
        "2t": "2t車",
        "2Ｔ": "2t車",
        "車輌費": "車両費",
        "車両費": "車両費",
        "作業員費": "作業員費",
        "資材費": "資材費",
        "資材配達・回収費": "資材配達・回収費",
        "諸経費": "諸経費",
        "養生費": "養生費",
        "養生器材費": "養生器材費",
        "器材費": "器材費",
        "人員輸送費": "人員輸送費",
        "人員輸送車": "人員輸送車",
        "一般作業員": "一般作業員",
        # Exact OCR aliases only: do not replace characters in site names.
        "一舩": "一般",
        "一舷": "一般",
        "ー般": "一般",
        "班長": "班長",
        "深夜時間外費": "深夜時間外費",
        "時間外費": "時間外費",
        "時間外費M": "時間外費",
        "ダンボール": "ダンボール",
        "エアキャップ": "エアキャップ",
        "エア袋": "エア袋",
        "ラベル": "ラベル",
        "小計": "小計",
    }
    if item in aliases:
        return aliases[item]

    if item == "小計":
        return "小計"
    return item


def _parse_shibusawa_line(line: str, page_label: str) -> dict[str, str] | None:
    line = _normalize_ocr_line(line)
    if not line:
        return None

    skip_markers = ["見積明細書", "作業日", "JWITS", "項目", "数量", "単価", "金額", "備考"]
    compact_line = line.replace(" ", "")
    if any(marker in compact_line for marker in skip_markers):
        return None
    if re.fullmatch(r"NO\.?\s*\d*", line, re.IGNORECASE):
        return None
    if re.fullmatch(r"\d{1,2}/\d{1,2}\s*\([^)]*\)\s*.*", line):
        m = re.match(r"^(\d{1,2}/\d{1,2})\s*\(([^)]*)\)\s*(.*)$", line)
        if m:
            date_key = m.group(1)
            weekday = m.group(2).strip()
            item = f"{date_key}({weekday})"
            remark = m.group(3).strip()
            return _row(page_label, item, remark=remark)
        m = re.match(r"^(\d{1,2}/\d{1,2}\s*\([^)]*\))\s*(.*)$", line)
        item = m.group(1).replace(" ", "")
        remark = m.group(2).strip()
        return _row(page_label, item, remark=remark)

    if line in {"車両費", "作業員費", "資材費"}:
        return _row(page_label, line)

    tokens = line.split()
    if not tokens:
        return None

    # Merge split subtotal labels such as "小 計".
    if len(tokens) >= 2 and tokens[0] == "小" and tokens[1] in {"計", "Ey", "EI", "Es"}:
        tokens = ["小計"] + tokens[2:]

    remark = ""
    if len(tokens) >= 2 and not _is_number_like(tokens[-1]) and any(_is_number_like(tok) for tok in tokens[:-1]):
        remark = tokens[-1]
        tokens = tokens[:-1]

    num_idx = next((i for i, tok in enumerate(tokens) if _is_number_like(tok)), None)
    if num_idx is None:
        return None

    item = "".join(tokens[:num_idx]).strip()
    if not item:
        return None
    nums = []
    for tok in tokens[num_idx:]:
        num = _normalize_numeric_token(tok)
        if num:
            nums.append(num)

    if not nums:
        return None

    quantity = ""
    unit_price = ""
    amount = ""
    if len(nums) == 1:
        if item == "小計":
            amount = nums[0]
        else:
            unit_price = nums[0]
    elif len(nums) == 2:
        if nums[0] == nums[1]:
            quantity = "1"
            unit_price, amount = nums
        else:
            quantity, unit_price = nums
    else:
        quantity, unit_price, amount = nums[:3]

    item = _canonical_item(item, quantity=quantity, unit_price=unit_price, amount=amount, remark=remark)
    return _row(page_label, item, quantity, "", unit_price, amount, remark)


def _extract_ocr_rows_for_page(doc: fitz.Document, page_index: int) -> list[dict[str, str]]:
    ocr_text = _ocr_page_text(doc, page_index)
    page_label = _parse_page_label(ocr_text, fallback=str(page_index))
    rows: list[dict[str, str]] = []
    for raw_line in ocr_text.splitlines():
        parsed = _parse_shibusawa_line(raw_line, page_label)
        if parsed is not None:
            rows.append(parsed)
            if parsed.get("項目") == "小計":
                break

    return rows



def _prepare_shibusawa_rows(rows: list[dict[str, str]], messages: list[str] | None = None) -> list[dict[str, str]]:
    """数量・単価・金額のうち一つの欠落を補い、参考の時間外費を除く。"""
    def parse_number(value: str) -> Decimal | None:
        value = unicodedata.normalize("NFKC", value).strip()
        if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?", value):
            return None
        return Decimal(value.replace(",", ""))

    def format_number(value: Decimal) -> str:
        text = format(value, ",f")
        return text.rstrip("0").rstrip(".") if "." in text else text

    result: list[dict[str, str]] = []
    for source in rows:
        row = source.copy()
        quantity = row.get("数量", "").strip()
        item = re.sub(r"\s+", "", row.get("項目", ""))
        # Reference overtime rates must be excluded before quantity inference.
        if "時間外費" in item and not quantity:
            continue
        if "諾経費" in item:
            item = item.replace("諾経費", "諸経費")
            row["項目"] = item
        if "ー再作業員" in item or "一再作業員" in item:
            item = item.replace("ー再作業員", "一般作業員").replace("一再作業員", "一般作業員")
            row["項目"] = item
        quantity_number = parse_number(quantity)
        unit_price = parse_number(row.get("単価", ""))
        amount = parse_number(row.get("金額", ""))
        recalculate_misc_quantity = item.rstrip("*") == "諸経費" and quantity_number != Decimal("1")
        if (not quantity or quantity_number == 0 or recalculate_misc_quantity) and unit_price is not None and unit_price != 0 and amount is not None:
            row["数量"] = format_number(amount / unit_price)
        if not row.get("金額", "").strip() and quantity_number is not None and unit_price is not None:
            row["金額"] = format_number(quantity_number * unit_price)
        if not row.get("単価", "").strip() and quantity_number is not None and quantity_number != 0 and amount is not None:
            row["単価"] = format_number(amount / quantity_number)
        if messages is not None:
            for field, formula in (("数量", "金額÷単価"), ("金額", "数量×単価"), ("単価", "金額÷数量")):
                if row.get(field, "") != source.get(field, ""):
                    messages.append(
                        f"明細ページ{row['ページ']}・{len(result) + 1}行目「{row['項目']}」："
                        f"{field}を{formula}で補完しました（{row[field]}）。"
                    )
        result.append(row)
    return result


def extract_shibusawa_detail_rows(pdf_path: str) -> pd.DataFrame:
    """明細を返す。画面用の補完通知は attrs['completion_messages'] に保持する。"""
    _ensure_tesseract_env()
    with fitz.open(pdf_path) as doc:
        rows: list[dict[str, str]] = []
        construction_section = False
        # 澁澤の見積書は表紙の次ページから明細を読み取る。
        for page_index in range(1, doc.page_count):
            page_text = doc[page_index].get_text()
            if not page_text.strip():
                page_text = _ocr_page_text(doc, page_index)
            # 明細中の「施工費（詳細別紙）」という項目ではなく、別紙の見出しで判定。
            heading = re.sub(r"\s+", "", unicodedata.normalize("NFKC", page_text[:400]))
            if "見積明細内訳表" in heading:
                construction_section = False
            elif re.search(r"[【\[「]施工費[】\]」]", heading):
                construction_section = True
            if construction_section:
                continue  # 「前ページより」の施工費続紙も除外する。
            try:
                page_rows = _extract_coordinate_rows_for_page(doc, page_index)
            except ValueError:
                # 罫線を検出できない場合も、実際のOCR結果だけから抽出する。
                page_rows = _extract_ocr_rows_for_page(doc, page_index)
            rows.extend(page_rows)
        messages: list[str] = []
        df = pd.DataFrame(_fill_shibusawa_units(_prepare_shibusawa_rows(rows, messages)), columns=COLUMNS)
        # DataFrameの列には入れず、CSV・Excel・転記の対象から分離する。
        df.attrs["completion_messages"] = messages
        return df



def save_detail_csv(df: pd.DataFrame, output_path: str) -> str:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(out, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
        return str(out)
    except PermissionError:
        fallback = out.with_name(f"{out.stem}_new{out.suffix}")
        df.to_csv(fallback, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
        return str(fallback)


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="澁澤PDFの明細をCSV化します。PDFを渡すだけで自動判定します。")
    parser.add_argument("pdf_path", help="入力PDFパス")
    parser.add_argument("-o", "--output", default=None, help="出力CSVパス")
    args = parser.parse_args()

    pdf_path = args.pdf_path
    output_path = args.output or f"{Path(pdf_path).stem}_detail.csv"

    df = extract_shibusawa_detail_rows(pdf_path)
    saved_path = save_detail_csv(df, output_path)

    print(f"rows: {len(df)}")
    print(f"saved: {saved_path}")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
