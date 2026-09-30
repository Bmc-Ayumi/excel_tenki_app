"""帝国倉庫PDFの明細抽出・加工と転記列対応。"""
import re
from dataclasses import dataclass
import fitz
import pandas as pd



TRANSFER_COL_MAP = {'B': 'A', 'E': 'D', 'F': 'E', 'K': 'F', 'I': 'H'}


@dataclass(frozen=True)
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str


def _word_close(a: Word, b: Word, tolerance: float = 2.0) -> bool:
    return (
        abs(a.x0 - b.x0) <= tolerance
        and abs(a.y0 - b.y0) <= tolerance
        and abs(a.x1 - b.x1) <= tolerance
        and abs(a.y1 - b.y1) <= tolerance
    )


def _dedupe_words(words: list[Word]) -> list[Word]:
    deduped: list[Word] = []
    for w in sorted(words, key=lambda item: (round(item.y0, 1), item.x0, item.text)):
        if any(_word_close(w, existing) for existing in deduped):
            continue
        deduped.append(w)
    return deduped


def load_words(page: fitz.Page, y_min: float, y_max: float) -> list[Word]:
    words: list[Word] = []
    for x0, y0, x1, y1, text, *_rest in page.get_text("words"):
        if y_min <= y0 <= y_max:
            words.append(Word(float(x0), float(y0), float(x1), float(y1), str(text)))
    return words


def cluster_rows(words: list[Word], tolerance: float = 2.8) -> list[float]:
    ys = sorted({round(w.y0, 1) for w in words})
    rows: list[float] = []
    for y in ys:
        if not rows or abs(y - rows[-1]) > tolerance:
            rows.append(y)
    return rows


def row_for_y(y: float, rows: list[float]) -> float:
    return min(rows, key=lambda row_y: abs(row_y - y))


def build_detail_rows(page: fitz.Page) -> list[list[str]]:
    # 固定の行座標に頼らず、ページ内の実際の文字位置から行を作る。
    words = load_words(page, 205, float(page.rect.height))
    row_centers = cluster_rows(words, tolerance=2.8)
    if not row_centers:
        return [["\u9805\u76ee", "", "", "\u6570\u91cf", "\u5358\u4f4d", "\u5358\u4fa1", "\u91d1\u984d", "\u5099\u8003"]]

    rows: dict[float, dict[int, list[tuple[float, str]]]] = {
        row: {i: [] for i in range(1, 8)} for row in row_centers
    }

    for w in words:
        row = row_for_y(w.y0, row_centers)
        if w.x0 < 90:
            col = 1
        elif w.x0 < 220:
            col = 2
        elif w.x0 < 300:
            col = 3
        elif w.x0 < 335:
            col = 4
        elif w.x0 < 390:
            col = 5
        elif w.x0 < 455:
            col = 6
        else:
            col = 7
        rows[row][col].append((w.x0, w.text))

    ordered_rows: list[list[str]] = [["\u9805\u76ee", "", "", "\u6570\u91cf", "\u5358\u4f4d", "\u5358\u4fa1", "\u91d1\u984d", "\u5099\u8003"]]
    for row_y in sorted(rows):
        cols = rows[row_y]
        values = []
        for idx in range(1, 8):
            parts = [text for _, text in sorted(cols[idx], key=lambda item: item[0])]
            values.append("".join(parts).strip())
        if any(values):
            if len(ordered_rows) > 1 and not any(values[:6]) and values[6]:
                prev = ordered_rows[-1]
                prev[6] = (prev[6] + " " + values[6]).strip() if prev[6] else values[6]
                continue
            ordered_rows.append(values)
    return ordered_rows


def extract_teisoh_misc_amount(page: fitz.Page) -> str:
    try:
        words = page.get_text("words") or []
        label_words = [
            (float(w[0]), float(w[1]))
            for w in words
            if str(w[4]).strip() in {"諸経費", "諸経費用"}
        ]
        numeric_words = [
            (float(w[0]), float(w[1]), str(w[4]).strip())
            for w in words
            if re.fullmatch(r"[0-9,]+(?:円)?", str(w[4]).strip())
        ]
        for x0, y0 in label_words:
            same_row = [
                (x, text)
                for x, y, text in numeric_words
                if abs(y - y0) <= 8 and x > x0
            ]
            if same_row:
                return re.sub(r"[^\d,]", "", sorted(same_row, key=lambda item: item[0])[0][1])
    except Exception:
        pass

    return ""


def is_teisoh_subtotal_row(row: pd.Series) -> bool:
    label = "".join(str(row.get(col, "")).strip() for col in ["A", "B"]).replace(" ", "")
    return label == "小計"


def is_teisoh_subtotal_output_row(values: list[object]) -> bool:
    text = "".join(str(v).strip() for v in values).replace(" ", "")
    if "小計" in text:
        return True
    return "小" in text and "計" in text


def append_teisoh_misc_row_from_amount(df: pd.DataFrame, amount_text: str) -> pd.DataFrame:
    if df.empty or not amount_text:
        return df
    misc_row = pd.DataFrame(
        [["諸経費", "", "", "1", "式", amount_text, amount_text, ""]],
        columns=df.columns,
    )
    return pd.concat([df, misc_row], ignore_index=True)


def normalize_teisoh_item_name(value: object) -> object:
    if not isinstance(value, str):
        return value

    text = value.strip()
    if text == "2t":
        return "2t車"
    if text == "4t":
        return "4t車"
    return text


def split_merged_quantity_and_price(quantity: object, unit_price: object) -> tuple[object, object]:
    def _split(value: object) -> tuple[str, str]:
        if not isinstance(value, str):
            return "", ""
        text = value.strip().replace(" ", "")
        m = re.match(r"^(-?\d{1,3})(\d{1,3}(?:,\d{3})+)$", text)
        if not m:
            return "", ""
        return m.group(1), m.group(2)

    quantity_text = "" if quantity is None else str(quantity).strip()
    unit_price_text = "" if unit_price is None else str(unit_price).strip()

    if quantity_text and not unit_price_text:
        q, p = _split(quantity_text)
        if q and p:
            return q, p

    if unit_price_text and not quantity_text:
        q, p = _split(unit_price_text)
        if q and p:
            return q, p

    return quantity, unit_price


def load_teisoh_detail_df(pdf_path: str) -> pd.DataFrame:
    doc = fitz.open(pdf_path)
    try:
        detail_rows: list[list[str]] = []
        # 1ページ目は表紙なので、2ページ目以降だけを明細として読む。
        for page_index in range(1, len(doc)):
            page_rows = build_detail_rows(doc[page_index])
            if not page_rows:
                continue
            if not detail_rows:
                detail_rows.extend(page_rows)
            else:
                detail_rows.extend(page_rows[1:])
        misc_amount = extract_teisoh_misc_amount(doc[0]) if len(doc) > 0 else ""
    finally:
        doc.close()

    source_df = pd.DataFrame(detail_rows[1:], columns=["A", "B", "C", "D", "E", "F", "G"]).fillna("")
    source_df["B"] = source_df["B"].map(normalize_teisoh_item_name)

    split_pairs = source_df.apply(
        lambda row: split_merged_quantity_and_price(row["D"], row["C"]),
        axis=1,
        result_type="expand",
    )
    source_df["D"] = split_pairs[0]
    source_df["C"] = split_pairs[1]
    source_df = source_df[~source_df.apply(is_teisoh_subtotal_row, axis=1)].reset_index(drop=True)

    output_rows: list[list[str]] = [["項目", "", "", "数量", "単位", "単価", "金額", "備考"]]
    for _, row in source_df.iterrows():
        item_text = str(row["B"]).strip()
        if not item_text:
            continue
        output_rows.append(
            [
                item_text,
                "",
                "",
                str(row["D"]).strip(),
                str(row["E"]).strip(),
                str(row["C"]).strip(),
                str(row["F"]).strip(),
                str(row["G"]).strip(),
            ]
        )

    output_rows = [output_rows[0]] + [row for row in output_rows[1:] if not is_teisoh_subtotal_output_row(row)]

    df = pd.DataFrame(output_rows, columns=["A", "B", "C", "D", "E", "F", "G", "H"])
    if misc_amount:
        df = append_teisoh_misc_row_from_amount(df, misc_amount)
    df.index += 1
    return df
