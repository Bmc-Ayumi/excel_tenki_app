import streamlit as st
from office import office_int, office_int_multi, TRANSFER_COL_MAP as OFFICE_COL_MAP
from comany import comany, TRANSFER_COL_MAP as COMANY_COL_MAP
from teisoh import load_teisoh_detail_df, TRANSFER_COL_MAP as TEISOH_COL_MAP
from openpyxl.utils import column_index_from_string, get_column_letter
import tempfile
from weberton import (
    COLUMNS as WEBERTON_COLUMNS,
    TRANSFER_COL_MAP as WEBERTON_COL_MAP,
    extract_weberton_pdf,
    validate_amounts as validate_weberton_amounts,
    prepare_csv_rows as prepare_weberton_rows,
    build_transfer_df as build_weberton_transfer_df,
)
import os
from openpyxl import load_workbook
import pyexcel as p
import re
import pandas as pd
import win32com.client as win32
import datetime
import pythoncom
import shutil
import glob
import tempfile

# ===== アプリ専用 Temp（ここに一時ファイルを集約）=====
APP_TEMP_DIR = os.path.join(tempfile.gettempdir(), "excel_tenki_app")
os.makedirs(APP_TEMP_DIR, exist_ok=True)

TESSERACT_DIR_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
]
for _tesseract_dir in TESSERACT_DIR_CANDIDATES:
    _tesseract_exe = os.path.join(_tesseract_dir, "tesseract.exe")
    if os.path.exists(_tesseract_exe):
        if _tesseract_dir not in os.environ.get("PATH", ""):
            os.environ["PATH"] = _tesseract_dir + os.pathsep + os.environ.get("PATH", "")
        _tessdata_dir = os.path.join(_tesseract_dir, "tessdata")
        if os.path.isdir(_tessdata_dir):
            os.environ.setdefault("TESSDATA_PREFIX", _tessdata_dir)
        break

def cleanup_app_temp():
    """このアプリが作った一時ファイルだけ削除する（Temp全体は触らない）"""
    try:
        shutil.rmtree(APP_TEMP_DIR, ignore_errors=True)
        os.makedirs(APP_TEMP_DIR, exist_ok=True)
    except Exception:
        pass


SUPPLIER_TARGET_SHEET = {
    "オフィスインテリア": "見積明細（建築・内装）",
    "コマニー": "見積明細（コマニー）",
}


def validate_dest_cell(dest_cell: str, expected_col: str = "B") -> int:
    m = re.match(r"^([A-Z]+)(\d+)$", (dest_cell or "").upper().strip())
    if not m:
        raise ValueError("転記先セルは 'B3' の形式で入力してください")
    col, row = m.group(1), int(m.group(2))
    if col != expected_col:
        raise ValueError(f"転記先セルは {expected_col}列スタートで入力してください（例：{expected_col}{row}）")
    return row

def get_row_from_cell(cell_addr: str) -> int:
    m = re.match(r"^([A-Z]+)(\d+)$", (cell_addr or "").upper().strip())
    if not m:
        raise ValueError("セルは 'C2' の形式で入力してください")
    return int(m.group(2))


import shutil

def write_df_to_template_com(
    template_path: str,
    out_path: str,
    sheet_name: str,
    start_cell: str,
    df_src,
    col_map: dict,
):
    # ✅ 新規のときだけコピー（同一ファイルならコピーしない）
    if os.path.abspath(template_path) != os.path.abspath(out_path):
        shutil.copy(template_path, out_path)


    start_cell = (start_cell or "").strip().upper()
    start_row = validate_dest_cell(start_cell, expected_col="B")

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # UpdateLinks=0 でリンク更新を避ける（余計な警告・遅延防止）
        wb = excel.Workbooks.Open(out_path, UpdateLinks=0)
        ws = wb.Worksheets(sheet_name)

        n_rows = len(df_src)
        if n_rows == 0:
            raise ValueError("転記データが空です")

        # 1列ずつ書き込み（式列は触らない＝潰れない）
        for tmpl_col_letter, src_col_letter in col_map.items():
            c = column_index_from_string(tmpl_col_letter)  # テンプレ列番号

            # 縦1列の2次元配列（Excel Rangeに渡す形）
            col_values = []
            for i in range(n_rows):
                v = df_src.iloc[i][src_col_letter] if src_col_letter in df_src.columns else ""
                if v is None:
                    v = ""
                col_values.append([v])

            r0 = start_row
            r1 = start_row + n_rows - 1
            ws.Range(ws.Cells(r0, c), ws.Cells(r1, c)).Value = col_values

        # 数式再計算
        excel.CalculateFullRebuild()

        wb.Save()  # out_path に保存（ロゴ保持）
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


def get_col_map(company: str) -> dict:
    maps = {
        "オフィス（１シート）": OFFICE_COL_MAP,
        "オフィス（複数シート）": OFFICE_COL_MAP,
        "コマニー": COMANY_COL_MAP,
        "帝国倉庫PDF": TEISOH_COL_MAP,
        "ウェバートンPDF": WEBERTON_COL_MAP,
    }
    if company not in maps:
        raise ValueError(f"未対応の仕入先です: {company}")
    return maps[company].copy()


st.set_page_config(page_title="Excel転記アプリ", layout="wide")


# ▼ 新規スタート（状態クリア）
if st.sidebar.button("🧹 新規スタート（データクリア）"):
    cleanup_app_temp()

    # ★ uploaderを作り直すための番号を進める
    st.session_state["uploader_reset_counter"] = st.session_state.get("uploader_reset_counter", 0) + 1

    for k in [
        "result_path",
        "template_path",
        "template_wb",
        "last_result_path",
        "download_name",
        "updated_template_df_map",
        "prev_mode",
        "original_template_path",
        "office_multi_start_sheet",
        "sheet_selector",
        "src_cell",
        "dest_cell",
        "output_name",
        "prev_company",
    ]:
        if k in st.session_state:
            del st.session_state[k]

    st.rerun()

# ===== 転記結果 自動復元 =====
RESULT_DIR = os.path.join(os.path.expanduser("~"), "Excel転記アプリ_results")
os.makedirs(RESULT_DIR, exist_ok=True)

def find_latest_result() -> str | None:
    files = glob.glob(os.path.join(RESULT_DIR, "*.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

# ===== 転記ロック（同時実行防止）=====
LOCK_FILE = os.path.join(RESULT_DIR, "transfer.lock")
LOCK_TIMEOUT_SEC = 10 * 60  # 10分

def _now_ts():
    return datetime.datetime.now().timestamp()

def is_lock_stale(timeout_sec=LOCK_TIMEOUT_SEC) -> bool:
    try:
        with open(LOCK_FILE, "r", encoding="utf-8") as f:
            ts = float(f.read().strip() or "0")
        return (_now_ts() - ts) > timeout_sec
    except:
        return False

def acquire_lock() -> bool:
    # 古いロックは自動解除
    if os.path.exists(LOCK_FILE) and is_lock_stale():
        try:
            os.remove(LOCK_FILE)
        except:
            pass

    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(str(_now_ts()))
        return True
    except FileExistsError:
        return False

def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except:
        pass


# --- SessionState 初期化（必須）---
if "last_result_path" not in st.session_state:
    st.session_state.last_result_path = None
if "template_wb" not in st.session_state:
    st.session_state.template_wb = None
if "template_path" not in st.session_state:
    st.session_state.template_path = None
if "original_template_path" not in st.session_state:
    st.session_state.original_template_path = None
if "updated_template_df_map" not in st.session_state:
    st.session_state.updated_template_df_map = {}

st.title("仕入先見積変換＆転記アプリ")


# --- 共通処理 ---
def save_uploaded_bytes(file_bytes, suffix=".xls"):
    os.makedirs(APP_TEMP_DIR, exist_ok=True)
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=APP_TEMP_DIR)
    tmp_file.write(file_bytes)
    tmp_file.close()
    return tmp_file.name

def convert_xls_to_xlsx(xls_path):
    xlsx_path = xls_path + "x"
    p.save_book_as(file_name=xls_path, dest_file_name=xlsx_path)
    return xlsx_path


# --- UI ---

# ▼ 既存UI
company = st.sidebar.selectbox(
    "仕入先を選択してください",
    ["オフィス（１シート）", "オフィス（複数シート）", "コマニー", "帝国倉庫PDF", "ウェバートンPDF"]
)
# 仕入先変更を検知
prev_company = st.session_state.get("prev_company")

if prev_company is None:
    st.session_state["prev_company"] = company
elif prev_company != company:
    # ★ 仕入先が変わったら、仕入先ファイル uploader だけ初期化
    st.session_state["uploader_reset_counter"] = st.session_state.get("uploader_reset_counter", 0) + 1

    # ★ 仕入先側の状態だけ消す
    for k in [
        "office_multi_start_sheet",
        "src_cell",
        "dest_cell",
        "preview_sheet_name",
        "current_supplier_name",
    ]:
        if k in st.session_state:
            del st.session_state[k]

    # ★ 処理済み見積の表示用もクリア
    converted_excel_path = None
    selected_start_sheet = None
    df_converted = None

    # ★ 新しい仕入先を記録
    st.session_state["prev_company"] = company
    st.rerun()

uploader_reset_counter = st.session_state.get("uploader_reset_counter", 0)

is_pdf_supplier = company in ("帝国倉庫PDF", "ウェバートンPDF")
supplier_file_type = ["pdf"] if is_pdf_supplier else ["xls", "xlsx"]
supplier_file_label = "仕入先PDFファイルをアップロードしてください" if is_pdf_supplier else "仕入先ファイル（xls/.xlsx）をアップロードしてください"

supplier_file = st.file_uploader(
    supplier_file_label,
    type=supplier_file_type,
    key=f"supplier_file_{uploader_reset_counter}"
)

converted_excel_path = None
selected_start_sheet = None
df_converted = None
preview_sheet_name = None

# ★ ここで先に処理開始シートを表示
if supplier_file:
    ext = os.path.splitext(supplier_file.name)[1].lower()
    supplier_file_bytes = supplier_file.getvalue()
    raw_path = save_uploaded_bytes(supplier_file_bytes, suffix=ext)
    if company == "ウェバートンPDF":
        try:
            weberton_records = extract_weberton_pdf(raw_path)
            validate_weberton_amounts(weberton_records)
            df_converted = pd.DataFrame(prepare_weberton_rows(weberton_records), columns=WEBERTON_COLUMNS)
        except (OSError, ValueError, RuntimeError) as exc:
            st.error(f"ウェバートンPDFの抽出に失敗しました: {exc}")
            st.stop()
        converted_excel_path = os.path.join(APP_TEMP_DIR, "weberton_pdf_converted.xlsx")
        with pd.ExcelWriter(converted_excel_path, engine="openpyxl") as writer:
            df_converted.to_excel(writer, sheet_name="ウェバートンPDF", index=False)
        preview_sheet_name = "ウェバートンPDF"
    elif company == "帝国倉庫PDF":
        df_converted = load_teisoh_detail_df(raw_path)
        converted_excel_path = os.path.join(APP_TEMP_DIR, "teisoh_pdf_converted.xlsx")
        with pd.ExcelWriter(converted_excel_path, engine="openpyxl") as writer:
            df_converted.to_excel(writer, sheet_name="帝国倉庫PDF", index=False)
        preview_sheet_name = "帝国倉庫PDF"
    else:
        xlsx_path = convert_xls_to_xlsx(raw_path) if ext == ".xls" else raw_path

        if company == "オフィス（複数シート）":
            raw_wb = load_workbook(xlsx_path, data_only=True)
            raw_sheet_options = raw_wb.sheetnames
            raw_wb.close()

            selected_start_sheet = st.selectbox(
                "処理開始シートを選択してください（このシート以降を処理）",
                raw_sheet_options,
                key="office_multi_start_sheet"
            )

# ★ selectbox のあとにテンプレート uploader
template_file = st.file_uploader(
    "見積テンプレートExcel（.xlsx）をアップロード",
    type=["xlsx"],
    key=f"template_file_{uploader_reset_counter}"
)

# ★ 変換処理
if supplier_file:
    if company == "オフィス（１シート）":
        converted_excel_path = office_int(xlsx_path)

    elif company == "オフィス（複数シート）":
        converted_excel_path = office_int_multi(xlsx_path, selected_start_sheet)

    elif company == "コマニー":
        converted_excel_path = comany(xlsx_path)

    if converted_excel_path is None:
        if company == "コマニー":
            st.error("この見積ファイルには『内訳書』シートが無いため、コマニーの変換ができませんでした。")
        else:
            st.error(f"{company} の変換に失敗しました。（対応シートが無い/形式が違う可能性があります）")
        st.stop()

    wb = load_workbook(converted_excel_path)

    if company == "オフィス（複数シート）" and selected_start_sheet:
        preview_sheet_name = selected_start_sheet
    else:
        preview_sheet_name = wb.sheetnames[0]

    ws = wb[preview_sheet_name]
    data = list(ws.values)
    df_converted = pd.DataFrame(data)
    df_converted.columns = [get_column_letter(i + 1) for i in range(len(df_converted.columns))]
    df_converted.index += 1
    df_converted = df_converted.fillna("")


# ★「初回だけ」テンプレを確定させる（2回目以降の rerun では上書きしない）
# テンプレは「新規モード」では毎回これを基準にしたいので、アップロードがあれば更新してOK
if template_file:
    template_bytes = template_file.getvalue()
    template_path = save_uploaded_bytes(template_bytes, suffix=".xlsx")
    st.session_state.original_template_path = template_path

    if not st.session_state.get("result_path"):
        st.session_state.template_path = template_path
        st.session_state.template_wb = load_workbook(template_path)


# テンプレート未読込 → 見積だけ単体表示
if df_converted is not None and not st.session_state.template_wb:
    st.subheader("処理済み見積ファイルのプレビュー")
    st.dataframe(df_converted, use_container_width=True)

# テンプレートが読み込まれている場合 → 両方表示
elif df_converted is not None and st.session_state.template_wb:
    sheet_name = st.selectbox("見積テンプレートシートを選んでください", st.session_state.template_wb.sheetnames, key="sheet_selector")
    ws_template = st.session_state.template_wb[sheet_name]
    template_data = list(ws_template.values)
    df_template = pd.DataFrame(template_data).fillna("")
    df_template.columns = [get_column_letter(i+1) for i in range(len(df_template.columns))]
    df_template.index += 1

    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"処理済み見積ファイルのプレビュー（{preview_sheet_name}）")
        st.dataframe(df_converted, use_container_width=True)
    with col2:
        # シート別の転記結果があればそれを表示
        if (
            "updated_template_df_map" in st.session_state
            and sheet_name in st.session_state.updated_template_df_map
        ):
            st.subheader(f"見積テンプレートシート：転記結果（{sheet_name}）")
            st.data_editor(
                st.session_state.updated_template_df_map[sheet_name],
                use_container_width=True,
                disabled=True,
                hide_index=False
            )
        else:
            st.subheader(f"見積テンプレートシート: {sheet_name}")
            st.data_editor(
                df_template,
                use_container_width=True,
                disabled=True,
                hide_index=False
            )


    # --- ★ ここから転記UI ---
    st.markdown("### 転記設定（開始セルと転記先セルを入力）")

    src = st.text_input("転記元の開始セル（例: C14）", key="src_cell")
    dest = st.text_input("転記先セル（例: B3）", key="dest_cell")
    output_name = st.text_input("出力ファイル名（拡張子不要）", value="merged_result", key="output_name")


    if st.button("✅ 一括転記する"):
        # ① 入力チェック（ロック前）
        src  = (st.session_state.get("src_cell")  or "").strip().upper()
        dest = (st.session_state.get("dest_cell") or "").strip().upper()

        if not src:
            st.error("転記元の開始セル（例: C2）を入力してください")
            st.stop()
        if not dest:
            st.error("転記先セル（例: B2）を入力してください")
            st.stop()

        # ② ロック獲得
        if not acquire_lock():
            st.warning("現在、他のユーザーが転記中です。少し待ってください。")
            st.stop()

        try:
            src_row = get_row_from_cell(src)
            if company == "ウェバートンPDF":
                src_row = max(2, src_row)  # 1行目の列名は転記しない
            df_to_write = df_converted.loc[src_row:].copy()

            cols_check = [c for c in ["A","B","C","D","E","F"] if c in df_to_write.columns]
            if cols_check:
                df_to_write = df_to_write[
                    ~df_to_write[cols_check]
                    .astype(str)
                    .apply(lambda x: x.str.strip())
                    .eq("")
                    .all(axis=1)
                ]

            if company == "ウェバートンPDF":
                source_columns = {
                    get_column_letter(i + 1): name
                    for i, name in enumerate(WEBERTON_COLUMNS)
                }
                df_to_write = build_weberton_transfer_df(
                    df_to_write.rename(columns=source_columns).to_dict("records")
                )

            # ここで必ず「今画面に入っている値」を取り直す
            output_name = (st.session_state.get("output_name") or "merged_result").strip()

            # ファイル名に使えない文字を除去（念のため）
            output_name = re.sub(r'[\\/:*?"<>|]', '_', output_name)

            ts = datetime.datetime.now().strftime("%m%d")

            # ===== 差し替え形式：継続でもファイル名を更新しつつ、最終的に1ファイルにする =====
            last_result_path = st.session_state.get("last_result_path")
            original_template_path = st.session_state.get("original_template_path")

            # 入力ファイル名（拡張子つき）
            new_filename = f"{ts}_{output_name}.xlsx"
            new_path = os.path.join(RESULT_DIR, new_filename)

            if last_result_path:
                # 継続：基本は上書き
                base_template_path = last_result_path
                result_path = last_result_path

                # ただし、ユーザーが別名を入れたら「その名前に切り替える」
                if os.path.basename(last_result_path) != new_filename:
                    # 既に同名があるなら消しておく（上書き許可）
                    if os.path.exists(new_path):
                        os.remove(new_path)

                    shutil.copy(last_result_path, new_path)
                    base_template_path = new_path
                    result_path = new_path

            else:
                # 1回目（新規）
                base_template_path = original_template_path
                if not base_template_path:
                    raise RuntimeError("テンプレートが未設定です。先にテンプレートExcelをアップロードしてください。")
                result_path = new_path

            # ======================================================================


            # ★★★ ここまで ★★★

            write_df_to_template_com(
                template_path=base_template_path,
                out_path=result_path,
                sheet_name=sheet_name,
                start_cell=dest,
                df_src=df_to_write,
                col_map=get_col_map(company),
            )
            # --- 差し替え：前回ファイル名と今回のファイル名が違うなら、前回ファイルを削除して1本化 ---
            if last_result_path and os.path.abspath(last_result_path) != os.path.abspath(result_path):
                try:
                    os.remove(last_result_path)
                except Exception:
                    pass


            st.session_state["download_name"] = os.path.basename(result_path)
            st.session_state["last_result_path"] = result_path  # 次回継続の基準
            st.session_state["result_path"] = result_path       # 画面表示/ダウンロード用
            st.session_state.template_path = result_path
            st.session_state.template_wb = load_workbook(result_path)


            st.success("転記が完了しました")
            release_lock()   # ✅ rerun前に解除（保険）
            st.rerun()

        except Exception as e:
            st.error(str(e))
            release_lock()
            st.stop()

        finally:
            release_lock()


# ★ 一括転記後だけ、ダウンロードボタンを表示
if "result_path" in st.session_state:
    # いま画面に入っている出力名を「DL名」に反映（転記後に変えてもOK）
    dl_base = (st.session_state.get("output_name") or "merged_result").strip()
    dl_base = re.sub(r'[\\/:*?"<>|]', '_', dl_base)  # 禁止文字除去

    # ts をファイル名に付けたいなら：転記時に使った ts を保存しておくのが理想
    # いったん簡易に「保存先ファイルの先頭4桁(0109_)を流用」する例：
    current_base = os.path.basename(st.session_state["result_path"])
    prefix = current_base.split("_", 1)[0]  # 0109 だけ取り出す（無い場合もある）

    # prefix が4桁じゃなければ付けない
    if prefix.isdigit() and len(prefix) == 4:
        dl_name = f"{prefix}_{dl_base}.xlsx"
    else:
        dl_name = f"{dl_base}.xlsx"

    with open(st.session_state["result_path"], "rb") as f:
        st.download_button(
            "転記済みファイルをダウンロード",
            f,
            file_name=dl_name,
            key="download_btn"  # key固定（表示が安定）
        )

    st.write("保存先:", st.session_state.get("result_path"))
    st.write("DL名:", dl_name)


                # --- ★ ここまで転記UI ---

