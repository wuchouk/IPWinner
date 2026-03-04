import re
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.drawing.image import Image
from openpyxl.utils.units import pixels_to_EMU
from io import BytesIO
from datetime import datetime, timezone, timedelta
import streamlit.components.v1 as st_components
import os, tempfile, json

APP_VERSION = "v9.1"
APP_LAST_UPDATE = "2026-03-04 20:15"

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="商標監控資料合併工具",
    page_icon="📋",
    layout="centered",
)

# 自訂 CSS：拖放上傳區域放大 + 視覺提示 + 檔案列表一次顯示更多
st.markdown("""
<style>
/* ---- 放大拖放上傳區域 ---- */
[data-testid="stFileUploaderDropzone"] {
    min-height: 220px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    border: 2px dashed #b0b8c8;
    border-radius: 12px;
    background: #f8fafd;
    transition: border-color 0.2s, background 0.2s;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #4F8BF9;
    background: #eef3fc;
}
/* 讓拖放區域內的文字和按鈕居中 */
[data-testid="stFileUploaderDropzone"] > div {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
}
/* 隱藏 file_uploader 分頁按鈕 */
[data-testid="stFileUploader"] nav[role="navigation"],
[data-testid="stFileUploader"] [data-testid="stPagination"],
[data-testid="stFileUploader"] .stPagination {
    display: none !important;
}
/* 檔案列表容器高度 */
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileList"] {
    max-height: 700px !important;
    overflow-y: auto !important;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# 偵測使用者瀏覽器時區（用於下載檔名時間戳）
# ============================================================
@st.cache_resource
def _create_tz_component():
    """建立瀏覽器時區偵測元件（使用 Streamlit declare_component）"""
    comp_dir = os.path.join(tempfile.gettempdir(), "_st_tz_detect")
    os.makedirs(comp_dir, exist_ok=True)
    with open(os.path.join(comp_dir, "index.html"), "w") as f:
        f.write("""<!DOCTYPE html>
<html><body><script>
function send(type, data) {
    window.parent.postMessage(
        Object.assign({isStreamlitMessage: true, type: type}, data), "*"
    );
}
send("streamlit:componentReady", {apiVersion: 1});
window.addEventListener("message", function(event) {
    if (event.data.type === "streamlit:render") {
        send("streamlit:setComponentValue",
             {value: new Date().getTimezoneOffset()});
    }
});
</script></body></html>""")
    return st_components.declare_component("_tz_detect", path=comp_dir)

_tz_component = _create_tz_component()
_client_tz_offset = _tz_component(default=0, key="_tz_offset", height=0)


def _get_client_now():
    """取得使用者本地時間（根據瀏覽器時區）"""
    if isinstance(_client_tz_offset, (int, float)):
        # getTimezoneOffset() 回傳「UTC - 本地」的分鐘數，所以要取反
        client_tz = timezone(timedelta(minutes=-int(_client_tz_offset)))
        return datetime.now(client_tz)
    return datetime.now(timezone.utc)


# ============================================================
# 合併紀錄持久化（JSON 檔案）
# ============================================================
_HISTORY_FILE = os.path.join(os.path.dirname(__file__), '.merge_history.json')


def _load_merge_history():
    """從 JSON 檔讀取今日合併紀錄；若非今日或檔案不存在則回傳空 list"""
    try:
        with open(_HISTORY_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if data.get('date') == _get_client_now().strftime('%Y-%m-%d'):
            return data.get('records', [])
    except (FileNotFoundError, json.JSONDecodeError, ValueError):
        pass
    return []


def _save_merge_history(records):
    """將合併紀錄寫入 JSON 檔（含今日日期）"""
    data = {
        'date': _get_client_now().strftime('%Y-%m-%d'),
        'records': records,
    }
    try:
        with open(_HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # Streamlit Cloud 等環境可能寫入失敗，靜默處理


# ============================================================
# 欄位對應設定
# ============================================================
DB1_MAPPING = {
    'A': 'B', 'B': 'C', 'C': 'D', 'D': 'E', 'E': 'F',
    'F': 'G', 'G': 'H', 'H': 'I', 'I': 'J', 'J': 'K', 'K': 'L',
}
DB1_IMAGE_SRC_COL = 1

# DB2 有兩種格式：
# 格式 A（有地區欄）：A:序号, B:商标文字, C:商标图样, D:地區, E:类别, F:申请号, G:申请人, H:申请日期, I:初审公告日期, J:初审公告期号, K:异议截止日期, L:群组, M:商品/服务
# 格式 B（無地區欄）：A:序号, B:商标文字, C:商标图样, D:类别, E:申请号, F:申请人, G:申请日期, H:初审公告日期, I:初审公告期号, J:异议截止日期, K:群组, L:商品/服务
DB2_MAPPING_WITH_REGION = {
    'B': 'B',   # 商标文字 → 他人商標
    'C': 'C',   # 商标图样 → 商標圖樣
    'D': 'D',   # 地區 → 地區
    'E': 'E',   # 类别 → 商標類別
    'F': 'F',   # 申请号 → 申請號
    'G': 'G',   # 申请人 → 申請人
    'H': 'H',   # 申请日期 → 申請日期
    'I': 'I',   # 初审公告日期 → 公告日期
    'K': 'J',   # 异议截止日期 → 異議期限
    'M': 'K',   # 商品/服务 → 商品/服務名稱（原文）
}
DB2_MAPPING_NO_REGION = {
    'B': 'B',   # 商标文字 → 他人商標
    'C': 'C',   # 商标图样 → 商標圖樣
    'D': 'E',   # 类别 → 商標類別
    'E': 'F',   # 申请号 → 申請號
    'F': 'G',   # 申请人 → 申請人
    'G': 'H',   # 申请日期 → 申請日期
    'H': 'I',   # 初审公告日期 → 公告日期
    'J': 'J',   # 异议截止日期 → 異議期限
    'L': 'K',   # 商品/服务 → 商品/服務名稱（原文）
}
DB2_IMAGE_SRC_COL = 2


def _detect_db2_has_region(ws):
    """偵測 DB2 工作表是否含有「地區」欄位（在 Row 1 的 D 欄）"""
    headers = _get_row1_headers(ws)
    return '地區' in headers

DB3_MAPPING = {
    'B': 'B', 'C': 'C',
    'D': 'D', 'E': 'E', 'F': 'F', 'G': 'G', 'H': 'H', 'I': 'I', 'J': 'J',
    'K': 'K', 'L': 'L',
}
DB3_IMAGE_SRC_COL = 2

# 合併檔（已合併過的檔案，可以再次與新資料合併）
MERGED_FILE_MAPPING = {
    'B': 'B', 'C': 'C', 'D': 'D', 'E': 'E', 'F': 'F',
    'G': 'G', 'H': 'H', 'I': 'I', 'J': 'J', 'K': 'K', 'L': 'L',
}
MERGED_FILE_IMAGE_SRC_COL = 2
MERGED_FILE_DATA_START = 3  # 合併檔的資料從第 3 列開始

# DB3 需要清理的欄位（去除前綴、轉換日期格式）
MONTH_MAP = {
    'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04',
    'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
    'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12',
}


def _convert_ddmonyyyy(text):
    """DD-MON-YYYY → YYYY-MM-DD，例如 '02-JAN-2026' → '2026-01-02'。
    無法辨識則回傳原文。"""
    m = re.match(r'^(\d{1,2})-([A-Z]{3})-(\d{4})$', text.strip())
    if m:
        day, mon, year = m.groups()
        month_num = MONTH_MAP.get(mon)
        if month_num:
            return f'{year}-{month_num}-{day.zfill(2)}'
    return text


def clean_db3_date(value):
    """將 DB3 日期格式轉為 YYYY-MM-DD，例如 'Reg: 22-AUG-2025' → '2025-08-22'。
    不管前綴是什麼（App: / Reg: / Opp 等）都去除，只保留日期並轉換格式。"""
    if not value or not isinstance(value, str):
        return value
    value = value.strip()
    # 去除任何英文前綴（含冒號和空格），只保留日期部分
    value = re.sub(r'^[A-Za-z]+[:\s]*\s*', '', value)
    return _convert_ddmonyyyy(value)


def clean_db3_opposition(value):
    """清理異議期限（J 欄）：去 Opp 前綴，並將所有 DD-MMM-YYYY 轉為 YYYY-MM-DD。
    保留其餘文字不動。
    例如 'Opp CN : 30-APR-2026\\nCA : 2 months...' → 'CN : 2026-04-30\\nCA : 2 months...'"""
    if not value or not isinstance(value, str):
        return value
    text = value.strip()
    # 去除開頭的 "Opp " 前綴
    text = re.sub(r'^Opp\s+', '', text)
    # 替換所有出現的 DD-MON-YYYY
    def _replace_date(m):
        return _convert_ddmonyyyy(m.group(0))
    text = re.sub(r'\d{1,2}-[A-Z]{3}-\d{4}', _replace_date, text)
    return text


def clean_class_column(value):
    """清理商標類別（E 欄）：去空格、頓號分隔、去前導 0。
    例如 '09 11 43' → '9、11、43'，'9, 11, 43' → '9、11、43'"""
    if not value:
        return value
    text = str(value).strip()
    if not text:
        return text
    # 以逗號、頓號、空白拆分
    parts = re.split(r'[,、\s]+', text)
    # 去前導 0 並過濾空字串
    cleaned = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # 去前導 0（純數字才處理）
        if p.isdigit():
            p = str(int(p))
        cleaned.append(p)
    return '、'.join(cleaned) if cleaned else text


def clean_db3_app_number(value):
    """去除申請號的前綴，例如 'App 26000442' → '26000442'"""
    if not value or not isinstance(value, str):
        return value
    return re.sub(r'^(App|Reg)\s+', '', value.strip())


# DB3 地區名稱對照表
DB3_REGION_MAP = {
    'EU trade marks': 'EUIPO',
    'International Register': 'WIPO',
    'United States of America': 'US (USPTO)',
}


def clean_db3_region(value):
    """轉換 DB3 地區名稱，例如 'EU trade marks' → 'EUIPO'"""
    if not value or not isinstance(value, str):
        return value
    value = value.strip()
    # 先檢查完全匹配
    if value in DB3_REGION_MAP:
        return DB3_REGION_MAP[value]
    # 處理帶括號的情況，例如 "EU trade marks (unpublished applications)" → "EUIPO(unpublished applications)"
    for original, replacement in DB3_REGION_MAP.items():
        if value.startswith(original):
            suffix = value[len(original):]  # 例如 " (unpublished applications)"
            return replacement + suffix.lstrip()  # "EUIPO(unpublished applications)"
    return value

MERGED_HEADERS = [
    '#', '他人商標', '商標圖樣', '地區', '商標類別',
    '申請號', '申請人', '申請日期', '公告日期', '異議期限',
    '商品/服務名稱（原文）', '商品/服務名稱（英文翻譯）',
]
MERGED_IMAGE_COL = 2
MERGED_HEADER_ROW = 3
MERGED_DATA_START = 4
SOURCE_DATA_START = 2


# ============================================================
# 核心函式
# ============================================================
def col_letter_to_index(letter):
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def find_merged_header_row(file_bytes):
    """找出合併檔的標頭所在列（掃描 Row 1~5）"""
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        ws = wb.active
        for check_row in range(1, min(6, ws.max_row + 1)):
            headers = set()
            for cell in ws[check_row]:
                if cell.value is not None:
                    headers.add(str(cell.value).strip())
            if '他人商標' in headers and '商標圖樣' in headers:
                wb.close()
                return check_row
        wb.close()
    except Exception:
        pass
    return None


def _get_row1_headers(ws):
    """取得工作表 Row 1 的所有非空值"""
    headers = set()
    for cell in ws[1]:
        if cell.value is not None:
            headers.add(str(cell.value).strip())
    return headers


def detect_db_type(file_bytes):
    """
    自動辨識檔案來自哪個資料庫。
    回傳 'db1', 'db2', 'db3', 'merged', 或 None（無法辨識）。

    注意：某些 xlsx 在 read_only 模式會讀不到完整資料，
    因此改用一般模式；DB1 (Markify) 的資料可能不在第一個 sheet，
    需要掃描所有 sheet。
    """
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

        # 先檢查所有 sheet（DB1 Markify 資料在非第一個 sheet）
        for ws in wb.worksheets:
            headers_row1 = _get_row1_headers(ws)

            # DB1：英文標頭，含 "Trademark"
            if 'Trademark' in headers_row1:
                wb.close()
                return 'db1'

            # DB2：簡體中文，含 "商标文字"
            if '商标文字' in headers_row1:
                wb.close()
                return 'db2'

            # DB3：繁體中文，含 "他人商標"（Row 1 有 # 和 他人商標）
            if '他人商標' in headers_row1:
                wb.close()
                return 'db3'

        # 都不符合 → 檢查 active sheet 的 Row 2~5（合併檔）
        ws = wb.active
        for check_row in range(2, min(6, ws.max_row + 1)):
            headers_check = set()
            for cell in ws[check_row]:
                if cell.value is not None:
                    headers_check.add(str(cell.value).strip())
            if '他人商標' in headers_check and '商標圖樣' in headers_check:
                wb.close()
                return 'merged'

        wb.close()
        return None
    except Exception:
        return None


def _find_data_sheets(wb, db_type):
    """找出含有資料的 sheet，掃描所有 tab 避免資料不在 active sheet 的問題"""
    # 各 DB 的 Row 1 關鍵字
    db_signature = {
        'db1': 'Trademark',
        'db2': '商标文字',
        'db3': '他人商標',
    }
    keyword = db_signature.get(db_type)
    if keyword:
        sheets = []
        for ws in wb.worksheets:
            headers = _get_row1_headers(ws)
            if keyword in headers:
                sheets.append(ws)
        return sheets if sheets else [wb.active]
    # merged 或其他：使用 active sheet
    return [wb.active]


def read_source_data(file_bytes, mapping, db_type=''):
    """讀取來源 Excel 並按 mapping 轉換欄位"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    rows = []

    sheets = _find_data_sheets(wb, db_type)

    for ws in sheets:
        # DB2：動態偵測該 sheet 是否含有「地區」欄位，選用對應的 mapping
        if db_type == 'db2':
            if _detect_db2_has_region(ws):
                active_mapping = DB2_MAPPING_WITH_REGION
                db2_default_region = None
            else:
                active_mapping = DB2_MAPPING_NO_REGION
                db2_default_region = 'China'  # 摩知輪本身是中國資料庫
        else:
            active_mapping = mapping
            db2_default_region = None

        # 合併檔需要動態找到標頭列，資料從標頭列 +1 開始
        if db_type == 'merged':
            header_row = find_merged_header_row(file_bytes) or 2
            data_start = header_row + 1
        else:
            data_start = SOURCE_DATA_START
        for row_idx, row in enumerate(
            ws.iter_rows(min_row=data_start, values_only=False),
            start=data_start,
        ):
            if row_idx > ws.max_row:
                break
            merged_row = {}
            for src_col_letter, dest_col_letter in active_mapping.items():
                src_idx = col_letter_to_index(src_col_letter)
                if src_idx < len(row):
                    merged_row[dest_col_letter] = row[src_idx].value
            # DB2：無地區欄時預設填入 China（摩知輪 = 中國資料庫）
            if db_type == 'db2' and db2_default_region:
                if not merged_row.get('D') or str(merged_row['D']).strip() == '':
                    merged_row['D'] = db2_default_region
            # DB3 資料清理
            if db_type == 'db3':
                if 'F' in merged_row:
                    merged_row['F'] = clean_db3_app_number(merged_row['F'])
                if 'D' in merged_row:
                    merged_row['D'] = clean_db3_region(merged_row['D'])
                for date_col in ['H', 'I']:
                    if date_col in merged_row:
                        merged_row[date_col] = clean_db3_date(merged_row[date_col])
                if 'J' in merged_row:
                    merged_row['J'] = clean_db3_opposition(merged_row['J'])
                # 空的公告日期 → 1900-01-00、空的異議期限 → 0
                if not merged_row.get('I') or str(merged_row['I']).strip() == '':
                    merged_row['I'] = '1900-01-00'
                if not merged_row.get('J') or str(merged_row['J']).strip() == '':
                    merged_row['J'] = '0'
            if any(v is not None for v in merged_row.values()):
                rows.append(merged_row)
    wb.close()
    return rows


def read_source_images(file_bytes, src_image_col, db_type=''):
    """讀取來源檔案中的圖片"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    images = {}
    sheets = _find_data_sheets(wb, db_type)
    row_offset_accum = 0  # 多 sheet 時累積行數偏移

    for sheet_idx, ws in enumerate(sheets):
        for img in ws._images:
            anchor = img.anchor
            if hasattr(anchor, '_from') and anchor._from.col == src_image_col:
                src_row = anchor._from.row + row_offset_accum
                img_data = BytesIO(img._data())
                images[src_row] = {
                    'data': img_data,
                    'width': img.width,
                    'height': img.height,
                    'from_colOff': anchor._from.colOff,
                    'from_rowOff': anchor._from.rowOff,
                    'to_col': anchor.to.col,
                    'to_row': anchor.to.row,
                    'to_colOff': anchor.to.colOff,
                    'to_rowOff': anchor.to.rowOff,
                    # 保留原始跨距（DB2 圖片會跨到下一列/欄）
                    'row_span': anchor.to.row - anchor._from.row,
                    'col_span': anchor.to.col - anchor._from.col,
                }
        # 累積偏移：當前 sheet 的資料行數（不含標頭）
        if sheet_idx < len(sheets) - 1:
            row_offset_accum += ws.max_row - SOURCE_DATA_START + 1
    wb.close()
    return images


def create_merged_file(all_rows, all_images, progress_bar=None):
    """建立合併後的 Excel 檔"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '商標監控結果清單'

    # 使用 scheme='minor' 讓字型參照 theme 設定（英文 Times New Roman / 中文 新細明體）
    header_font = Font(scheme='minor', bold=True, size=11)
    header_fill = PatternFill('solid', fgColor='D9E1F2')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    data_font = Font(scheme='minor', size=10)
    align_center_top = Alignment(horizontal='center', vertical='top', wrap_text=True)
    align_left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # ── Row 1：LOGO 水平＋垂直置中（A1:L1 合併） ──
    ws.merge_cells('A1:L1')
    ws.row_dimensions[1].height = 58
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    # 插入 LOGO 圖片（精確置中於合併儲存格）
    logo_path = os.path.join(os.path.dirname(__file__), 'logo.jpg')
    if os.path.exists(logo_path):
        logo_img = Image(logo_path)
        logo_img.height = 55
        logo_img.width = int(logo_img.height * (863 / 133))  # 保持範本比例

        # 計算各欄寬度（像素）以求出水平置中偏移
        _cw_chars = [4, 23, 23, 15, 15, 16, 25, 15, 20, 20, 60, 60]  # A~L
        _cw_px = [int(w * 7 + 5) for w in _cw_chars]
        _total_px = sum(_cw_px)
        _left_px = (_total_px - logo_img.width) / 2  # 左留白（像素）

        # 找出 LOGO 起始落在第幾欄、欄內偏移多少
        _cum = 0
        _logo_col, _logo_col_off_px = 0, 0
        for _i, _cpx in enumerate(_cw_px):
            if _cum + _cpx > _left_px:
                _logo_col = _i
                _logo_col_off_px = int(_left_px - _cum)
                break
            _cum += _cpx

        # 垂直置中：row height 58pt → EMU；圖片高度 → EMU
        _row_h_emu = int(58 * 12700)
        _img_h_emu = pixels_to_EMU(logo_img.height)
        _row_off_emu = max(0, (_row_h_emu - _img_h_emu) // 2)

        # 計算 LOGO 結束位置
        _right_px = _left_px + logo_img.width
        _cum2 = 0
        _logo_to_col, _logo_to_col_off_px = len(_cw_px) - 1, _cw_px[-1]
        for _i, _cpx in enumerate(_cw_px):
            if _cum2 + _cpx >= _right_px:
                _logo_to_col = _i
                _logo_to_col_off_px = int(_right_px - _cum2)
                break
            _cum2 += _cpx

        _from_marker = AnchorMarker(
            col=_logo_col,
            colOff=pixels_to_EMU(_logo_col_off_px),
            row=0,
            rowOff=_row_off_emu,
        )
        _to_marker = AnchorMarker(
            col=_logo_to_col,
            colOff=pixels_to_EMU(_logo_to_col_off_px),
            row=0,
            rowOff=_row_off_emu + _img_h_emu,
        )
        logo_img.anchor = TwoCellAnchor(_from=_from_marker, to=_to_marker)
        ws.add_image(logo_img)

    # ── Row 2：監控商標 / 監控地區 / 監控類別（留空讓使用者手動填寫） ──
    ws.merge_cells('A2:C2')
    ws.merge_cells('D2:G2')
    ws.merge_cells('H2:L2')
    ws.row_dimensions[2].height = 38
    info_font = Font(scheme='minor', bold=True, size=14)
    info_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for col_letter, label in [('A', '監控商標：'), ('D', '監控地區：'), ('H', '監控類別：')]:
        ci = col_letter_to_index(col_letter) + 1
        cell = ws.cell(row=2, column=ci, value=label)
        cell.font = info_font
        cell.alignment = info_align

    # ── Row 3：欄位標題 ──
    for col_idx, header in enumerate(MERGED_HEADERS, start=1):
        cell = ws.cell(row=MERGED_HEADER_ROW, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    total = len(all_rows)
    col_letters = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    for i, row_data in enumerate(all_rows):
        row_num = MERGED_DATA_START + i
        num_cell = ws.cell(row=row_num, column=1, value=f'=ROW()-{MERGED_HEADER_ROW}')
        num_cell.font = data_font
        num_cell.alignment = align_center_top
        for cl in col_letters:
            ci = col_letter_to_index(cl) + 1
            v = row_data.get(cl)
            # 申請號（F 欄）：強制存為文字
            if cl == 'F' and v is not None:
                v = str(v).strip()
            # 商標類別（E 欄）：去空格、頓號分隔、去前導 0
            if cl == 'E' and v is not None:
                v = clean_class_column(v)
            cell = ws.cell(row=row_num, column=ci, value=v if v is not None else '')
            cell.font = data_font
            cell.alignment = align_left_top
        if progress_bar and i % 50 == 0:
            progress_bar.progress(
                0.3 + 0.3 * (i / max(total, 1)),
                text=f'寫入資料 {i}/{total}...',
            )

    col_widths = {
        'A': 4, 'B': 23, 'C': 23, 'D': 15, 'E': 15,
        'F': 16, 'G': 25, 'H': 15, 'I': 20, 'J': 20, 'K': 60, 'L': 60,
    }
    for cl, w in col_widths.items():
        ws.column_dimensions[cl].width = w

    for r in range(MERGED_DATA_START, MERGED_DATA_START + total):
        ws.row_dimensions[r].height = 100

    img_total = len(all_images)
    for idx, (img_info, row_offset) in enumerate(all_images):
        try:
            img_info['data'].seek(0)
            img = Image(img_info['data'])
            if img_info['width'] and img_info['height']:
                img.width = img_info['width']
                img.height = img_info['height']
            else:
                img.width = 80
                img.height = 80
            max_size = 100
            if img.width > max_size or img.height > max_size:
                ratio = min(max_size / img.width, max_size / img.height)
                img.width = int(img.width * ratio)
                img.height = int(img.height * ratio)
            new_row = img_info['orig_row'] + row_offset
            new_col = MERGED_IMAGE_COL
            # 圖片頂部至少留 100000 EMU (~7.6pt) 的間距，避免壓到上方框線
            MIN_ROW_OFF = 100000
            orig_row_off = img_info.get('from_rowOff', 0)
            row_off = max(orig_row_off, MIN_ROW_OFF)
            _from = AnchorMarker(
                col=new_col,
                colOff=img_info.get('from_colOff', 0),
                row=new_row,
                rowOff=row_off,
            )
            _to = AnchorMarker(
                col=new_col + img_info.get('col_span', 0),
                colOff=img_info.get('to_colOff', pixels_to_EMU(img.width)),
                row=new_row + img_info.get('row_span', 0),
                rowOff=img_info.get('to_rowOff', pixels_to_EMU(img.height)),
            )
            img.anchor = TwoCellAnchor(_from=_from, to=_to)
            ws.add_image(img)
        except Exception:
            pass
        if progress_bar and idx % 20 == 0:
            progress_bar.progress(
                0.6 + 0.35 * (idx / max(img_total, 1)),
                text=f'寫入圖片 {idx}/{img_total}...',
            )

    ws.freeze_panes = 'A4'
    last_row = MERGED_HEADER_ROW + len(all_rows)
    ws.auto_filter.ref = f'A{MERGED_HEADER_ROW}:L{last_row}'

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Patch theme XML：將 minor font 的 latin 字型改為 Times New Roman
    # 這樣 scheme='minor' 的 cell 會用 Times New Roman 顯示英文，新細明體顯示中文
    from zipfile import ZipFile
    patched = BytesIO()
    with ZipFile(output, 'r') as zin:
        with ZipFile(patched, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'xl/theme/theme1.xml':
                    text = data.decode('utf-8')
                    text = re.sub(
                        r'(<a:minorFont>\s*<a:latin typeface=")[^"]*(")',
                        r'\1Times New Roman\2',
                        text,
                    )
                    data = text.encode('utf-8')
                zout.writestr(item, data)
    patched.seek(0)
    return patched, len(all_rows)


# ============================================================
# Streamlit 介面
# ============================================================
DB_LABELS = {
    'merged': '合併檔',
    'db1': '資料庫 1',
    'db2': '資料庫 2',
    'db3': '資料庫 3',
}

st.title("📋 商標監控資料合併工具")
st.markdown("上傳各資料庫的原始 Excel 檔（最多 15 個），系統會自動辨識來源並合併。也可同時放入舊的合併檔，系統會一起整合。")

st.divider()

# 用遞增 key 來重置 file_uploader
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

# 單一上傳窗口
uploaded_files = st.file_uploader(
    "將檔案拖放至此處，或點擊 Browse files 選擇檔案",
    type=["xlsx"],
    accept_multiple_files=True,
    help="支援同時上傳多個 .xlsx 檔案，系統會自動辨識來自哪個資料庫",
    key=f"file_uploader_{st.session_state.uploader_key}",
)

# 注入 JS 強制顯示所有已上傳檔案（移除 Streamlit 內建的分頁隱藏）
# st_components.html 的 iframe 有 allow-same-origin，可以存取 parent document
if uploaded_files and len(uploaded_files) > 3:
    st_components.html("""<script>
    (function(){
        try {
            var doc = window.parent.document;
            function showAll() {
                var items = doc.querySelectorAll(
                    '[data-testid="stFileUploaderFile"]'
                );
                items.forEach(function(el){ el.style.display = 'flex'; });
                var pag = doc.querySelectorAll(
                    '[data-testid="stFileUploader"] nav[role="navigation"], ' +
                    '[data-testid="stFileUploader"] [data-testid="stPagination"]'
                );
                pag.forEach(function(el){ el.style.display = 'none'; });
            }
            showAll();
            setTimeout(showAll, 500);
            setTimeout(showAll, 1500);
        } catch(e) { /* sandbox 擋住就算了，CSS 會處理分頁隱藏 */ }
    })();
    </script>""", height=0)

# 檢查上傳數量
if uploaded_files and len(uploaded_files) > 15:
    st.error("⚠️ 最多只能上傳 15 個檔案，請減少檔案數量。")
    st.stop()

# 辨識並分類檔案
if uploaded_files:
    classified = {'merged': [], 'db1': [], 'db2': [], 'db3': []}
    unknown_files = []

    for f in uploaded_files:
        file_bytes = f.getvalue()
        db_type = detect_db_type(file_bytes)
        if db_type:
            classified[db_type].append((f.name, file_bytes))
        else:
            unknown_files.append(f.name)

    # 顯示辨識結果
    st.subheader("📂 檔案辨識結果")

    cols = st.columns(4)
    for i, (db_key, label) in enumerate(DB_LABELS.items()):
        with cols[i]:
            files_list = classified[db_key]
            count = len(files_list)
            if count > 0:
                st.success(f"**{label}**　{count} 個檔案")
                for fname, _ in files_list:
                    st.caption(f"　📄 {fname}")
            else:
                if db_key == 'merged':
                    st.info(f"**{label}**　無")
                else:
                    st.warning(f"**{label}**　未偵測到")

    if unknown_files:
        st.error(
            f"⚠️ 以下 {len(unknown_files)} 個檔案無法辨識來源，將被忽略：\n\n"
            + "\n".join(f"- {name}" for name in unknown_files)
        )

    # 確認至少有檔案可以合併
    total_files = sum(len(v) for v in classified.values())
    if total_files == 0:
        st.error("沒有可辨識的檔案，請確認上傳的是正確的原始檔。")
        st.stop()

    st.divider()

    # 初始化 session_state
    if 'merge_done' not in st.session_state:
        st.session_state.merge_done = False

    # 合併按鈕（僅在尚未合併時顯示）
    if not st.session_state.merge_done:
        if st.button("🚀 開始合併", type="primary", use_container_width=True):
            progress_bar = st.progress(0, text="開始處理...")
            logs = []

            try:
                all_rows = []
                all_images = []
                row_counts = {}
                img_counts = {}

                # 定義處理順序和對應的設定
                db_configs = {
                    'merged': {'mapping': MERGED_FILE_MAPPING, 'img_col': MERGED_FILE_IMAGE_SRC_COL},
                    'db1': {'mapping': DB1_MAPPING, 'img_col': DB1_IMAGE_SRC_COL},
                    'db2': {'mapping': DB2_MAPPING_NO_REGION, 'img_col': DB2_IMAGE_SRC_COL},
                    'db3': {'mapping': DB3_MAPPING, 'img_col': DB3_IMAGE_SRC_COL},
                }

                step = 0
                total_steps = total_files * 2  # 每個檔案讀資料 + 讀圖片

                # 按 合併檔 → db1 → db2 → db3 順序處理
                for db_key in ['merged', 'db1', 'db2', 'db3']:
                    files_list = classified[db_key]
                    if not files_list:
                        continue

                    config = db_configs[db_key]
                    label = DB_LABELS[db_key]
                    db_row_count = 0
                    db_img_count = 0

                    for fname, file_bytes in files_list:
                        # 讀取資料
                        step += 1
                        progress_bar.progress(
                            0.05 + 0.20 * (step / total_steps),
                            text=f'讀取 {label}：{fname}...',
                        )
                        rows = read_source_data(file_bytes, config['mapping'], db_type=db_key)
                        logs.append(f"{label} / {fname}：{len(rows)} 筆資料")

                        # 讀取圖片
                        step += 1
                        progress_bar.progress(
                            0.05 + 0.20 * (step / total_steps),
                            text=f'讀取 {label} 圖片：{fname}...',
                        )
                        images = read_source_images(file_bytes, config['img_col'], db_type=db_key)
                        logs.append(f"{label} / {fname}：{len(images)} 張圖片")

                        # 計算圖片位移
                        if db_key == 'merged':
                            src_header_row = find_merged_header_row(file_bytes) or 2
                            src_data_start = src_header_row + 1
                        else:
                            src_data_start = SOURCE_DATA_START
                        row_offset = (MERGED_DATA_START - src_data_start) + len(all_rows)
                        for src_row, img_info in images.items():
                            img_info['orig_row'] = src_row
                            all_images.append((img_info, row_offset))

                        all_rows.extend(rows)
                        db_row_count += len(rows)
                        db_img_count += len(images)

                    row_counts[db_key] = db_row_count
                    img_counts[db_key] = db_img_count

                # 建立合併檔
                progress_bar.progress(0.30, text=f'建立合併檔（{len(all_rows)} 筆，{len(all_images)} 張圖片）...')
                output_bytes, count = create_merged_file(all_rows, all_images, progress_bar)
                progress_bar.progress(1.0, text="合併完成！")

                # 執行記錄彙總
                logs.append("───────────────────")
                for db_key, label in DB_LABELS.items():
                    rc = row_counts.get(db_key, 0)
                    ic = img_counts.get(db_key, 0)
                    fc = len(classified[db_key])
                    if fc > 0:
                        logs.append(f"{label}：{fc} 個檔案 → {rc} 筆 / {ic} 張圖片")
                logs.append(f"合計：{count} 筆 / {len(all_images)} 張圖片")

                # 儲存結果到 session_state，下載後頁面 rerun 時仍可顯示
                st.session_state.merge_done = True
                st.session_state.merge_output = output_bytes.getvalue()
                st.session_state.merge_count = count
                st.session_state.merge_img_count = len(all_images)
                st.session_state.merge_logs = logs
                client_now = _get_client_now()
                st.session_state.merge_filename = f"{client_now.strftime('%Y%m%d_%H%M')}_合併檔.xlsx"
                st.session_state.merge_active_dbs = [
                    (db_key, DB_LABELS[db_key], row_counts.get(db_key, 0), img_counts.get(db_key, 0), len(classified[db_key]))
                    for db_key in DB_LABELS if len(classified[db_key]) > 0
                ]
                st.rerun()

            except Exception as e:
                progress_bar.empty()
                st.error(f"❌ 合併失敗：{str(e)}")
                import traceback
                with st.expander("錯誤詳情"):
                    st.code(traceback.format_exc())

    # 下載後自動重置的 callback（同時寫入合併紀錄到 JSON 檔）
    def _reset_after_download():
        # 從 JSON 檔讀取既有紀錄
        history = _load_merge_history()
        # 記錄本次合併
        record = {
            'time': _get_client_now().strftime('%H:%M'),
            'filename': st.session_state.get('merge_filename', ''),
            'logs': st.session_state.get('merge_logs', []),
            'count': st.session_state.get('merge_count', 0),
            'img_count': st.session_state.get('merge_img_count', 0),
        }
        history.append(record)
        _save_merge_history(history)
        # 清除本次合併的暫存
        for key in ['merge_done', 'merge_output', 'merge_count', 'merge_img_count',
                     'merge_logs', 'merge_filename', 'merge_active_dbs']:
            st.session_state.pop(key, None)
        st.session_state.uploader_key += 1

    # 合併完成後：持久顯示結果（即使下載觸發 rerun 也不會消失）
    if st.session_state.merge_done:
        st.balloons()
        st.success("合併完成！")

        # 顯示各來源統計
        active_dbs = st.session_state.merge_active_dbs
        result_cols = st.columns(len(active_dbs)) if active_dbs else st.columns(1)
        for i, (db_key, label, rc, ic, fc) in enumerate(active_dbs):
            with result_cols[i]:
                st.metric(f"{label}（{fc} 檔）", f"{rc} 筆", f"{ic} 張圖片")

        st.markdown(f"### 合計：{st.session_state.merge_count} 筆資料 / {st.session_state.merge_img_count} 張圖片")

        # 下載按鈕（點擊後自動重置頁面）
        st.download_button(
            label="⬇️ 下載合併檔",
            data=st.session_state.merge_output,
            file_name=st.session_state.merge_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            on_click=_reset_after_download,
        )

        # 執行記錄
        with st.expander("📝 執行記錄"):
            for log in st.session_state.merge_logs:
                st.text(log)

# ============================================================
# 合併紀錄（從 JSON 檔讀取，當日有效，隔日自動清空）
# ============================================================
_persisted_history = _load_merge_history()
if _persisted_history:
    st.divider()
    with st.expander(f"📋 今日合併紀錄（{len(_persisted_history)} 次）", expanded=False):
        for i, rec in enumerate(reversed(_persisted_history), 1):
            st.markdown(f"**{i}. {rec['time']}　→　{rec['filename']}**　"
                        f"（{rec['count']} 筆 / {rec['img_count']} 張圖片）")
            for log_line in rec['logs']:
                if log_line.startswith('──'):
                    break
                st.caption(f"　　{log_line}")
            if i < len(_persisted_history):
                st.markdown("---")

# 頁尾
st.divider()
st.caption("商標監控資料合併工具 · 輸出為簡易版格式（無標題列），可事後手動加上事務所標題。")
st.caption(f"{APP_VERSION} · Last updated {APP_LAST_UPDATE}")
