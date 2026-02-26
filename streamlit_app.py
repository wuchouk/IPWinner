import re
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.drawing.image import Image
from openpyxl.utils.units import pixels_to_EMU
from io import BytesIO
from datetime import datetime

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="商標監控資料合併工具",
    page_icon="📋",
    layout="centered",
)

# ============================================================
# 欄位對應設定
# ============================================================
DB1_MAPPING = {
    'A': 'B', 'B': 'C', 'C': 'D', 'D': 'E', 'E': 'F',
    'F': 'G', 'G': 'H', 'H': 'I', 'I': 'J', 'J': 'K', 'K': 'L',
}
DB1_IMAGE_SRC_COL = 1

DB2_MAPPING = {
    'B': 'B', 'C': 'C', 'D': 'D', 'E': 'E', 'F': 'F',
    'G': 'G', 'H': 'H', 'I': 'I', 'K': 'J', 'M': 'K',
}
DB2_IMAGE_SRC_COL = 2

DB3_MAPPING = {
    'B': 'B', 'C': 'C',
    'D': 'D', 'E': 'E', 'F': 'F', 'G': 'G', 'H': 'H', 'I': 'I', 'J': 'J',
    'K': 'K', 'L': 'L',
}
DB3_IMAGE_SRC_COL = 2

# DB3 需要清理的欄位（去除前綴、轉換日期格式）
MONTH_MAP = {
    'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04',
    'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
    'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12',
}


def clean_db3_date(value):
    """將 DB3 日期格式轉為 YYYY-MM-DD，例如 '02-JAN-2026' → '2026-01-02'"""
    if not value or not isinstance(value, str):
        return value
    value = value.strip()
    # 去除前綴 "App: " 或 "Opp "
    value = re.sub(r'^(App:\s*|Opp\s+)', '', value)
    # 轉換 DD-MON-YYYY → YYYY-MM-DD
    m = re.match(r'^(\d{1,2})-([A-Z]{3})-(\d{4})$', value)
    if m:
        day, mon, year = m.groups()
        month_num = MONTH_MAP.get(mon)
        if month_num:
            return f'{year}-{month_num}-{day.zfill(2)}'
    return value


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
MERGED_HEADER_ROW = 2
MERGED_DATA_START = 3
SOURCE_DATA_START = 2


# ============================================================
# 核心函式
# ============================================================
def col_letter_to_index(letter):
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def detect_db_type(file_bytes):
    """
    自動辨識檔案來自哪個資料庫。
    回傳 'db1', 'db2', 'db3', 或 None（無法辨識）。
    """
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active
        headers = set()
        for cell in ws[1]:
            if cell.value is not None:
                headers.add(str(cell.value).strip())
        max_col = ws.max_column
        wb.close()

        # DB1：英文標頭，含 "Trademark"
        if 'Trademark' in headers:
            return 'db1'

        # DB2：簡體中文，含 "商标文字"
        if '商标文字' in headers:
            return 'db2'

        # DB3：繁體中文，含 "他人商標"（與 DB2 的 "商标文字" 區分）
        if '他人商標' in headers:
            return 'db3'

        return None
    except Exception:
        return None


def read_source_data(file_bytes, mapping, db_type=''):
    """讀取來源 Excel 並按 mapping 轉換欄位"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=SOURCE_DATA_START, values_only=False),
        start=SOURCE_DATA_START,
    ):
        if row_idx > ws.max_row:
            break
        merged_row = {}
        for src_col_letter, dest_col_letter in mapping.items():
            src_idx = col_letter_to_index(src_col_letter)
            if src_idx < len(row):
                merged_row[dest_col_letter] = row[src_idx].value
        # DB3 資料清理
        if db_type == 'db3':
            if 'F' in merged_row:
                merged_row['F'] = clean_db3_app_number(merged_row['F'])
            if 'D' in merged_row:
                merged_row['D'] = clean_db3_region(merged_row['D'])
            for date_col in ['H', 'I', 'J']:
                if date_col in merged_row:
                    merged_row[date_col] = clean_db3_date(merged_row[date_col])
            # 空的公告日期 → 1900-01-00、空的異議期限 → 0
            if not merged_row.get('I') or str(merged_row['I']).strip() == '':
                merged_row['I'] = '1900-01-00'
            if not merged_row.get('J') or str(merged_row['J']).strip() == '':
                merged_row['J'] = '0'
        if any(v is not None for v in merged_row.values()):
            rows.append(merged_row)
    wb.close()
    return rows


def read_source_images(file_bytes, src_image_col):
    """讀取來源檔案中的圖片"""
    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    ws = wb.active
    images = {}
    for img in ws._images:
        anchor = img.anchor
        if hasattr(anchor, '_from') and anchor._from.col == src_image_col:
            src_row = anchor._from.row
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
            }
    wb.close()
    return images


def create_merged_file(all_rows, all_images, progress_bar=None):
    """建立合併後的 Excel 檔"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '商標監控結果清單'

    header_font = Font(name='Arial', bold=True, size=11)
    header_fill = PatternFill('solid', fgColor='D9E1F2')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    data_font = Font(name='Arial', size=10)

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
        ws.cell(row=row_num, column=1, value=f'=ROW()-{MERGED_HEADER_ROW}')
        for cl in col_letters:
            ci = col_letter_to_index(cl) + 1
            v = row_data.get(cl)
            cell = ws.cell(row=row_num, column=ci, value=v if v is not None else '')
            cell.font = data_font
        if progress_bar and i % 50 == 0:
            progress_bar.progress(
                0.3 + 0.3 * (i / max(total, 1)),
                text=f'寫入資料 {i}/{total}...',
            )

    col_widths = {
        'A': 6, 'B': 30, 'C': 12, 'D': 20, 'E': 12,
        'F': 18, 'G': 30, 'H': 14, 'I': 14, 'J': 14, 'K': 40, 'L': 40,
    }
    for cl, w in col_widths.items():
        ws.column_dimensions[cl].width = w

    for r in range(MERGED_DATA_START, MERGED_DATA_START + total):
        ws.row_dimensions[r].height = 50

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
            _from = AnchorMarker(
                col=new_col,
                colOff=img_info.get('from_colOff', 0),
                row=new_row,
                rowOff=img_info.get('from_rowOff', 0),
            )
            _to = AnchorMarker(
                col=new_col,
                colOff=img_info.get('to_colOff', pixels_to_EMU(img.width)),
                row=new_row,
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

    ws.freeze_panes = 'A3'
    last_row = MERGED_HEADER_ROW + len(all_rows)
    ws.auto_filter.ref = f'A{MERGED_HEADER_ROW}:L{last_row}'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(all_rows)


# ============================================================
# Streamlit 介面
# ============================================================
DB_LABELS = {
    'db1': '資料庫 1',
    'db2': '資料庫 2',
    'db3': '資料庫 3',
}

st.title("📋 商標監控資料合併工具")
st.markdown("上傳各資料庫的原始 Excel 檔（最多 15 個），系統會自動辨識來源並合併。")

st.divider()

# 單一上傳窗口
uploaded_files = st.file_uploader(
    "選擇要合併的 Excel 檔案",
    type=["xlsx"],
    accept_multiple_files=True,
    help="支援同時上傳多個檔案，系統會自動辨識來自哪個資料庫",
)

# 檢查上傳數量
if uploaded_files and len(uploaded_files) > 15:
    st.error("⚠️ 最多只能上傳 15 個檔案，請減少檔案數量。")
    st.stop()

# 辨識並分類檔案
if uploaded_files:
    classified = {'db1': [], 'db2': [], 'db3': []}
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

    cols = st.columns(3)
    for i, (db_key, label) in enumerate(DB_LABELS.items()):
        with cols[i]:
            files_list = classified[db_key]
            count = len(files_list)
            if count > 0:
                st.success(f"**{label}**　{count} 個檔案")
                for fname, _ in files_list:
                    st.caption(f"　📄 {fname}")
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

    # 合併按鈕
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
                'db1': {'mapping': DB1_MAPPING, 'img_col': DB1_IMAGE_SRC_COL},
                'db2': {'mapping': DB2_MAPPING, 'img_col': DB2_IMAGE_SRC_COL},
                'db3': {'mapping': DB3_MAPPING, 'img_col': DB3_IMAGE_SRC_COL},
            }

            step = 0
            total_steps = total_files * 2  # 每個檔案讀資料 + 讀圖片

            # 按 db1 → db2 → db3 順序處理
            for db_key in ['db1', 'db2', 'db3']:
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
                    images = read_source_images(file_bytes, config['img_col'])
                    logs.append(f"{label} / {fname}：{len(images)} 張圖片")

                    # 計算圖片位移
                    row_offset = (MERGED_DATA_START - SOURCE_DATA_START) + len(all_rows)
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

            # 顯示結果
            st.balloons()

            result_cols = st.columns(3)
            for i, (db_key, label) in enumerate(DB_LABELS.items()):
                with result_cols[i]:
                    rc = row_counts.get(db_key, 0)
                    ic = img_counts.get(db_key, 0)
                    fc = len(classified[db_key])
                    st.metric(
                        f"{label}（{fc} 檔）",
                        f"{rc} 筆",
                        f"{ic} 張圖片",
                    )

            st.markdown(f"### 合計：{count} 筆資料 / {len(all_images)} 張圖片")

            # 下載按鈕
            filename = f"合併檔_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇️ 下載合併檔",
                data=output_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            # 執行記錄
            with st.expander("📝 執行記錄"):
                for log in logs:
                    st.text(log)

        except Exception as e:
            progress_bar.empty()
            st.error(f"❌ 合併失敗：{str(e)}")
            import traceback
            with st.expander("錯誤詳情"):
                st.code(traceback.format_exc())

# 頁尾
st.divider()
st.caption("商標監控資料合併工具 · 輸出為簡易版格式（無標題列），可事後手動加上事務所標題。")
