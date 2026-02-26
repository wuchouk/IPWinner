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
# 欄位對應設定（與桌面版相同）
# ============================================================
DB1_MAPPING = {
    'A': 'B', 'B': 'C', 'C': 'D', 'D': 'E', 'E': 'F',
    'F': 'G', 'G': 'H', 'H': 'I', 'I': 'J', 'J': 'K', 'K': 'L',
}
DB1_IMAGE_SRC_COL = 1
DB1_EXPECTED_HEADERS = [
    'Trademark', 'Logotype', 'Databases', 'Classes',
    'Application number', 'Owner/Applicant', 'Application date',
    'Publication date', 'Deadline for opposition',
    'Goods & Services', 'Goods & Services (Translated)',
]

DB2_MAPPING = {
    'B': 'B', 'C': 'C', 'D': 'D', 'E': 'E', 'F': 'F',
    'G': 'G', 'H': 'H', 'I': 'I', 'K': 'J', 'M': 'K',
}
DB2_IMAGE_SRC_COL = 2
DB2_EXPECTED_HEADERS = [
    '序号', '商标文字', '商标图样', '地區', '类别',
    '申请号', '申请人', '申请日期', '初审公告日期',
]

DB3_MAPPING = {
    'B': 'B', 'C': 'C', 'K': 'K', 'L': 'L',
    'N': 'D', 'O': 'E', 'P': 'F', 'Q': 'G', 'R': 'H', 'S': 'I', 'T': 'J',
}
DB3_IMAGE_SRC_COL = 2
DB3_EXPECTED_HEADERS = ['#', '他人商標', '商標圖樣']

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


def read_source_data(file_bytes, mapping, expected_headers, header_row=1, data_start_row=2):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    headers = [cell.value for cell in ws[header_row]]
    matched = sum(
        1 for exp in expected_headers
        if any(h and str(h).strip() == str(exp).strip() for h in headers)
    )
    if matched < min(3, len(expected_headers)):
        wb.close()
        actual = [str(h) for h in headers[:15] if h]
        raise ValueError(
            f"標頭驗證失敗！預期含有 {expected_headers[:5]}，"
            f"實際標頭為 {actual}，僅匹配 {matched} 個"
        )
    rows = []
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=data_start_row, values_only=False), start=data_start_row
    ):
        if row_idx > ws.max_row:
            break
        merged_row = {}
        for src_col_letter, dest_col_letter in mapping.items():
            src_idx = col_letter_to_index(src_col_letter)
            if src_idx < len(row):
                merged_row[dest_col_letter] = row[src_idx].value
        if any(v is not None for v in merged_row.values()):
            rows.append(merged_row)
    wb.close()
    return rows


def read_source_images(file_bytes, src_image_col):
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
            progress_bar.progress(0.3 + 0.3 * (i / max(total, 1)), text=f'寫入資料 {i}/{total}...')

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
st.title("📋 商標監控資料合併工具")
st.markdown("上傳 3 個資料庫的原始 Excel 檔，自動合併為統一格式（含商標圖片）。")

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**資料庫 1**")
    st.caption("Trademark, Logotype, Databases...")
    file_db1 = st.file_uploader("選擇檔案", type=["xlsx"], key="db1")

with col2:
    st.markdown("**資料庫 2**")
    st.caption("序号, 商标文字, 商标图样...")
    file_db2 = st.file_uploader("選擇檔案", type=["xlsx"], key="db2")

with col3:
    st.markdown("**資料庫 3**")
    st.caption("他人商標, 商標圖樣, 地區...")
    file_db3 = st.file_uploader("選擇檔案", type=["xlsx"], key="db3")

st.divider()

# 狀態顯示
if file_db1 and file_db2 and file_db3:
    st.success("✅ 三個檔案已上傳，可以開始合併！")

    if st.button("🚀 開始合併", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="開始處理...")
        status = st.empty()
        logs = []

        try:
            # 讀取原始位元組
            db1_bytes = file_db1.getvalue()
            db2_bytes = file_db2.getvalue()
            db3_bytes = file_db3.getvalue()

            # ---- 資料庫 1 ----
            progress_bar.progress(0.05, text="讀取資料庫 1...")
            db1_rows = read_source_data(db1_bytes, DB1_MAPPING, DB1_EXPECTED_HEADERS)
            logs.append(f"資料庫 1：{len(db1_rows)} 筆資料")

            progress_bar.progress(0.10, text="讀取資料庫 1 圖片...")
            db1_images = read_source_images(db1_bytes, DB1_IMAGE_SRC_COL)
            logs.append(f"資料庫 1：{len(db1_images)} 張圖片")

            # ---- 資料庫 2 ----
            progress_bar.progress(0.15, text="讀取資料庫 2...")
            db2_rows = read_source_data(db2_bytes, DB2_MAPPING, DB2_EXPECTED_HEADERS)
            logs.append(f"資料庫 2：{len(db2_rows)} 筆資料")

            progress_bar.progress(0.20, text="讀取資料庫 2 圖片...")
            db2_images = read_source_images(db2_bytes, DB2_IMAGE_SRC_COL)
            logs.append(f"資料庫 2：{len(db2_images)} 張圖片")

            # ---- 資料庫 3 ----
            progress_bar.progress(0.25, text="讀取資料庫 3...")
            db3_rows = read_source_data(db3_bytes, DB3_MAPPING, DB3_EXPECTED_HEADERS)
            logs.append(f"資料庫 3：{len(db3_rows)} 筆資料")

            progress_bar.progress(0.28, text="讀取資料庫 3 圖片...")
            db3_images = read_source_images(db3_bytes, DB3_IMAGE_SRC_COL)
            logs.append(f"資料庫 3：{len(db3_images)} 張圖片")

            # ---- 組合 ----
            all_rows = db1_rows + db2_rows + db3_rows
            all_images = []

            ro1 = MERGED_DATA_START - SOURCE_DATA_START  # = 1
            for r, info in db1_images.items():
                info['orig_row'] = r
                all_images.append((info, ro1))

            ro2 = ro1 + len(db1_rows)
            for r, info in db2_images.items():
                info['orig_row'] = r
                all_images.append((info, ro2))

            ro3 = ro2 + len(db2_rows)
            for r, info in db3_images.items():
                info['orig_row'] = r
                all_images.append((info, ro3))

            progress_bar.progress(0.30, text=f"建立合併檔（{len(all_rows)} 筆，{len(all_images)} 張圖片）...")

            # ---- 建立合併檔 ----
            output_bytes, count = create_merged_file(all_rows, all_images, progress_bar)

            progress_bar.progress(1.0, text="合併完成！")

            total_imgs = len(db1_images) + len(db2_images) + len(db3_images)
            logs.append(f"───────────────────")
            logs.append(f"合計：{count} 筆 / {total_imgs} 張圖片")

            # 顯示結果
            st.balloons()

            result_cols = st.columns(3)
            with result_cols[0]:
                st.metric("資料庫 1", f"{len(db1_rows)} 筆", f"{len(db1_images)} 張圖片")
            with result_cols[1]:
                st.metric("資料庫 2", f"{len(db2_rows)} 筆", f"{len(db2_images)} 張圖片")
            with result_cols[2]:
                st.metric("資料庫 3", f"{len(db3_rows)} 筆", f"{len(db3_images)} 張圖片")

            st.markdown(f"### 合計：{count} 筆資料 / {total_imgs} 張圖片")

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

else:
    missing = []
    if not file_db1:
        missing.append("資料庫 1")
    if not file_db2:
        missing.append("資料庫 2")
    if not file_db3:
        missing.append("資料庫 3")
    st.info(f"請上傳以下檔案：{'、'.join(missing)}")

# 頁尾
st.divider()
st.caption("商標監控資料合併工具 · 輸出為簡易版格式（無標題列），可事後手動加上事務所標題。")
