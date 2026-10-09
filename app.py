import html
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
import io
from datetime import datetime, timedelta, timezone

# 1. 페이지 설정
st.set_page_config(page_title="T2 승객 수 파일 저장", layout="wide", initial_sidebar_state="collapsed")

# KST(한국시간) 기준 날짜 세팅
KST = timezone(timedelta(hours=9))
now_kst_time = datetime.now(KST)
today_date_str = now_kst_time.strftime("%Y-%m-%d")
tomorrow_date_str = (now_kst_time + timedelta(days=1)).strftime("%Y-%m-%d")

SHEET_NAME = "보안검색_데이터_공유" 

@st.cache_resource(show_spinner=False)
def get_gspread_client():
    creds_dict = dict(st.secrets["gcp"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_gspread_client()
    return client.open(SHEET_NAME)

# ⭐ [핵심 1] 꼬리표(날짜) 달고 데이터 저장 + 과거 데이터 자동 청소
def update_pax_data(new_df, target_date_str):
    new_df['조회일자'] = target_date_str
    spreadsheet = get_spreadsheet()
    try:
        sheet = spreadsheet.worksheet("pax_data")
        data = sheet.get_all_values()
        if len(data) > 1:
            existing_df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in existing_df.columns:
                existing_df['조회일자'] = today_date_str
        else:
            existing_df = pd.DataFrame(columns=['조회일자', '편명', '승객수', '출발지'])
    except:
        sheet = spreadsheet.add_worksheet(title="pax_data", rows=1000, cols=20)
        existing_df = pd.DataFrame(columns=['조회일자', '편명', '승객수', '출발지'])

    combined = pd.concat([existing_df, new_df], ignore_index=True)
    combined = combined[combined['조회일자'] >= today_date_str]
    combined.drop_duplicates(subset=['조회일자', '편명'], keep='last', inplace=True)

    sheet.clear()
    data_to_save = [combined.columns.values.tolist()] + combined.fillna("").astype(str).values.tolist()
    sheet.update(range_name="A1", values=data_to_save)
    load_pax_data.clear()
    return True

# ⭐ [핵심 2] 파일 목록도 꼬리표 달고 저장
def update_file_list(new_files, target_date_str):
    new_df = pd.DataFrame({'조회일자': [target_date_str]*len(new_files), '파일명': new_files})
    spreadsheet = get_spreadsheet()
    try:
        sheet = spreadsheet.worksheet("file_list")
        data = sheet.get_all_values()
        if len(data) > 1:
            existing_df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in existing_df.columns:
                existing_df['조회일자'] = today_date_str
        else:
            existing_df = pd.DataFrame(columns=['조회일자', '파일명'])
    except:
        sheet = spreadsheet.add_worksheet(title="file_list", rows=100, cols=5)
        existing_df = pd.DataFrame(columns=['조회일자', '파일명'])

    combined = pd.concat([existing_df, new_df], ignore_index=True)
    combined = combined[combined['조회일자'] >= today_date_str]
    combined.drop_duplicates(subset=['조회일자', '파일명'], keep='last', inplace=True)

    sheet.clear()
    data_to_save = [combined.columns.values.tolist()] + combined.fillna("").astype(str).values.tolist()
    sheet.update(range_name="A1", values=data_to_save)
    load_file_list.clear()

@st.cache_data(ttl=1800, max_entries=1, show_spinner=False)
def load_file_list():
    try:
        spreadsheet = get_spreadsheet()
        sheet = spreadsheet.worksheet("file_list")
        data = sheet.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in df.columns: df['조회일자'] = today_date_str
            return df
    except: pass
    return pd.DataFrame()

@st.cache_data(ttl=21600, max_entries=1, show_spinner=False)
def load_pax_data():
    try:
        spreadsheet = get_spreadsheet()
        sheet = spreadsheet.worksheet("pax_data")
        data = sheet.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in df.columns: df['조회일자'] = today_date_str
            return df
    except: pass
    return pd.DataFrame()

# ⭐ 특정 날짜 데이터 비우기 (강제 비우기)
def clear_date_data(target_date_str):
    spreadsheet = get_spreadsheet()
    try:
        sheet = spreadsheet.worksheet("pax_data")
        data = sheet.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in df.columns: df['조회일자'] = today_date_str
            df = df[(df['조회일자'] != target_date_str) & (df['조회일자'] >= today_date_str)]
            sheet.clear()
            sheet.update(range_name="A1", values=[df.columns.values.tolist()] + df.fillna("").astype(str).values.tolist())
    except: pass

    try:
        sheet = spreadsheet.worksheet("file_list")
        data = sheet.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            if '조회일자' not in df.columns: df['조회일자'] = today_date_str
            df = df[(df['조회일자'] != target_date_str) & (df['조회일자'] >= today_date_str)]
            sheet.clear()
            sheet.update(range_name="A1", values=[df.columns.values.tolist()] + df.fillna("").astype(str).values.tolist())
    except: pass

    load_pax_data.clear()
    load_file_list.clear()

if "toast_msg" in st.session_state:
    st.toast(st.session_state["toast_msg"], icon="✅")
    del st.session_state["toast_msg"]
     
# --- [디자인 및 PDF 압축 CSS] ---
st.markdown("""
    <style>
    /* 상단 Fork 버튼과 GitHub 고양이 아이콘 숨김 */
    [data-testid="stToolbarActions"] {
        display: none !important;
    }

    .main .block-container { padding-top: 0px !important; padding-bottom: 0px !important; margin-top: -15px !important; }
    div[data-testid="stVerticalBlock"] { gap: 0px !important; }
    .element-container { margin-bottom: 0px !important; }
    iframe { margin-bottom: 0px !important; min-height: 45px !important; }
    
    .file-box { background-color:#f0f7ff; padding:15px; border-radius:5px; margin-bottom:15px; border: 1px solid #3b82f6; display: block; overflow: visible; }
    .file-item { font-size:13px; margin: 0 0 6px 10px !important; line-height: 1.5 !important; color: #1f2937; font-weight: normal; word-break: break-all; }
    .file-box-title { font-size:14px; font-weight:bold; color:#1E3A8A; margin: 0 0 10px 0 !important; line-height: 1.4 !important; }
    
    .merged-table { width: 100%; border-collapse: collapse; text-align: center; font-family: sans-serif; margin-bottom: 0px !important; }
    .merged-table tr { border: none !important; } 
    .merged-table th { background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; padding: 4px; font-weight: bold; }
    .merged-table td { border: 1px solid #dee2e6 !important; padding: 3px; vertical-align: middle; font-weight: bold !important; }
    
    .sum-cell { font-weight: bold; color: #1E3A8A; vertical-align: middle !important; }
    
    .total-banner { background-color: #f0f7ff !important; padding: 4px 8px !important; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 2px; margin-top: 2px; }
    .carrier-banner { background-color: #ffffff !important; padding: 4px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 4px; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap; }
    .carrier-item { font-size: 14px; font-weight: bold; }
    .print-row { display: flex; flex-direction: row; gap: 15px; width: 100%; }
    .print-col { flex: 1; min-width: 0; margin-bottom: 0px !important; }
    
    @media print {
        .no-print, header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"], iframe, [data-testid="stHtml"] { display: none !important; }
        html, body { height: auto !important; min-height: auto !important; padding-bottom: 0 !important; margin-bottom: 0 !important; padding-top: 0 !important; }
        .appview-container, .main, .block-container, .element-container { padding-top: 0 !important; margin-top: 0 !important; padding-bottom: 0 !important; margin-bottom: 0 !important; }
        div[data-testid="stVerticalBlock"] { gap: 0 !important; }
        body { zoom: 75%; }
        .print-row { display: flex !important; flex-direction: row !important; }
        table { page-break-inside: auto; margin-bottom: 0px !important; }
        tr { page-break-inside: avoid; page-break-after: auto; }
        thead { display: table-header-group; }
        @page { size: A4; margin-top: 12mm !important; margin-bottom: 12mm !important; margin-left: 10mm !important; margin-right: 10mm !important; }
        @page :first { margin-top: 0mm !important; }
    }
    </style>
""", unsafe_allow_html=True)
     
def clean_flight_no(val):
    if pd.isna(val): return ""
    val = str(val).strip().replace(" ", "").upper()
    match = re.match(r'([A-Z]+)(\d+)', val)
    if match: return f"{match.group(1)}{int(match.group(2)):03d}"
    return val
     
def smart_read(file):
    filename = file.name.lower()
    df = None
    try:
        if filename.endswith('.csv'):
            encodings = ['utf-8', 'cp949', 'euc-kr', 'utf-16', 'utf-8-sig']
            for enc in encodings:
                try:
                    file.seek(0)
                    df = pd.read_csv(file, encoding=enc)
                    break
                except: pass
        elif filename.endswith('.xls'):
            try:
                file.seek(0)
                df = pd.read_excel(file, engine='xlrd')
            except:
                try:
                    file.seek(0)
                    raw_data = file.read()
                    for enc in ['cp949', 'euc-kr', 'utf-8']:
                        try:
                            html_str = raw_data.decode(enc)
                            dfs = pd.read_html(io.StringIO(html_str))
                            if dfs: 
                                df = dfs[0]
                                break
                        except: pass
                except: pass
        else:
            file.seek(0)
            df = pd.read_excel(file, engine='openpyxl')
    except:
        try:
            file.seek(0)
            df = pd.read_excel(file)
        except: return None
        
    if df is None or df.empty: return None
    all_data = [df.columns.tolist()] + df.values.tolist()
    header_idx = -1
    for i, row in enumerate(all_data[:20]):
        row_str = "".join([str(x).upper() for x in row])
        if 'FLT' in row_str or '편명' in row_str or 'FLIGHT' in row_str:
            header_idx = i
            break
            
    if header_idx > 0:
        new_header = all_data[header_idx]
        new_data = all_data[header_idx+1:]
        df = pd.DataFrame(new_data, columns=new_header)
        
    df.columns = [str(c) if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    return df
     
def parse_dl_pax(df):
    if df is None or df.empty: return None
    all_rows = [df.columns.tolist()] + df.values.tolist()
    pax_row_idx = -1
    pax_row_data = []
    header_row_data = []
    
    for i, row in enumerate(all_rows):
        for cell in row:
            if str(cell).replace(" ", "").strip() == '환승객':
                pax_row_idx = i
                pax_row_data = row
                break
        if pax_row_idx != -1: break
        
    if pax_row_idx != -1:
        header_row_data = all_rows[0]
        dl_data = []
        for col_idx, cell in enumerate(header_row_data):
            cell_str = str(cell)
            if 'DL' in cell_str.upper() and re.search(r'DL\s*\d+', cell_str, re.IGNORECASE):
                flt_no = re.search(r'(DL\s*\d+)', cell_str, re.IGNORECASE).group(1).replace(" ", "").upper()
                flt_no = clean_flight_no(flt_no) 
                
                if col_idx < len(pax_row_data):
                    pax_val = str(pax_row_data[col_idx]).replace(",", "").strip()
                    try:
                        pax_count = int(float(pax_val))
                        dl_data.append({'편명': flt_no, '승객수': pax_count})
                    except: pass
        if dl_data: return pd.DataFrame(dl_data)
    return None
     
def find_col(df, keywords):
    if df is None or df.empty: return None
    for col in df.columns:
        clean_col = str(col).replace(" ", "").replace("/", "").replace("_", "").replace(".", "").upper()
        for key in keywords:
            if key.upper() in clean_col: return col
    return None
     
def format_route(val, option):
    if pd.isna(val): return ""
    val = str(val).strip()
    val = re.sub(r'\([가-힣\s]+\)', '', val).strip()
    match = re.search(r'(.*?)\s*\(([A-Za-z0-9]+)\)', val)
    
    if match:
        city = match.group(1).split('/')[0].strip() 
        code = match.group(2).strip().upper()       
        if code == "HND": city = "하네다"
        elif code == "NRT": city = "나리타"
            
        if option == "한글 (도시명)": return city
        elif option == "영어 (쓰리코드)": return code
        else: return f"{city}({code})"
            
    if '/' in val: val = val.split('/')[0].strip()
        
    val_upper = val.upper()
    if val_upper == "HND" or "HND" in val_upper:
        if option == "한글 (도시명)": return "하네다"
        elif option == "영어 (쓰리코드)": return "HND"
        else: return "하네다(HND)"
    elif val_upper == "NRT" or "NRT" in val_upper:
        if option == "한글 (도시명)": return "나리타"
        elif option == "영어 (쓰리코드)": return "NRT"
        else: return "나리타(NRT)"
        
    return val
     
def generate_table_html(df, title, count, color, opt_airline, opt_peak, font_size):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 style='text-align:center; color:{color}; font-size:16px; margin-top:2px; margin-bottom:5px;'>{display_title}</h3>"
    if df.empty: return html + "<div style='text-align:center; padding:20px; border:1px solid #ddd;'>데이터 없음</div></div>"
    
    df = df.sort_values('시간').reset_index(drop=True)
    
    html += f'<table class="merged-table" style="font-size: {font_size}px !important;"><thead><tr>'
    html += f'<th style="width:14%; font-size:{font_size}px !important;">예상시간</th>'
    html += f'<th style="width:12%; font-size:{font_size}px !important;">시간</th>'
    html += f'<th style="width:14%; font-size:{font_size}px !important;">편명</th>'
    html += f'<th style="font-size:{font_size}px !important;">출발지</th>'
    html += f'<th style="width:11%; font-size:{font_size}px !important;">게이트</th>'
    html += f'<th style="width:11%; font-size:{font_size}px !important;">승객</th>'
    html += f'<th style="width:11%; font-size:{font_size}px !important;">합계</th>'
    html += f'</tr></thead><tbody>'
    
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    processed_hours = set()
    
    for i, row in df.iterrows():
        current_h = row['hour_val']
        flt = str(row['편명']).upper()
        row_style_css = ""
        
        if opt_airline:
            if flt.startswith("DL"): row_style_css = "background-color: #E3F2FD;" 
            elif flt.startswith("OZ"): row_style_css = "background-color: #FDF4F7;" 
        elif opt_peak:
            if current_h == 16: row_style_css = "background-color: #F4FAFD;" 
            elif current_h == 17: row_style_css = "background-color: #FFFDF0;" 
            elif current_h == 18: row_style_css = "background-color: #FFF5F8;" 
        td_style = f' style="{row_style_css} font-size: {font_size}px !important; font-weight: bold !important;"'
        
        html += f'<tr>'
        html += f'<td{td_style}></td><td{td_style}>{row["시간"]}</td><td{td_style}>{row["편명"]}</td><td{td_style}>{row.get("출발지", "")}</td><td{td_style}>{row["게이트"]}</td><td{td_style}>{row["p_display"]}</td>'
        
        if current_h not in processed_hours:
            sum_font = font_size + 1
            html += f'<td rowspan="{hour_counts[current_h]}" class="sum-cell" style="background-color: #ffffff !important; font-size: {sum_font}px !important; font-weight: bold !important;"><div style="position: relative; z-index: 10;">{hour_sums[current_h]:,}</div></td>'
            processed_hours.add(current_h)
        html += '</tr>'
    return html + '</tbody></table></div>'
     
# --- [날짜별 등록 현황] ---
full_files_df = load_file_list()
full_pax_df = load_pax_data()

tomorrow_kst = now_kst_time + timedelta(days=1)
today_label = f"{now_kst_time.month}월 {now_kst_time.day}일"
tomorrow_label = f"{tomorrow_kst.month}월 {tomorrow_kst.day}일"

def get_date_status(date_str):
    if not full_files_df.empty and '파일명' in full_files_df.columns:
        files = full_files_df[full_files_df['조회일자'] == date_str]['파일명'].tolist()
    else:
        files = []
    if not full_pax_df.empty:
        pax = full_pax_df[full_pax_df['조회일자'] == date_str].copy()
    else:
        pax = pd.DataFrame()
    flights = len(pax)
    total = 0
    if flights and '승객수' in pax.columns:
        total = int(pd.to_numeric(pax['승객수'].astype(str).str.replace(',', '').str.strip(), errors='coerce').fillna(0).sum())
    return files, pax, flights, total

# --- [사이드바: 비상용 잡지 기능] ---
with st.sidebar:
    st.header("🚨 비상용 잡지 보기")
    st.caption("게이트 서버 장애 시에만 사용합니다. 게이트 파일을 올리면 가운데 화면이 잡지 표로 바뀌고, 파일을 지우면 다시 파일 저장 화면으로 돌아옵니다.")
    gate_files = st.file_uploader("게이트 파일 (.xls, .xlsx, .csv)", accept_multiple_files=True, key="gate_uploader")

    st.divider()
    date_option = st.radio("📅 표시 날짜 선택", ["오늘", "내일 (+1일)"], index=0)

    if date_option == "내일 (+1일)": target_date = now_kst_time + timedelta(days=1)
    else: target_date = now_kst_time

    display_date_str = target_date.strftime("%Y년 %m월 %d일")

    st.divider()
    route_option = st.radio("🌍 출발지 표기 방식", ["한글+영어 (혼합)", "한글 (도시명)", "영어 (쓰리코드)"], index=0)
    st.divider()
    vis_option = st.radio("🎨 시각화 옵션", ["적용 안 함", "1. ✈ 항공사별 색상 표시 (DL:연하늘, OZ:연분홍)", "2. ⏰ 첨두시간 색상 표시 (16~18시)"], index=0)
    opt_airline = (vis_option == "1. ✈ 항공사별 색상 표시 (DL:연하늘, OZ:연분홍)")
    opt_peak = (vis_option == "2. ⏰ 첨두시간 색상 표시 (16~18시)")
    st.divider()
    time_range = st.slider("조회 시간대 (시)", 0, 24, (0, 24))
    st.divider()
    base_font_size = st.slider("🔠 표 글자 크기 조절 (px)", min_value=10, max_value=17, value=12, step=1)

st.markdown(f"""
    <style>
    .merged-table, .merged-table th, .merged-table td {{ font-size: {base_font_size}px !important; font-weight: bold !important; }}
    .sum-cell {{ font-size: {base_font_size + 1}px !important; font-weight: bold !important; }}
    </style>
""", unsafe_allow_html=True)

emergency_mode = bool(gate_files)

# --- [메인: 파일 저장 화면] ---
if not emergency_mode:
    st.markdown("""
        <style>
        .pg-title { text-align:center; font-size:30px; font-weight:700; margin:48px 0 4px 0; color:#31333f; }
        .pg-lead { text-align:center; color:#6b7280; font-size:15px; margin:0 0 24px 0; }
        .st-cards { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:8px; }
        .st-card { border:1px solid #e6e6eb; border-radius:10px; padding:14px 16px; background:#fff; }
        .st-card .d { font-weight:700; font-size:16px; color:#31333f; }
        .st-card .tag { display:inline-block; font-size:12px; padding:2px 8px; border-radius:20px; margin-left:6px; font-weight:600; vertical-align:2px; }
        .tag-live { background:#fde8e8; color:#b91c1c; }
        .tag-wait { background:#eef2ff; color:#3730a3; }
        .st-card .n { margin-top:8px; font-size:14px; color:#6b7280; }
        .st-card .n b { color:#1E3A8A; font-size:18px; }
        .st-card .n.empty b { color:#9ca3af; }
        .step-h { font-weight:700; font-size:17px; line-height:1.6; margin:22px 0 0 0; padding-bottom:22px; color:#31333f; display:flex; align-items:center; flex-wrap:wrap; }
        .step-h .num { display:inline-flex; width:24px; height:24px; border-radius:50%; background:#31333f; color:#fff; font-size:13px; align-items:center; justify-content:center; margin-right:8px; flex-shrink:0; }
        .step-h .sub { font-weight:400; font-size:13px; color:#6b7280; margin-left:8px; }
        .note-warn { margin:6px 0 22px 0; background:#ffecec; border-left:4px solid #ff4b4b; padding:10px 14px; border-radius:6px; font-size:14px; color:#7f1d1d; }
        .reg-box { background:#f0f7ff; border:1px solid #3b82f6; border-radius:8px; padding:14px 16px; margin-bottom:8px; }
        .reg-box .h { font-weight:700; color:#1E3A8A; font-size:14px; margin-bottom:8px; }
        .reg-box .f { font-size:13px; margin:0 0 5px 8px; color:#1f2937; word-break:break-all; }
        .reg-box .none { font-size:13px; color:#6b7280; }
        .link-h { font-weight:700; font-size:17px; margin:30px 0 8px 0; padding-top:14px; border-top:1px solid #e6e6eb; color:#31333f; }
        div[data-testid="stFileUploader"] { margin-bottom:10px; }
        @media (max-width: 640px) { .st-cards { grid-template-columns:1fr; } }
        </style>
    """, unsafe_allow_html=True)

    _left, center, _right = st.columns([1, 2.4, 1])
    with center:
        st.markdown("<div class='pg-title'>💾 승객 수 파일 저장</div>", unsafe_allow_html=True)
        st.markdown("<div class='pg-lead'>T2 보안검색 환승부 · 실시간 잡지에 쓰일 승객수 파일을 등록합니다</div>", unsafe_allow_html=True)

        today_files, today_pax, today_flights, today_total = get_date_status(today_date_str)
        tom_files, tom_pax, tom_flights, tom_total = get_date_status(tomorrow_date_str)

        def status_line(files, flights, total):
            if flights == 0:
                return "<div class='n empty'><b>아직 등록된 파일 없음</b></div>"
            return f"<div class='n'>파일 <b>{len(files)}</b>개 · <b>{flights:,}</b>편 · 총 <b>{total:,}</b>명</div>"

        st.markdown(f"""
            <div class='st-cards'>
                <div class='st-card'>
                    <div class='d'>오늘 {today_label} <span class='tag tag-live'>실시간 잡지 표시 중</span></div>
                    {status_line(today_files, today_flights, today_total)}
                </div>
                <div class='st-card'>
                    <div class='d'>내일 {tomorrow_label} <span class='tag tag-wait'>등록 대기</span></div>
                    {status_line(tom_files, tom_flights, tom_total)}
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 1. 날짜 선택
        st.markdown("<div class='step-h'><span class='num'>1</span>어느 날짜 데이터를 올리나요?</div>", unsafe_allow_html=True)
        upload_target = st.radio(
            "업로드할 데이터 날짜",
            [f"내일 ({tomorrow_label})", f"오늘 ({today_label})"],
            index=0, horizontal=True, label_visibility="collapsed", key="upload_target"
        )
        is_today = upload_target.startswith("오늘")
        target_date_str = today_date_str if is_today else tomorrow_date_str
        target_word = "오늘" if is_today else "내일"
        target_label = today_label if is_today else tomorrow_label

        if is_today:
            st.markdown("<div class='note-warn'>⚠ <b>오늘 데이터는 지금 실시간 잡지에 표시 중입니다.</b> 내일 파일을 여기에 올리면 실시간 잡지 승객수가 바뀝니다.</div>", unsafe_allow_html=True)

        saved_files = today_files if is_today else tom_files
        saved_pax_df = today_pax if is_today else tom_pax
        is_upload_locked = len(saved_files) >= 3

        # 2. 파일 올리기
        st.markdown("<div class='step-h'><span class='num'>2</span>승객수 파일 올리기<span class='sub'>.xls · .xlsx · .csv · 여러 개 가능 (날짜당 최대 3개)</span></div>", unsafe_allow_html=True)
        if is_upload_locked:
            st.error(f"🚨 **업로드 제한됨**\n\n{target_word}({target_label})에 이미 3개의 파일이 등록되어 있습니다. 아래 3번에서 데이터를 먼저 비워주세요.")

        uploaded_pax_files = st.file_uploader(
            "승객수 파일 (.xls, .xlsx, .csv)",
            accept_multiple_files=True,
            key="pax_uploader",
            disabled=is_upload_locked,
            label_visibility="collapsed"
        )

        save_clicked = st.button(
            f"💾 {target_word}({target_label}) 데이터로 저장",
            type="primary",
            use_container_width=True,
            disabled=(not uploaded_pax_files) or is_upload_locked
        )

        if save_clicked and uploaded_pax_files and not is_upload_locked:
            with st.spinner(f"📤 파일을 처리하고 저장하는 중..."):
                p_temp = []
                new_file_names = []
                for f in uploaded_pax_files:
                    df = smart_read(f)
                    if df is not None:
                        dl_df = parse_dl_pax(df)
                        if dl_df is not None:
                            p_temp.append(dl_df)
                            new_file_names.append(f.name)
                        else:
                            f_c = find_col(df, ['FLT', '편명', 'FLIGHT'])
                            p_c = find_col(df, ['TS', 'PAX', '승객수', 'T/S', 'TTL', 'TOTAL'])
                            r_c = find_col(df, ['FROM', 'ROUTE', '출발지'])
                            if f_c and p_c:
                                tmp = df[[f_c, p_c]].copy()
                                if r_c: tmp['출발지'] = df[r_c].astype(str)
                                tmp.columns = ['편명', '승객수', '출발지'] if r_c else ['편명', '승객수']
                                tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                                p_temp.append(tmp)
                                new_file_names.append(f.name)

                upload_ok = False
                if p_temp:
                    combined_df = pd.concat(p_temp).drop_duplicates('편명')
                    upload_ok = update_pax_data(combined_df, target_date_str)
                    if upload_ok:
                        update_file_list(new_file_names, target_date_str)

            if upload_ok:
                st.session_state["toast_msg"] = f"{target_word}({target_label}) 데이터 저장 완료!"
            elif not p_temp:
                st.session_state["toast_msg"] = "⚠ 인식 가능한 데이터를 찾지 못했습니다."
            st.rerun()

        # 3. 등록된 파일 확인
        st.markdown("<div class='step-h'><span class='num'>3</span>등록된 파일 확인</div>", unsafe_allow_html=True)
        if not saved_pax_df.empty:
            if saved_files:
                file_lines = "".join([f"<div class='f'>• {html.escape(str(fname))}</div>" for fname in saved_files])
            else:
                file_lines = "<div class='f'>• 데이터 적용 완료</div>"
            st.markdown(f"<div class='reg-box'><div class='h'>📂 {target_word} {target_label} 등록 파일</div>{file_lines}</div>", unsafe_allow_html=True)

            # 오늘 데이터는 실시간 잡지에 표시 중이므로 비밀번호 잠금, 내일 데이터는 바로 비우기
            if is_today:
                with st.expander("🚨 오늘 데이터 강제 비우기 (관리자용)"):
                    st.markdown("<span style='font-size:12px; color:gray;'>실시간 잡지 표출에 문제가 생길 수 있으므로 가급적 지우지 마세요.</span>", unsafe_allow_html=True)

                    admin_pw = st.text_input("비밀번호 입력", type="password", placeholder="비밀번호 4자리")

                    # 🔑 관리자 비밀번호 (현재는 "0000")
                    if admin_pw == "0000":
                        if st.button("🗑 강제 비우기 실행", use_container_width=True, type="primary"):
                            clear_date_data(target_date_str)
                            st.session_state["toast_msg"] = "오늘 데이터를 강제로 비웠습니다."
                            st.rerun()
                    elif admin_pw != "":
                        st.error("비밀번호가 일치하지 않습니다.")
            else:
                if st.button(f"🗑 내일({target_label}) 데이터 비우기", use_container_width=True):
                    clear_date_data(target_date_str)
                    st.session_state["toast_msg"] = "데이터를 모두 비웠습니다."
                    st.rerun()
        else:
            st.markdown(f"<div class='reg-box'><div class='h'>📂 {target_word} {target_label} 등록 파일</div><div class='none'>아직 등록된 파일이 없습니다.</div></div>", unsafe_allow_html=True)

        # 바로가기
        st.markdown("<div class='link-h'>🔗 바로가기</div>", unsafe_allow_html=True)
        l1, l2 = st.columns(2)
        with l1: st.link_button("🔄 실시간 잡지", "https://live-magazine-t2.streamlit.app/", use_container_width=True)
        with l2: st.link_button("✈ 인천공항 도착편", "https://www.airport.kr/ap_ko/872/subview.do", use_container_width=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        l3, l4 = st.columns(2)
        with l3: st.link_button("📧 네이버 메일", "https://mail.naver.com", use_container_width=True)
        with l4: st.link_button("⏪ 이전 버전", "https://t2-magazine-old-dby3dpnaxzhq7eoitpqrm7.streamlit.app/", use_container_width=True)

# --- [메인: 비상용 잡지 화면 (게이트 파일을 올렸을 때만)] ---
else:
    p_all, g_all = [], []

    # 사이드바에서 고른 '표시 날짜'의 승객수 데이터를 사용
    display_date_key = target_date.strftime("%Y-%m-%d")
    if not full_pax_df.empty:
        disp_pax_df = full_pax_df[full_pax_df['조회일자'] == display_date_key].copy()
    else:
        disp_pax_df = pd.DataFrame()

    if not disp_pax_df.empty:
        if '출발지' in disp_pax_df.columns:
            disp_pax_df['출발지'] = disp_pax_df['출발지'].apply(lambda x: format_route(x, route_option))
        p_all.append(disp_pax_df)

    for f in gate_files:
        df = smart_read(f)
        if df is not None:
            f_c = find_col(df, ['FLT', '편명', 'FLIGHT'])
            g_c = find_col(df, ['GN', 'GATE', '게이트', 'G/N'])
            t_c = find_col(df, ['TIME', 'STA', '시간'])
            r_c = find_col(df, ['FROM', 'ROUTE', '출발지'])
            e_c = find_col(df, ['출구', '입국장', 'EXIT'])

            if f_c and g_c and t_c:
                cols_to_extract = [f_c, g_c, t_c]
                col_names = ['편명', '게이트', '시간']

                if r_c:
                    cols_to_extract.append(r_c)
                    col_names.append('출발지')
                if e_c:
                    cols_to_extract.append(e_c)
                    col_names.append('출구')

                tmp = df[cols_to_extract].copy()
                tmp.columns = col_names

                if r_c: tmp['출발지'] = tmp['출발지'].apply(lambda x: format_route(x, route_option))
                tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                g_all.append(tmp)

    if not p_all:
        st.warning(f"⚠ {display_date_str} 승객수 데이터가 없습니다. 왼쪽 메뉴의 '표시 날짜'를 확인하거나, 게이트 파일을 지우고 파일 저장 화면에서 승객수 파일을 먼저 등록해 주세요.")
    elif not g_all:
        st.warning("⚠ 게이트 파일에서 편명·게이트·시간 정보를 찾지 못했습니다. 파일을 확인해 주세요.")
    else:
        df_p = pd.concat(p_all).drop_duplicates('편명')
        df_g = pd.concat(g_all).drop_duplicates('편명')
        final = pd.merge(df_g, df_p, on='편명', how='inner', suffixes=('', '_p'))
    
        if '출발지' in final.columns:
            final = final[~final['출발지'].astype(str).str.contains('PUS|김해|부산', case=False, na=False)]
    
        if not final.empty:
            final['p_val'] = pd.to_numeric(final['승객수'], errors='coerce').fillna(0).astype(int)
        
            def format_pax_display(val):
                if pd.isna(val) or str(val).strip() == '': return ""
                try:
                    cleaned_val = str(val).replace(',', '').strip()
                    if cleaned_val == '': return ""
                    return f"{int(float(cleaned_val)):,}"
                except: return ""
                
            final['p_display'] = final['승객수'].apply(format_pax_display)
            final['hour'] = final['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
            final = final[(final['hour'] >= time_range[0]) & (final['hour'] <= time_range[1])]
        
            if '출구' not in final.columns: final['출구'] = ""
            final['g_num'] = pd.to_numeric(final['게이트'], errors='coerce').fillna(0)
        
            def get_zone(row):
                if row['g_num'] > 0:
                    return '서편' if 0 < row['g_num'] <= 250 else '동편'
                else:
                    exit_val = str(row.get('출구', '')).strip().upper()
                    if exit_val == 'A': return '서편'
                    if exit_val == 'B': return '동편'
                    return '동편'
            def get_gate_str(row):
                if row['g_num'] > 0:
                    return str(int(row['g_num']))
                else:
                    exit_val = str(row.get('출구', '')).strip().upper()
                    if exit_val in ['A', 'B']: return '-'
                    return '-'
        
            final['구역'] = final.apply(get_zone, axis=1)
            final['게이트'] = final.apply(get_gate_str, axis=1)
        
            total_p = final['p_val'].sum()
            def c_sum(c): return final[final['편명'].str.startswith(c, na=False)]['p_val'].sum()
            ke_s, oz_s, dl_s = c_sum('KE'), c_sum('OZ'), c_sum('DL')
        
            st.components.v1.html(
                """
                <style>
                body { margin: 0; padding: 0; overflow: hidden; display: flex; gap: 10px; }
                .custom-btn {
                    background-color: white; border: 1px solid #dcdcdc; color: #31333f;
                    padding: 6px 15px; font-size: 14px; border-radius: 6px; cursor: pointer;
                    font-family: sans-serif; box-shadow: 0px 1px 3px rgba(0,0,0,0.1);
                }
                .custom-btn:hover { border-color: #ff4b4b; color: #ff4b4b; }
                </style>
                <button class="custom-btn" onclick="window.parent.print()">📄 PDF 저장</button>
                <button class="custom-btn" onclick="takePic()" id="pic-btn">📸 전체 사진으로 저장</button>
            
                <script>
                function takePic() {
                    var btn = document.getElementById('pic-btn');
                    btn.innerText = "⏳ 캡처 중... 잠시만요!";
                    try {
                        var win = window.parent;
                        var doc = win.document;
                        if (!win.html2canvas) {
                            var script = doc.createElement('script');
                            script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
                            script.onload = function() { doCap(win, doc, btn); };
                            script.onerror = function() { alert("⚠ 에러"); btn.innerText = "📸 전체 사진으로 저장"; };
                            doc.head.appendChild(script);
                        } else { doCap(win, doc, btn); }
                    } catch(e) { btn.innerText = "📸 전체 사진으로 저장"; }
                }
            
                function doCap(win, doc, btn) {
                    var target = doc.querySelector('.block-container') || doc.querySelector('.main');
                    var hides = doc.querySelectorAll('[data-testid="stSidebar"], header, iframe, [data-testid="stHtml"]');
                    var appView = doc.querySelector('.appview-container') || doc.querySelector('[data-testid="stAppViewContainer"]');
                    var mainView = doc.querySelector('.main');
                
                    var oldAppOverflow = appView ? appView.style.overflow : '';
                    var oldAppHeight = appView ? appView.style.height : '';
                    var oldMainOverflow = mainView ? mainView.style.overflow : '';
                    var oldMainHeight = mainView ? mainView.style.height : '';
                    if(appView) { appView.style.overflow = 'visible'; appView.style.height = 'auto'; }
                    if(mainView) { mainView.style.overflow = 'visible'; mainView.style.height = 'auto'; }
                
                    hides.forEach(function(e){ e.dataset.old = e.style.display; e.style.display = 'none'; });
                    setTimeout(function() {
                        win.html2canvas(target, { scale: 6, useCORS: true, backgroundColor: '#ffffff' }).then(function(canvas) {
                            var link = doc.createElement('a'); link.download = '잡지.png'; link.href = canvas.toDataURL('image/png'); link.click();
                        }).finally(function() {
                            if(appView) { appView.style.overflow = oldAppOverflow; appView.style.height = oldAppHeight; }
                            if(mainView) { mainView.style.overflow = oldMainOverflow; mainView.style.height = oldMainHeight; }
                            hides.forEach(function(e){ e.style.display = e.dataset.old || ''; }); btn.innerText = "📸 전체 사진으로 저장";
                        });
                    }, 800);
                }
                </script>
                """, height=45
            )
        
            st.markdown(f"""
                <div class="total-banner" style="position: relative;">
                    <div style='margin:0; color:#1E3A8A; font-size: 18px; font-weight: bold;'>📊 총 승객수: {total_p:,}명</div>
                    <div style="position: absolute; right: 15px; top: 50%; transform: translateY(-50%); font-weight: bold; color: #1E3A8A; font-size: 16px;">{display_date_str}</div>
                </div>
                <div class="carrier-banner">
                    <span class="carrier-item">KE: <span style="color:#1E3A8A;">{ke_s:,}</span>명</span>
                    <span class="carrier-item">OZ: <span style="color:#1E3A8A;">{oz_s:,}</span>명</span>
                    <span class="carrier-item">DL: <span style="color:#1E3A8A;">{dl_s:,}</span>명</span>
                </div>
                <hr style="margin: 2px 0 10px 0; border: 0; border-top: 1px solid #e5e7eb;">
            """, unsafe_allow_html=True)
        
            west_p = final[final['구역'] == '서편']['p_val'].sum()
            east_p = final[final['구역'] == '동편']['p_val'].sum()
        
            w_html = generate_table_html(final[final['구역'] == '서편'], "⬅ 서편", west_p, "#DC2626", opt_airline, opt_peak, base_font_size)
            e_html = generate_table_html(final[final['구역'] == '동편'], "➡ 동편", east_p, "#2563EB", opt_airline, opt_peak, base_font_size)
            st.markdown(f'<div class="print-row">{e_html}{w_html}</div>', unsafe_allow_html=True)
