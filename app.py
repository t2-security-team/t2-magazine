import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
import io
from datetime import datetime, timedelta, timezone

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

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
    # 오늘보다 이전인 과거 데이터는 몰래 싹 청소해줌 (용량 쾌적)
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
    
    /* ⭐ 수정사항 1번: PDF 인쇄 시 사이드바 및 버튼 완벽 숨김 */
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
     
# --- [사이드바 설정] ---
with st.sidebar:
    st.header("🔗 빠른 사이트 이동")
    st.link_button("✈ 인천공항 도착편 조회", "https://www.airport.kr/ap_ko/872/subview.do", use_container_width=True)
    st.link_button("📧 네이버 메일함 열기", "https://mail.naver.com", use_container_width=True)
    st.link_button("⏪ 이전 버전으로 이동", "https://t2-magazine-old-dby3dpnaxzhq7eoitpqrm7.streamlit.app/", use_container_width=True)
    st.link_button("🔄 실시간 연동 버전으로 이동", "https://live-magazine-t2.streamlit.app/", use_container_width=True)
    st.divider()
    
    st.header("📂 데이터 업로드")
    
    today_ui_str = f"오늘 ({now_kst_time.month}월 {now_kst_time.day}일)"
    tomorrow_ui_str = f"내일 ({(now_kst_time + timedelta(days=1)).month}월 {(now_kst_time + timedelta(days=1)).day}일)"
    
    upload_target = st.radio("📅 업로드할 데이터 날짜", [today_ui_str, tomorrow_ui_str], index=1, horizontal=True)
    target_date_str = today_date_str if "오늘" in upload_target else tomorrow_date_str
    
    full_files_df = load_file_list()
    if not full_files_df.empty:
        saved_files = full_files_df[full_files_df['조회일자'] == target_date_str]['파일명'].tolist()
    else:
        saved_files = []
        
    full_pax_df = load_pax_data()
    if not full_pax_df.empty:
        saved_pax_df = full_pax_df[full_pax_df['조회일자'] == target_date_str]
    else:
        saved_pax_df = pd.DataFrame()
    
    is_upload_locked = len(saved_files) >= 3
    
    if is_upload_locked:
        st.error(f"🚨 **업로드 제한됨**\n\n해당 날짜에 이미 3개의 파일이 등록되어 있습니다. 아래의 데이터 비우기 버튼을 먼저 눌러주세요.")
    
    uploaded_pax_files = st.file_uploader(
        "1. 승객수 파일 (.xls, .xlsx, .csv)", 
        accept_multiple_files=True, 
        key="pax_uploader",
        disabled=is_upload_locked
    )
    
    if uploaded_pax_files and not is_upload_locked:
        if st.button("💾 파일 저장", use_container_width=True):
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
                st.session_state["toast_msg"] = f"{upload_target} 데이터 저장 완료!"
            elif not p_temp:
                st.session_state["toast_msg"] = "⚠ 인식 가능한 데이터를 찾지 못했습니다."
            st.rerun()
     
    if not saved_pax_df.empty:
        st.markdown("<div class='file-box'>", unsafe_allow_html=True)
        st.markdown(f"<p class='file-box-title'>✅ 현재 적용중인 데이터</p>", unsafe_allow_html=True)
        
        if saved_files:
            for fname in saved_files:
                st.markdown(f"<p class='file-item'>• {fname}</p>", unsafe_allow_html=True)
        else:
            st.markdown("<p class='file-item'>• 데이터 적용 완료</p>", unsafe_allow_html=True)
            
        if "오늘" in upload_target:
            if st.button(f"🗑 데이터 비우기", use_container_width=True):
                st.session_state.show_today_warning = True

            if st.session_state.get("show_today_warning", False):
                st.error("🚨 **[경고] 이 데이터를 비우면 오늘 잡지를 볼 수 없습니다!**\n\n진행하시겠습니까?")
                col1, col2 = st.columns(2)
                if col1.button("강제 비우기"):
                    clear_date_data(target_date_str)
                    st.session_state.show_today_warning = False
                    st.session_state["toast_msg"] = "데이터를 모두 비웠습니다."
                    st.rerun()
                if col2.button("취소", type="primary"):
                    st.session_state.show_today_warning = False
                    st.rerun()
        else:
            if st.button(f"🗑 데이터 비우기", use_container_width=True):
                clear_date_data(target_date_str)
                st.session_state.show_today_warning = False
                st.session_state["toast_msg"] = "데이터를 모두 비웠습니다."
                st.rerun()
                
        st.markdown("</div>", unsafe_allow_html=True)
     
    # ⭐ 수정사항 2번: 게이트 업로드 창을 '비상용'으로 접어두기
    with st.expander("🚨 수동 게이트 업로드 (게이트 서버 장애시에만 사용)"):
        gate_files = st.file_uploader("2. 게이트 파일 (.xls, .xlsx, .csv)", accept_multiple_files=True)
    
    st.divider()
    date_option = st.radio("📅 표시 날짜 선택", ["어제 (-1일)", "오늘", "내일 (+1일)"], index=1)
    
    if date_option == "어제 (-1일)": target_date = now_kst_time - timedelta(days=1)
    elif date_option == "내일 (+1일)": target_date = now_kst_time + timedelta(days=1)
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
     
# --- [메인 로직] ---
p_all, g_all = [], []
     
if not saved_pax_df.empty:
    if '출발지' in saved_pax_df.columns:
        saved_pax_df['출발지'] = saved_pax_df['출발지'].apply(lambda x: format_route(x, route_option))
    p_all.append(saved_pax_df)
     
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
     
if not (p_all and g_all):
    st.markdown("<h2 style='text-align: center;'>✈ T2 보안검색 환승부 잡지 ✈</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #4B5563; margin-bottom: 30px;'>👋 환영합니다! 좌측 사이드바에서 데이터를 업로드하시거나, 아래 링크를 통해 원하시는 시스템으로 이동해 주세요.</p>", unsafe_allow_html=True)
    
    with st.expander("📢 시스템 이용 안내", expanded=True):
        st.markdown("""
        * **데이터 업로드**: 좌측 메뉴에서 **'내일'** 날짜를 선택한 후 내일자 승객수 파일을 올려주세요.
        * **실시간 연동**: 내일 데이터를 미리 업로드해 두어도, 자정 전까지는 '실시간 잡지'에서 오늘의 데이터를 정상적으로 확인하실 수 있습니다.
        """)
        
    st.divider()
    st.markdown("### 🔗 빠른 시스템 이동")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("##### 🔄 실시간 잡지")
        st.markdown("<span style='font-size: 13px; color: #6b7280;'>현재 운영 중인 실시간 데이터 연동 버전입니다.</span>", unsafe_allow_html=True)
        st.link_button("이동하기", "https://live-magazine-t2.streamlit.app/", use_container_width=True)
        
    with col2:
        st.markdown("##### 💾 승객 수 파일저장")
        st.markdown("<span style='font-size: 13px; color: #6b7280;'>날짜별 승객수 데이터를 미리 저장하고 관리하는 현재 버전입니다.</span>", unsafe_allow_html=True)
        st.link_button("이동하기", "https://t2-pax-magazine.streamlit.app/", use_container_width=True)
        
    with col3:
        st.markdown("##### ⏪ 초기 버전")
        st.markdown("<span style='font-size: 13px; color: #6b7280;'>과거에 사용하던 구형 잡지 버전입니다.</span>", unsafe_allow_html=True)
        st.link_button("이동하기", "https://t2-magazine-old-dby3dpnaxzhq7eoitpqrm7.streamlit.app/", use_container_width=True)
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
            
            // ⭐ 수정사항 1번: 사진 캡처 시 사이드바 및 버튼 완벽 숨김 처리!
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
