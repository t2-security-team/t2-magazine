import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime, timedelta, timezone

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# ⭐ [구글 시트 연동 설정]
SHEET_NAME = "보안검색_데이터_공유" # 만드신 구글 시트 이름

# ⭐ [속도 개선] 인증 클라이언트와 스프레드시트 연결을 캐싱해서, 앱이 켜져있는 동안
# 구글 드라이브에서 이름으로 시트를 매번 새로 찾는 무거운 API 호출을 하지 않도록 함
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

def save_to_sheet(df, sheet_name):
    try:
        spreadsheet = get_spreadsheet()
        try:
            sheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
        sheet.clear()
        data_to_save = [df.columns.values.tolist()] + df.fillna("").astype(str).values.tolist()
        sheet.update(range_name="A1", values=data_to_save)
        load_from_sheet.clear() # ⭐ 방금 쓴 내용을 다음 화면에서 바로 보이도록 캐시 무효화
        return True
    except Exception as e:
        st.sidebar.error(f"⚠ 데이터 저장 실패: {e}")
        return False

# [추가] 업로드된 파일 이름 목록을 보관하는 함수
def append_file_names(new_names):
    if not new_names: return
    try:
        spreadsheet = get_spreadsheet()
        try:
            sheet = spreadsheet.worksheet("file_list")
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title="file_list", rows="100", cols="1")

        existing_list = load_file_names()

        combined = list(set(existing_list + new_names))
        sheet.clear()
        df = pd.DataFrame(combined, columns=["파일명"])
        data_to_save = [df.columns.values.tolist()] + df.values.tolist()
        sheet.update(range_name="A1", values=data_to_save)
        load_file_names.clear() # ⭐ 방금 쓴 내용을 다음 화면에서 바로 보이도록 캐시 무효화
    except Exception as e:
        st.sidebar.error(f"⚠ 파일 목록 저장 실패: {e}")

# [추가] 파일 이름 목록을 불러오는 함수
# ⭐ [속도/할당량 개선] 20초 동안은 캐시된 값을 재사용해서, 슬라이더/라디오 조작만으로 구글시트 읽기 API가
# 계속 호출되어 분당 요청 한도(429 Quota exceeded)에 걸리는 것을 방지
@st.cache_data(ttl=20, show_spinner=False)
def load_file_names():
    try:
        spreadsheet = get_spreadsheet()
        sheet = spreadsheet.worksheet("file_list")
        data = sheet.get_all_values()
        if len(data) > 1:
            return [row[0] for row in data[1:] if row and row[0].strip() != ""]
    except gspread.exceptions.WorksheetNotFound:
        pass
    except Exception as e:
        st.sidebar.error(f"⚠ 파일 목록 불러오기 실패: {e}")
    return []

@st.cache_data(ttl=20, show_spinner=False)
def load_from_sheet(sheet_name):
    try:
        spreadsheet = get_spreadsheet()
        sheet = spreadsheet.worksheet(sheet_name)
        data = sheet.get_all_values()
        if len(data) > 1:
            return pd.DataFrame(data[1:], columns=data[0])
    except gspread.exceptions.WorksheetNotFound:
        pass
    except Exception as e:
        st.sidebar.error(f"⚠ 데이터 불러오기 실패: {e}")
    return pd.DataFrame()

def clear_sheet(sheet_name):
    try:
        spreadsheet = get_spreadsheet()
        sheet = spreadsheet.worksheet(sheet_name)
        sheet.clear()
        load_from_sheet.clear() # ⭐ 비운 내용이 바로 반영되도록 캐시 무효화
        load_file_names.clear()
    except gspread.exceptions.WorksheetNotFound:
        pass
    except Exception as e:
        st.sidebar.error(f"⚠ 데이터 비우기 실패: {e}")

# --- [업로드 완료 알림창 (토스트)] ---
# rerun 직후 화면 상단에서 알림을 띄우기 위해 session_state에 메시지를 저장해두었다가 표시
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
    
    /* ⭐ [수정] 파일 목록 박스 짤림 방지 및 여백 강화 */
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
        .no-print, header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"], iframe { display: none !important; }
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

# --- [도구함] ---
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
                    dfs = pd.read_html(file, encoding='utf-8')
                    if dfs: df = dfs[0]
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
    st.divider()
    
    st.header("📂 데이터 업로드")
    
    # [업로드 및 구글 시트 저장 로직]
    uploaded_pax_files = st.file_uploader("1. 승객수 파일 (.xls, .xlsx, .csv)", accept_multiple_files=True, key="pax_uploader")
    
    if uploaded_pax_files:
        with st.spinner("📤 업로드한 파일을 처리하고 저장하는 중..."):
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
                            if r_c: tmp['출발지'] = df[r_c].astype(str) # 원본 텍스트 그대로 저장
                            tmp.columns = ['편명', '승객수', '출발지'] if r_c else ['편명', '승객수']
                            tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                            p_temp.append(tmp)
                            new_file_names.append(f.name)
            upload_ok = False
            if p_temp:
                combined_df = pd.concat(p_temp).drop_duplicates('편명')
                upload_ok = save_to_sheet(combined_df, "pax_data")
                if upload_ok:
                    append_file_names(new_file_names) # ⭐ 업로드한 파일 이름도 시트에 별도 보관

        # ⭐ 다음 새로고침 때 알림창(토스트)을 띄우기 위해 메시지 저장
        if upload_ok:
            st.session_state["toast_msg"] = f"{len(new_file_names)}개 파일 업로드 완료!"
        elif not p_temp:
            st.session_state["toast_msg"] = "⚠ 인식 가능한 데이터를 찾지 못했습니다."
        st.rerun()

    # 구글 시트에 데이터 및 파일 목록이 있는지 확인
    saved_pax_df = load_from_sheet("pax_data")
    saved_files = load_file_names()
    
    if not saved_pax_df.empty:
        st.markdown("<div class='file-box'>", unsafe_allow_html=True)
        st.markdown("<p class='file-box-title'>✅ 현재 공유중인 승객 데이터</p>", unsafe_allow_html=True)
        
        # ⭐ 보관된 파일 목록 출력
        if saved_files:
            for fname in saved_files:
                st.markdown(f"<p class='file-item'>• {fname}</p>", unsafe_allow_html=True)
        else:
            st.markdown("<p class='file-item'>• 데이터 적용 완료</p>", unsafe_allow_html=True)
            
        if st.button("🗑 전체 데이터 비우기", use_container_width=True):
            clear_sheet("pax_data")
            clear_sheet("file_list") # ⭐ 비울 때 파일 목록도 같이 비우기
            st.session_state["toast_msg"] = "데이터를 모두 비웠습니다."
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    gate_files = st.file_uploader("2. 게이트 파일 (.xls, .xlsx, .csv)", accept_multiple_files=True)
    
    st.divider()
    date_option = st.radio("📅 표시 날짜 선택", ["어제 (-1일)", "오늘", "내일 (+1일)"], index=1)
    
    KST = timezone(timedelta(hours=9))
    today_date = datetime.now(KST)
    if date_option == "어제 (-1일)": target_date = today_date - timedelta(days=1)
    elif date_option == "내일 (+1일)": target_date = today_date + timedelta(days=1)
    else: target_date = today_date
        
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

# 구글 시트에 저장된 데이터를 불러옵니다.
if not saved_pax_df.empty:
    if '출발지' in saved_pax_df.columns:
        # 불러온 데이터에 출발지 표기 방식 옵션 실시간 적용
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
    with st.expander("💡 홈페이지 이용 방법 (필독)", expanded=True):
        st.markdown("""
        ### 🌐 데이터 공유 방식 안내
        * **자동 공유:** 1번째 파일(승객수 파일)을 업로드하면 서버에 보관되며, **모든 팀원이 동일한 데이터를 볼 수 있습니다.**
        * **비우기 버튼:** 다음 날 데이터를 넣기 전, 사이드바의 **[🗑 전체 데이터 비우기]** 버튼을 누르면 서버 데이터가 초기화됩니다.
        """)
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
                        script.onerror = function() { 
                            alert("⚠ 라이브러리를 불러올 수 없습니다."); 
                            btn.innerText = "📸 전체 사진으로 저장"; 
                        };
                        doc.head.appendChild(script);
                    } else {
                        doCap(win, doc, btn);
                    }
                } catch(e) {
                    alert("⚠ 브라우저 보안 설정으로 인해 캡처가 차단되었습니다.");
                    btn.innerText = "📸 전체 사진으로 저장";
                }
            }
            
            function doCap(win, doc, btn) {
                var target = doc.querySelector('.block-container') || doc.querySelector('.main');
                var hides = doc.querySelectorAll('[data-testid="stSidebar"], header, iframe');
                
                var appView = doc.querySelector('.appview-container') || doc.querySelector('[data-testid="stAppViewContainer"]');
                var mainView = doc.querySelector('.main');
                
                var oldAppOverflow = appView ? appView.style.overflow : '';
                var oldAppHeight = appView ? appView.style.height : '';
                var oldMainOverflow = mainView ? mainView.style.overflow : '';
                var oldMainHeight = mainView ? mainView.style.height : '';

                var oldTargetPaddingTop = target.style.paddingTop;
                var oldTargetMarginTop = target.style.marginTop;
                var oldTargetWidth = target.style.width;
                var oldTargetMaxWidth = target.style.maxWidth;

                if(appView) { appView.style.overflow = 'visible'; appView.style.height = 'auto'; }
                if(mainView) { mainView.style.overflow = 'visible'; mainView.style.height = 'auto'; }

                target.style.paddingTop = '10px';
                target.style.marginTop = '0px';
                target.style.width = '1100px'; 
                target.style.maxWidth = '1100px';

                hides.forEach(function(e){ e.dataset.old = e.style.display; e.style.display = 'none'; });
                
                setTimeout(function() {
                    win.html2canvas(target, { 
                        scale: 6, 
                        useCORS: true, 
                        backgroundColor: '#ffffff'
                    }).then(function(canvas) {
                        var link = doc.createElement('a');
                        link.download = '보안검색_잡지_전체.png';
                        link.href = canvas.toDataURL('image/png');
                        link.click();
                    }).catch(function(err) {
                        alert("사진 생성 중 오류가 발생했습니다.");
                    }).finally(function() {
                        if(appView) { appView.style.overflow = oldAppOverflow; appView.style.height = oldAppHeight; }
                        if(mainView) { mainView.style.overflow = oldMainOverflow; mainView.style.height = oldMainHeight; }
                        
                        target.style.paddingTop = oldTargetPaddingTop;
                        target.style.marginTop = oldTargetMarginTop;
                        target.style.width = oldTargetWidth;
                        target.style.maxWidth = oldTargetMaxWidth;

                        hides.forEach(function(e){ e.style.display = e.dataset.old || ''; });
                        btn.innerText = "📸 전체 사진으로 저장";
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
