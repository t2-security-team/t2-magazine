import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime, timedelta, timezone

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# ⭐️ [구글 시트 연동 설정]
SHEET_NAME = "보안검색_데이터_공유"

def get_gspread_client():
    creds_dict = dict(st.secrets["gcp"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

# [수정] 데이터와 파일명을 함께 저장
def save_data_and_files(df, file_names):
    client = get_gspread_client()
    sh = client.open(SHEET_NAME)
    
    # 1. 승객 데이터 저장
    try: sheet = sh.worksheet("pax_data")
    except: sheet = sh.add_worksheet(title="pax_data", rows="1000", cols="20")
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.fillna("").astype(str).values.tolist())
    
    # 2. 파일 목록 저장
    try: file_sheet = sh.worksheet("file_list")
    except: file_sheet = sh.add_worksheet(title="file_list", rows="100", cols="1")
    file_sheet.clear()
    file_sheet.update([["파일명"]] + [[name] for name in file_names])

# [수정] 데이터와 파일명을 함께 불러오기
def load_all():
    client = get_gspread_client()
    df = pd.DataFrame()
    files = []
    try:
        sh = client.open(SHEET_NAME)
        # 데이터 로드
        pax_data = sh.worksheet("pax_data").get_all_values()
        if len(pax_data) > 1: df = pd.DataFrame(pax_data[1:], columns=pax_data[0])
        # 파일 목록 로드
        f_list = sh.worksheet("file_list").get_all_values()
        if len(f_list) > 1: files = [row[0] for row in f_list[1:]]
    except: pass
    return df, files

# --- [도구함 함수들: clean_flight_no, smart_read, parse_dl_pax, find_col, format_route, generate_table_html 동일] ---
def clean_flight_no(val):
    if pd.isna(val): return ""
    val = str(val).strip().replace(" ", "").upper()
    match = re.match(r'([A-Z]+)(\d+)', val)
    if match: return f"{match.group(1)}{int(match.group(2)):03d}"
    return val

def smart_read(file):
    try:
        if file.name.endswith('.csv'):
            for enc in ['utf-8', 'cp949', 'euc-kr']:
                try: file.seek(0); return pd.read_csv(file, encoding=enc)
                except: continue
        else: file.seek(0); return pd.read_excel(file)
    except: return None
    return None

def parse_dl_pax(df):
    if df is None or df.empty: return None
    all_rows = [df.columns.tolist()] + df.values.tolist()
    pax_row_idx = -1
    pax_row_data = []
    for i, row in enumerate(all_rows):
        for cell in row:
            if str(cell).replace(" ", "").strip() == '환승객':
                pax_row_idx = i; pax_row_data = row; break
        if pax_row_idx != -1: break
    if pax_row_idx != -1:
        header = all_rows[0]
        dl_data = []
        for col_idx, cell in enumerate(header):
            if 'DL' in str(cell).upper() and re.search(r'DL\s*\d+', str(cell), re.IGNORECASE):
                flt_no = clean_flight_no(re.search(r'(DL\s*\d+)', str(cell), re.IGNORECASE).group(1))
                if col_idx < len(pax_row_data):
                    try: dl_data.append({'편명': flt_no, '승객수': int(float(str(pax_row_data[col_idx]).replace(",", "")))})
                    except: pass
        if dl_data: return pd.DataFrame(dl_data)
    return None

def find_col(df, keywords):
    for col in df.columns:
        clean = str(col).replace(" ", "").upper()
        for key in keywords:
            if key.upper() in clean: return col
    return None

def format_route(val, option):
    if pd.isna(val): return ""
    val = re.sub(r'\([가-힣\s]+\)', '', str(val)).strip()
    match = re.search(r'(.*?)\s*\(([A-Za-z0-9]+)\)', val)
    if match:
        city, code = match.group(1).split('/')[0].strip(), match.group(2).strip().upper()
        if code == "HND": city = "하네다"
        elif code == "NRT": city = "나리타"
        return city if option == "한글 (도시명)" else code if option == "영어 (쓰리코드)" else f"{city}({code})"
    return val

def generate_table_html(df, title, count, color, opt_airline, opt_peak, font_size):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 style='text-align:center; color:{color}; font-size:16px; margin-top:2px; margin-bottom:5px;'>{display_title}</h3>"
    if df.empty: return html + "<div style='text-align:center; padding:20px; border:1px solid #ddd;'>데이터 없음</div></div>"
    df = df.sort_values('시간').reset_index(drop=True)
    html += f'<table class="merged-table" style="font-size: {font_size}px !important;"><thead><tr><th>예상시간</th><th>시간</th><th>편명</th><th>출발지</th><th>게이트</th><th>승객</th><th>합계</th></tr></thead><tbody>'
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    processed_hours = set()
    for i, row in df.iterrows():
        current_h = row['hour_val']
        html += f'<tr><td></td><td>{row["시간"]}</td><td>{row["편명"]}</td><td>{row.get("출발지", "")}</td><td>{row["게이트"]}</td><td>{row["p_display"]}</td>'
        if current_h not in processed_hours:
            html += f'<td rowspan="{hour_counts[current_h]}" class="sum-cell">{hour_sums[current_h]:,}</td>'
            processed_hours.add(current_h)
        html += '</tr>'
    return html + '</tbody></table></div>'

# --- [사이드바] ---
with st.sidebar:
    st.header("📂 데이터 업로드")
    uploaded_files = st.file_uploader("승객수 파일 업로드", accept_multiple_files=True, key="pax_uploader")
    if uploaded_files:
        p_temp = []
        names = []
        for f in uploaded_files:
            df = smart_read(f)
            if df is not None:
                dl = parse_dl_pax(df)
                if dl is not None: p_temp.append(dl)
                else:
                    f_c, p_c, r_c = find_col(df, ['FLT', '편명']), find_col(df, ['TS', 'PAX', '승객수']), find_col(df, ['FROM', '출발지'])
                    if f_c and p_c:
                        tmp = df[[f_c, p_c]].copy()
                        if r_c: tmp['출발지'] = df[r_c].astype(str)
                        tmp.columns = ['편명', '승객수', '출발지'] if r_c else ['편명', '승객수']
                        tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                        p_temp.append(tmp)
                names.append(f.name)
        if p_temp: save_data_and_files(pd.concat(p_temp).drop_duplicates('편명'), names)
        st.rerun()

    saved_df, saved_names = load_all()
    if saved_names:
        st.markdown("<div class='file-box'>", unsafe_allow_html=True)
        st.markdown("<p class='file-box-title'>✅ 현재 보관중인 파일:</p>", unsafe_allow_html=True)
        for name in saved_names: st.markdown(f"<p class='file-item'>• {name}</p>", unsafe_allow_html=True)
        if st.button("🗑️ 전체 데이터 비우기"):
            clear_sheet("pax_data"); clear_sheet("file_list"); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
    # (나머지 게이트 파일 로직 및 UI는 이전과 동일)
