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

# 데이터와 파일명을 함께 저장
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

# 데이터와 파일명을 함께 불러오기
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

def clear_sheet(sheet_name):
    client = get_gspread_client()
    try:
        sheet = client.open(SHEET_NAME).worksheet(sheet_name)
        sheet.clear()
    except: pass

# --- [도구함] ---
def clean_flight_no(val):
    if pd.isna(val): return ""
    val = str(val).strip().replace(" ", "").upper()
    match = re.match(r'([A-Z]+)(\d+)', val)
    if match: return f"{match.group(1)}{int(match.group(2)):03d}"
    return val

def smart_read(file):
    filename = file.name.lower()
    try:
        if filename.endswith('.csv'):
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
    st.header("🔗 빠른 사이트 이동")
    st.link_button("✈ 인천공항 도착편 조회", "https://www.airport.kr/ap_ko/872/subview.do", use_container_width=True)
    st.link_button("📧 네이버 메일함 열기", "https://mail.naver.com", use_container_width=True)
    st.divider()
    
    st.header("📂 데이터 업로드")
    
    uploaded_files = st.file_uploader("1. 승객수 파일 (.xls, .xlsx, .csv)", accept_multiple_files=True, key="pax_uploader")
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
        st.markdown("<p class='file-box-title'>✅ 현재 공유중인 승객 데이터</p>", unsafe_allow_html=True)
        for name in saved_names: st.markdown(f"<p class='file-item'>• {name}</p>", unsafe_allow_html=True)
        if st.button("🗑 전체 데이터 비우기"):
            clear_sheet("pax_data"); clear_sheet("file_list"); st.rerun()
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

# --- [메인 로직] ---
p_all, g_all = [], []
if not saved_df.empty:
    saved_df['출발지'] = saved_df['출발지'].apply(lambda x: format_route(x, route_option))
    p_all.append(saved_df)

for f in gate_files:
    df = smart_read(f)
    if df is not None:
        f_c, g_c, t_c = find_col(df, ['FLT', '편명']), find_col(df, ['GN', 'GATE', '게이트']), find_col(df, ['TIME', 'STA', '시간'])
        if f_c and g_c and t_c:
            tmp = df[[f_c, g_c, t_c]].copy()
            tmp.columns = ['편명', '게이트', '시간']
            tmp['편명'] = tmp['편명'].apply(clean_flight_no)
            g_all.append(tmp)

if not (p_all and g_all):
    st.markdown("<h2 style='text-align: center;'>✈ T2 보안검색 환승부 잡지 ✈</h2>", unsafe_allow_html=True)
else:
    df_p = pd.concat(p_all).drop_duplicates('편명')
    df_g = pd.concat(g_all).drop_duplicates('편명')
    final = pd.merge(df_g, df_p, on='편명', how='inner')
    
    final['p_val'] = pd.to_numeric(final['승객수'], errors='coerce').fillna(0).astype(int)
    final['p_display'] = final['승객수'].apply(lambda x: f"{int(float(str(x).replace(',',''))):,}" if str(x).strip() else "")
    final['hour'] = final['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    final = final[(final['hour'] >= time_range[0]) & (final['hour'] <= time_range[1])]
    
    final['g_num'] = pd.to_numeric(final['게이트'], errors='coerce').fillna(0)
    final['구역'] = final['g_num'].apply(lambda x: '서편' if 0 < x <= 250 else '동편')
    
    total_p = final['p_val'].sum()
    ke_s, oz_s, dl_s = final[final['편명'].str.startswith('KE')]['p_val'].sum(), final[final['편명'].str.startswith('OZ')]['p_val'].sum(), final[final['편명'].str.startswith('DL')]['p_val'].sum()

    st.markdown(f"""
        <div class="total-banner">📊 총 승객수: {total_p:,}명</div>
        <div class="carrier-banner">
            <span class="carrier-item">KE: {ke_s:,}명</span>
            <span class="carrier-item">OZ: {oz_s:,}명</span>
            <span class="carrier-item">DL: {dl_s:,}명</span>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: st.markdown(generate_table_html(final[final['구역'] == '동편'], "➡ 동편", final[final['구역'] == '동편']['p_val'].sum(), "#2563EB", opt_airline, opt_peak, base_font_size), unsafe_allow_html=True)
    with col2: st.markdown(generate_table_html(final[final['구역'] == '서편'], "⬅ 서편", final[final['구역'] == '서편']['p_val'].sum(), "#DC2626", opt_airline, opt_peak, base_font_size), unsafe_allow_html=True)
