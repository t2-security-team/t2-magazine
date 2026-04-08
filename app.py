import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# --- [디자인 및 PDF 압축 CSS] ---
st.markdown("""
    <style>
    .main .block-container { padding-top: 0px !important; padding-bottom: 0px !important; margin-top: -15px !important; }
    div[data-testid="stVerticalBlock"] { gap: 0px !important; }
    .merged-table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: center; font-family: sans-serif; }
    .merged-table th { background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; padding: 4px; font-weight: bold; }
    .merged-table td { border: 1px solid #dee2e6 !important; padding: 3px; vertical-align: middle; }
    .sum-cell { background-color: #ffffff !important; font-weight: bold; color: #1E3A8A; font-size: 13px; }
    .total-banner { background-color: #f0f7ff !important; padding: 8px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 5px; }
    .carrier-banner { background-color: #ffffff !important; padding: 4px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 10px; display: flex; justify-content: center; gap: 20px; }
    .print-row { display: flex; flex-direction: row; gap: 15px; width: 100%; align-items: flex-start; }
    .print-col { flex: 1; min-width: 0; }
    </style>
""", unsafe_allow_html=True)

# --- [도구함] ---
def get_expected_time(time_str):
    """도착시간에서 10분을 뺀 예상시간 계산"""
    try:
        t = datetime.strptime(str(time_str), "%H:%M")
        exp = t - timedelta(minutes=10)
        return exp.strftime("%H:%M")
    except:
        return time_str

def clean_flight_no(val):
    if pd.isna(val): return ""
    val = str(val).strip().replace(" ", "").upper()
    match = re.match(r'([A-Z]+)(\d+)', val)
    if match: 
        airline, num = match.group(1), int(match.group(2))
        return f"{airline}{num:03d}"
    return val

def smart_read(file):
    try:
        filename = file.name.lower()
        if filename.endswith('.csv'):
            try: return pd.read_csv(file, encoding='utf-8')
            except: return pd.read_csv(file, encoding='cp949')
        return pd.read_excel(file, engine='openpyxl')
    except: return None

def parse_dl_pax(df):
    if df is None or df.empty: return None
    all_rows = [df.columns.tolist()] + df.values.tolist()
    pax_row_idx = -1
    for i, row in enumerate(all_rows):
        if any('환승객' in str(cell).replace(" ", "") for cell in row):
            pax_row_idx, pax_row_data = i, row
            break
    if pax_row_idx != -1:
        header = all_rows[0]
        dl_data = []
        for col_idx, cell in enumerate(header):
            cell_str = str(cell)
            if 'DL' in cell_str.upper() and re.search(r'DL\s*\d+', cell_str, re.IGNORECASE):
                flt = re.search(r'(DL\s*\d+)', cell_str, re.IGNORECASE).group(1).replace(" ", "").upper()
                if col_idx < len(pax_row_data):
                    try:
                        pax_count = int(float(str(pax_row_data[col_idx]).replace(",", "")))
                        dl_data.append({'편명': clean_flight_no(flt), '승객수': pax_count})
                    except: pass
        return pd.DataFrame(dl_data) if dl_data else None
    return None

def find_col(df, keywords):
    for col in df.columns:
        clean_col = str(col).replace(" ", "").replace("/", "").upper()
        if any(key.upper() in clean_col for key in keywords): return col
    return None

def clean_route(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    match = re.search(r'(.*?)\s*(\([A-Za-z0-9]+\))', val)
    if match: return f"{match.group(1).split('/')[0].strip()}{match.group(2).strip()}"
    return val.split('/')[0].strip()

def generate_table_html(df, title, count, color, opt_airline, opt_peak):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 style='text-align:center; color:{color}; font-size:16px; margin-bottom:5px;'>{display_title}</h3>"
    if df.empty: return html + "<p style='text-align:center;'>데이터 없음</p></div>"
    
    df = df.sort_values('시간').reset_index(drop=True)
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    
    html += '<table class="merged-table"><thead><tr>'
    html += '<th style="width:14%;">예상시간</th><th style="width:12%;">시간</th><th>출발지</th>'
    html += '<th style="width:14%;">편명</th><th style="width:11%;">게이트</th><th style="width:11%;">승객</th>'
    html += '<th style="width:11%;">합계</th></tr></thead><tbody>'
    
    processed_hours = set()
    for i, row in df.iterrows():
        curr_h = row['hour_val']
        flt = str(row['편명']).upper()
        exp_time = get_expected_time(row['시간']) # 예상시간 계산 로직 추가
        
        row_style = ""
        sum_bg = "#ffffff"
        
        # 색상 옵션 적용
        if opt_peak:
            if curr_h == 16: row_style, sum_bg = 'background-color: #F4FAFD;', "#F4FAFD"
            elif curr_h == 17: row_style, sum_bg = 'background-color: #FFFDF0;', "#FFFDF0"
            elif curr_h == 18: row_style, sum_bg = 'background-color: #FFF5F8;', "#FFF5F8"
        if opt_airline:
            if flt.startswith("DL"): row_style = 'background-color: #FFFDE7;'
            elif flt.startswith("OZ"): row_style = 'background-color: #FDF4F7;'

        html += f'<tr style="{row_style}">'
        html += f'<td>{exp_time}</td>' # 예상시간 출력
        html += f'<td>{row["시간"]}</td><td>{row["출발지"]}</td><td>{row["편명"]}</td>'
        html += f'<td>{row["게이트"]}</td><td>{row["p_val"]:,}</td>'
        
        if curr_h not in processed_hours:
            html += f'<td rowspan="{hour_counts[curr_h]}" class="sum-cell" style="background-color: {sum_bg} !important;">{hour_sums[curr_h]:,}</td>'
            processed_hours.add(curr_h)
        html += '</tr>'
    
    return html + '</tbody></table></div>'

# --- [사이드바 및 로직] ---
with st.sidebar:
    st.header("📂 데이터 업로드")
    pax_files = st.file_uploader("1. 승객수 파일", accept_multiple_files=True)
    gate_files = st.file_uploader("2. 게이트 파일", accept_multiple_files=True)
    view_option = st.radio("표시 방식", ("⬜ 기본", "✈️ 항공사별", "⏰ 첨두시간"))
    time_range = st.slider("시간대", 0, 24, (0, 24))

opt_airline, opt_peak = "항공사별" in view_option, "첨두시간" in view_option

if pax_files and gate_files:
    p_all, g_all = [], []
    for f in pax_files:
        df = smart_read(f)
        if df is not None:
            dl = parse_dl_pax(df)
            if dl is not None: p_all.append(dl)
            else:
                f_c, p_c, r_c = find_col(df, ['편명','FLT']), find_col(df, ['TS','PAX','승객']), find_col(df, ['FROM','출발지'])
                if f_c and p_c:
                    tmp = df[[f_c, p_c]].copy()
                    if r_c: tmp['출발지'] = df[r_c].apply(clean_route)
                    tmp.columns = ['편명', '승객수', '출발지'] if r_c else ['편명', '승객수']
                    tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                    p_all.append(tmp)

    for f in gate_files:
        df = smart_read(f)
        if df is not None:
            f_c, g_c, t_c, r_c = find_col(df, ['편명','FLT']), find_col(df, ['GATE','게이트']), find_col(df, ['TIME','시간']), find_col(df, ['FROM','출발지'])
            if f_c and g_c and t_c:
                tmp = df[[f_c, g_c, t_c]].copy()
                if r_c: tmp['출발지'] = df[r_c].apply(clean_route)
                tmp.columns = ['편명', '게이트', '시간', '출발지'] if r_c else ['편명', '게이트', '시간']
                tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                g_all.append(tmp)

    if p_all and g_all:
        df_p = pd.concat(p_all).drop_duplicates('편명')
        df_g = pd.concat(g_all).drop_duplicates('편명')
        final = pd.merge(df_g, df_p, on='편명', how='inner')
        
        if not final.empty:
            final['p_val'] = pd.to_numeric(final['승객수'], errors='coerce').fillna(0).astype(int)
            final['hour'] = final['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
            final = final[(final['hour'] >= time_range[0]) & (final['hour'] <= time_range[1])]
            final['g_num'] = pd.to_numeric(final['게이트'], errors='coerce').fillna(0)
            final['구역'] = final['g_num'].apply(lambda x: '서편' if 0 < x <= 250 else '동편')
            final['게이트'] = final['g_num'].astype(int).astype(str)
            
            # 상단 배너
            total_p = final['p_val'].sum()
            ke_s = final[final['편명'].str.startswith('KE', na=False)]['p_val'].sum()
            oz_s = final[final['편명'].str.startswith('OZ', na=False)]['p_val'].sum()
            dl_s = final[final['편명'].str.startswith('DL', na=False)]['p_val'].sum()

            # 저장 버튼 UI (기존 코드와 동일)
            st.components.v1.html("""
                <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                <button onclick="takePic()" style="padding:8px 15px; cursor:pointer; border-radius:5px; border:1px solid #ddd;">📸 전체 사진으로 저장</button>
                <script>
                function takePic() {
                    const target = window.parent.document.querySelector('.block-container');
                    html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#ffffff' }).then(canvas => {
                        const link = document.createElement('a');
                        link.download = '보안검색_잡지.png';
                        link.href = canvas.toDataURL();
                        link.click();
                    });
                }
                </script>
            """, height=50)

            st.markdown(f'<div class="total-banner"><h3>📊 총 승객수: {total_p:,}명</h3></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="carrier-banner"><span>KE: {ke_s:,}명</span><span>OZ: {oz_s:,}명</span><span>DL: {dl_s:,}명</span></div>', unsafe_allow_html=True)

            w_html = generate_table_html(final[final['구역'] == '서편'], "⬅️ 서편", final[final['구역'] == '서편']['p_val'].sum(), "#DC2626", opt_airline, opt_peak)
            e_html = generate_table_html(final[final['구역'] == '동편'], "➡️ 동편", final[final['구역'] == '동편']['p_val'].sum(), "#2563EB", opt_airline, opt_peak)
            
            st.markdown(f'<div class="print-row">{e_html}{w_html}</div>', unsafe_allow_html=True)
else:
    st.info("파일을 업로드해주세요.")
