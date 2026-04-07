import streamlit as st
import pandas as pd
import re

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# --- [디자인 및 PDF 압축 CSS] ---
st.markdown("""
    <style>
    /* 1. 웹 화면 상단 여백 최소화 */
    .main .block-container {
        padding-top: 1rem !important; 
        padding-bottom: 1rem !important;
    }
    iframe {
        margin-bottom: 5px !important;
        min-height: 45px !important;
    }

    /* 2. 표 스타일 압축 (폰트를 모든 기기에서 잘 보이도록 수정) */
    .merged-table { width: 100%; border-collapse: collapse; font-size: 11px; text-align: center; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }
    .merged-table th { background-color: #f8f9fa !important; border: 1px solid #dee2e6; padding: 4px; font-weight: bold; }
    .merged-table td { border: 1px solid #dee2e6; padding: 3px; vertical-align: middle; }
    .sum-cell { background-color: #fdfdfd !important; font-weight: bold; color: #1E3A8A; font-size: 12px; }
    
    /* 3. 요약 배너 스타일 */
    .total-banner { background-color: #f0f7ff !important; padding: 10px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 5px; }
    .carrier-banner { background-color: #ffffff !important; padding: 5px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 10px; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap; }
    .carrier-item { font-size: 14px; font-weight: bold; }

    /* 4. 서편/동편 가로 배치를 위한 레이아웃 */
    .print-row { display: flex; flex-direction: row; gap: 15px; width: 100%; }
    .print-col { flex: 1; min-width: 0; }

    /* 5. 인쇄(PDF) 화면 디테일 여백 설정 */
    @media print {
        /* 불필요한 요소 및 숨겨진 꼬리말(Footer) 완벽 제거 */
        .no-print, header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"], iframe { 
            display: none !important; 
            height: 0 !important; 
            margin: 0 !important; 
            padding: 0 !important;
        }
        
        /* ✅ [추가] 불필요한 빈 페이지(3페이지)가 생기지 않도록 높이 제한 해제 및 하단 여백 제거 */
        html, body {
            height: auto !important;
            min-height: auto !important;
            padding-bottom: 0 !important;
            margin-bottom: 0 !important;
        }
        
        .appview-container, .main, .block-container, .element-container { 
            padding-top: 0 !important; 
            margin-top: 0 !important; 
            padding-bottom: 0 !important; /* 하단 여백 파괴 */
            margin-bottom: 0 !important;
        }
        
        div[data-testid="stVerticalBlock"] { gap: 0 !important; }
        
        body { zoom: 75%; }
        .print-row { display: flex !important; flex-direction: row !important; }
        
        /* 표가 다음 장으로 넘어갈 때 줄이 반으로 잘리지 않도록 보호 */
        table { page-break-inside: auto; }
        tr { page-break-inside: avoid; page-break-after: auto; }
        thead { display: table-header-group; }

        /* 기본 페이지 여백 (2페이지 이상) */
        @page { 
            size: A4; 
            margin-top: 12mm !important; 
            margin-bottom: 10mm !important;
            margin-left: 10mm !important;
            margin-right: 10mm !important;
        }
        
        /* 1페이지 전용 여백 (맨 위 바짝 붙임) */
        @page :first {
            margin-top: 0mm !important; 
        }
    }
    </style>
    """, unsafe_allow_html=True)

# --- [도구함] ---
def smart_read(file):
    try:
        filename = file.name.lower()
        if filename.endswith('.csv'):
            try: return pd.read_csv(file, encoding='utf-8')
            except: return pd.read_csv(file, encoding='cp949')
        return pd.read_excel(file, engine='openpyxl')
    except:
        try: return pd.read_excel(file)
        except: return None

def find_col(df, keywords):
    if df is None or df.empty: return None
    for col in df.columns:
        clean_col = str(col).replace(" ", "").replace("/", "").replace("_", "").replace(".", "").upper()
        for key in keywords:
            if key.upper() in clean_col: return col
    return None

def clean_flight_no(val):
    if pd.isna(val): return ""
    val = str(val).strip().replace(" ", "").upper()
    match = re.match(r'([A-Z]+)(\d+)', val)
    if match:
        return f"{match.group(1)}{int(match.group(2))}"
    return val

def clean_route(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    match = re.search(r'(.*?)\s*(\([A-Za-z0-9]+\))', val)
    if match:
        city = match.group(1).split('/')[0].strip()
        code = match.group(2).strip()
        return f"{city}{code}"
    else:
        if '/' in val:
            return val.split('/')[0].strip()
        return val

def generate_table_html(df, title, count, color):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 style='text-align:center; color:{color}; font-size:16px; margin-top:5px; margin-bottom:5px;'>{display_title}</h3>"
    if df.empty:
        html += "<div style='text-align:center; padding:20px; border:1px solid #ddd;'>데이터 없음</div></div>"
        return html
    
    df = df.sort_values('시간').reset_index(drop=True)
    html += '<table class="merged-table"><thead><tr>'
    
    html += '<th style="width:14%;">예상시간</th>'
    html += '<th style="width:12%;">시간</th>'
    html += '<th>출발지</th>' 
    html += '<th style="width:14%;">편명</th>'
    html += '<th style="width:11%;">게이트</th>'
    html += '<th style="width:11%;">승객</th>'
    html += '<th style="width:11%;">합계</th>'
    
    html += '</tr></thead><tbody>'
    
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    processed_hours = set()
    
    for i, row in df.iterrows():
        html += f'<tr><td></td><td>{row["시간"]}</td><td>{row["출발지"]}</td><td>{row["편명"]}</td>'
        html += f'<td>{row["게이트"]}</td><td>{row["p_val"]:,}</td>'
        curr_h = row['hour_val']
        if curr_h not in processed_hours:
            html += f'<td rowspan="{hour_counts[curr_h]}" class="sum-cell">{hour_sums[curr_h]:,}</td>'
            processed_hours.add(curr_h)
        html += '</tr>'
    return html + '</tbody></table></div>'

# --- [사이드바 설정] ---
with st.sidebar:
    st.header("📂 데이터 업로드")
    pax_files = st.file_uploader("1. 승객수 파일 (.xlsx, .csv)", accept_multiple_files=True)
    gate_files = st.file_uploader("2. 게이트 파일 (.xlsx, .csv)", accept_multiple_files=True)
    st.divider()
    # ✅ 수정됨: 기본값을 0~24시로 꽉 차게 고정
    time_range = st.slider("조회 시간대 (시)", 0, 24, (0, 24))

# --- [메인 로직] ---
if not (pax_files and gate_files):
    st.markdown("<h2 style='text-align: center;'>✈️ T2 보안검색 환승부 잡지 ✈️</h2>", unsafe_allow_html=True)
    with st.expander("💡 홈페이지 이용 방법 및 주의사항 (필독)", expanded=True):
        st.markdown("""
        ### 1. 파일 업로드 방법
        * **1번째 파일 업로드 (승객수 파일):** 이메일로 받은 승객수(T/S, Pax) 데이터 업로드
        * **2번째 파일 업로드 (게이트 파일):** 인천공항 게이트 및 도착시간 데이터 업로드
        * **-인천공항 도착편 날짜, 시간대 설정 후 총 3개 다운로드 (대한항공, 아시아나, 델타)**

        ### 2. 중요: 파일 형식 필수 변환
        * 본 시스템은 **.xlsx 형식만 지원**합니다.
        * **인천공항 도착편** 다운로드 파일은 그대로 올리면 읽히지 않습니다.
        * **방법:** 파일을 열어 **[다른 이름으로 저장]** → 파일 형식을 **[Excel 통합 문서 (*.xlsx)]**로 선택하여 저장 후 업로드하세요.

        ### 3. 기타 안내사항
        * **델타 이메일 :** 승객수가 사진으로 왔을 때에는 이메일로 받은 대한항공 잡지 밑에 직접 입력해주세요.
        """)
else:
    p_all, g_all = [], []
    
    for f in pax_files:
        df = smart_read(f)
        if df is not None:
            f_c = find_col(df, ['FLT', '편명', 'FLIGHT'])
            p_c = find_col(df, ['TS', 'PAX', '승객수', 'T/S'])
            r_c = find_col(df, ['FROM', 'ROUTE', '출발지'])
            if f_c and p_c:
                tmp = df[[f_c, p_c]].copy()
                if r_c: tmp['출발지'] = df[r_c].apply(clean_route)
                tmp.columns = ['편명', '승객수', '출발지'] if r_c else ['편명', '승객수']
                tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                p_all.append(tmp)

    for f in gate_files:
        df = smart_read(f)
        if df is not None:
            f_c = find_col(df, ['FLT', '편명', 'FLIGHT'])
            g_c = find_col(df, ['GN', 'GATE', '게이트', 'G/N'])
            t_c = find_col(df, ['TIME', 'STA', '시간'])
            r_c = find_col(df, ['FROM', 'ROUTE', '출발지'])
            if f_c and g_c and t_c:
                tmp = df[[f_c, g_c, t_c]].copy()
                if r_c: tmp['출발지'] = df[r_c].apply(clean_route)
                tmp.columns = ['편명', '게이트', '시간', '출발지'] if r_c else ['편명', '게이트', '시간']
                tmp['편명'] = tmp['편명'].apply(clean_flight_no)
                g_all.append(tmp)

    if p_all and g_all:
        df_p = pd.concat(p_all).drop_duplicates('편명')
        df_g = pd.concat(g_all).drop_duplicates('편명')
        final = pd.merge(df_g, df_p, on='편명', how='inner', suffixes=('', '_p'))
        
        if not final.empty:
            final['p_val'] = pd.to_numeric(final['승객수'], errors='coerce').fillna(0).astype(int)
            final['hour'] = final['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
            final = final[(final['hour'] >= time_range[0]) & (final['hour'] <= time_range[1])]
            final['g_num'] = pd.to_numeric(final['게이트'], errors='coerce').fillna(0)
            final['구역'] = final['g_num'].apply(lambda x: '서편' if 0 < x <= 250 else '동편')
            final['게이트'] = final['g_num'].astype(int).astype(str)
            
            total_p = final['p_val'].sum()
            def c_sum(c): return final[final['편명'].str.startswith(c, na=False)]['p_val'].sum()
            ke_s, oz_s, dl_s = c_sum('KE'), c_sum('OZ'), c_sum('DL')

            # --- 결과 출력 ---
            st.components.v1.html(
                """
                <style>
                body { margin: 0; padding: 0; overflow: hidden; }
                .custom-print-btn {
                    background-color: white;
                    border: 1px solid #dcdcdc;
                    color: #31333f;
                    padding: 6px 15px;
                    font-size: 14px;
                    border-radius: 6px;
                    cursor: pointer;
                    transition: all 0.2s;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                    box-shadow: 0px 1px 3px rgba(0,0,0,0.1);
                    margin-bottom: 5px;
                }
                .custom-print-btn:hover {
                    border-color: #ff4b4b;
                    color: #ff4b4b;
                }
                </style>
                <button class="custom-print-btn" onclick="window.parent.print()">📄 PDF 저장</button>
                """,
                height=45
            )
            
            st.markdown(f"""
                <div class="total-banner"><h3 style='margin:0; color:#1E3A8A;'>📊 총 승객수: {total_p:,}명</h3></div>
                <div class="carrier-banner">
                    <span class="carrier-item">KE: <span style="color:#1E3A8A;">{ke_s:,}</span>명</span>
                    <span class="carrier-item">OZ: <span style="color:#1E3A8A;">{oz_s:,}</span>명</span>
                    <span class="carrier-item">DL: <span style="color:#1E3A8A;">{dl_s:,}</span>명</span>
                </div>
            """, unsafe_allow_html=True)

            st.divider()
            
            # 서편/동편 병렬 배치
            west_p = final[final['구역'] == '서편']['p_val'].sum()
            east_p = final[final['구역'] == '동편']['p_val'].sum()
            w_html = generate_table_html(final[final['구역'] == '서편'], "⬅️ 서편", west_p, "#DC2626")
            e_html = generate_table_html(final[final['구역'] == '동편'], "➡️ 동편", east_p, "#2563EB")
            
            st.markdown(f'<div class="print-row">{w_html}{e_html}</div>', unsafe_allow_html=True)
