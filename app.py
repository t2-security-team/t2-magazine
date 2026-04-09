import streamlit as st
import pandas as pd
import re

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# --- [디자인 및 PDF 압축 CSS (기본 고정 스타일)] ---
st.markdown("""
    <style>
    /* 웹 화면 및 캡처 시 상단 및 요소 간 기본 공백 극한으로 제거 */
    .main .block-container { padding-top: 0px !important; padding-bottom: 0px !important; margin-top: -15px !important; }
    div[data-testid="stVerticalBlock"] { gap: 0px !important; }
    .element-container { margin-bottom: 0px !important; }
    iframe { margin-bottom: 0px !important; min-height: 45px !important; }
    
    /* 표 기본 스타일 (글자 크기는 아래 동적 CSS에서 덮어씌움) */
    .merged-table { width: 100%; border-collapse: collapse; text-align: center; font-family: sans-serif; margin-bottom: 0px !important; }
    .merged-table tr { border: none !important; } 
    .merged-table th { background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; padding: 4px; font-weight: bold; }
    .merged-table td { border: 1px solid #dee2e6 !important; padding: 3px; vertical-align: middle; }
    
    /* 합계 셀 기본 스타일 */
    .sum-cell { background-color: #ffffff !important; font-weight: bold; color: #1E3A8A; vertical-align: middle !important; }
    
    /* 배너 여백 최소화 */
    .total-banner { background-color: #f0f7ff !important; padding: 8px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 2px; margin-top: 2px; }
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
    if match: 
        airline = match.group(1)
        num = int(match.group(2))
        return f"{airline}{num:03d}"
    return val

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
        if dl_data:
            return pd.DataFrame(dl_data)
    return None

def find_col(df, keywords):
    if df is None or df.empty: return None
    for col in df.columns:
        clean_col = str(col).replace(" ", "").replace("/", "").replace("_", "").replace(".", "").upper()
        for key in keywords:
            if key.upper() in clean_col: return col
    return None

def clean_route(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    match = re.search(r'(.*?)\s*(\([A-Za-z0-9]+\))', val)
    if match: return f"{match.group(1).split('/')[0].strip()}{match.group(2).strip()}"
    if '/' in val: return val.split('/')[0].strip()
    return val

def generate_table_html(df, title, count, color, opt_airline, opt_peak):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 style='text-align:center; color:{color}; font-size:16px; margin-top:2px; margin-bottom:5px;'>{display_title}</h3>"
    if df.empty: return html + "<div style='text-align:center; padding:20px; border:1px solid #ddd;'>데이터 없음</div></div>"
    
    df = df.sort_values('시간').reset_index(drop=True)
    html += '<table class="merged-table"><thead><tr><th style="width:14%;">예상시간</th><th style="width:12%;">시간</th><th>출발지</th><th style="width:14%;">편명</th><th style="width:11%;">게이트</th><th style="width:11%;">승객</th><th style="width:11%;">합계</th></tr></thead><tbody>'
    
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    processed_hours = set()
    
    for i, row in df.iterrows():
        curr_h = row['hour_val']
        flt = str(row['편명']).upper()
        
        row_style_css = ""
        
        # 색상 옵션
        if opt_airline:
            if flt.startswith("DL"):
                row_style_css = "background-color: #E3F2FD;" # 아주 연한 하늘색으로 변경됨
            elif flt.startswith("OZ"):
                row_style_css = "background-color: #FDF4F7;" # 아주 연한 분홍색
            
        elif opt_peak:
            if curr_h == 16:
                row_style_css = "background-color: #F4FAFD;" 
            elif curr_h == 17:
                row_style_css = "background-color: #FFFDF0;" 
            elif curr_h == 18:
                row_style_css = "background-color: #FFF5F8;" 

        td_style = f' style="{row_style_css}"' if row_style_css else ""

        html += f'<tr>'
        html += f'<td{td_style}></td><td{td_style}>{row["시간"]}</td><td{td_style}>{row["출발지"]}</td><td{td_style}>{row["편명"]}</td><td{td_style}>{row["게이트"]}</td><td{td_style}>{row["p_val"]:,}</td>'
        
        if curr_h not in processed_hours:
            html += f'<td rowspan="{hour_counts[curr_h]}" class="sum-cell"><div style="position: relative; z-index: 10;">{hour_sums[curr_h]:,}</div></td>'
            processed_hours.add(curr_h)
        html += '</tr>'
    return html + '</tbody></table></div>'

# --- [사이드바 설정] ---
with st.sidebar:
    st.header("📂 데이터 업로드")
    pax_files = st.file_uploader("1. 승객수 파일 (.xlsx, .csv)", accept_multiple_files=True)
    gate_files = st.file_uploader("2. 게이트 파일 (.xlsx, .csv)", accept_multiple_files=True)
    
    st.divider()
    st.markdown("### 🎨 시각화 옵션 (체크박스)")
    opt_airline = st.checkbox("1. ✈️ 항공사별 색상 표시 (DL:연하늘, OZ:연분홍)")
    opt_peak = st.checkbox("2. ⏰ 첨두시간 색상 표시 (16~18시)")
    
    st.divider()
    time_range = st.slider("조회 시간대 (시)", 0, 24, (0, 24))
    
    # 🔠 글자 크기 조절 슬라이더
    st.divider()
    st.markdown("### 🔠 글자 크기 조절")
    # 기본값 12px, 최소 8px, 최대 24px
    base_font_size = st.slider("표 글자 크기 (px)", min_value=8, max_value=24, value=12, step=1)

# 슬라이더 값에 따라 동적으로 CSS를 생성하여 덮어씌움
st.markdown(f"""
    <style>
    /* 표 내부 글자 크기를 슬라이더 값에 맞춤 */
    .merged-table, .merged-table th, .merged-table td {{
        font-size: {base_font_size}px !important;
    }}
    /* 합계 셀은 기본 글자보다 1px 더 크게 설정 */
    .sum-cell {{
        font-size: {base_font_size + 1}px !important;
    }}
    </style>
""", unsafe_allow_html=True)


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
            dl_df = parse_dl_pax(df)
            if dl_df is not None:
                p_all.append(dl_df)
            else:
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
                    btn.innerText = "⏳ 전체 화면 캡처 중... 잠시만요!";
                    try {
                        var win = window.parent;
                        var doc = win.document;
                        
                        if (!win.html2canvas) {
                            var script = doc.createElement('script');
                            script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
                            script.onload = function() { doCap(win, doc, btn); };
                            script.onerror = function() { 
                                alert("⚠️ 라이브러리를 불러올 수 없습니다."); 
                                btn.innerText = "📸 전체 사진으로 저장"; 
                            };
                            doc.head.appendChild(script);
                        } else {
                            doCap(win, doc, btn);
                        }
                    } catch(e) {
                        alert("⚠️ 브라우저 보안 설정으로 인해 캡처가 차단되었습니다.");
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

                    if(appView) { appView.style.overflow = 'visible'; appView.style.height = 'auto'; }
                    if(mainView) { mainView.style.overflow = 'visible'; mainView.style.height = 'auto'; }

                    target.style.paddingTop = '0px';
                    target.style.marginTop = '0px';

                    hides.forEach(function(e){ e.dataset.old = e.style.display; e.style.display = 'none'; });
                    
                    setTimeout(function() {
                        win.html2canvas(target, { 
                            scale: 2, 
                            useCORS: true, 
                            backgroundColor: '#ffffff',
                            scrollY: 0,
                            windowWidth: target.scrollWidth,
                            windowHeight: target.scrollHeight 
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

                            hides.forEach(function(e){ e.style.display = e.dataset.old || ''; });
                            btn.innerText = "📸 전체 사진으로 저장";
                        });
                    }, 800);
                }
                </script>
                """, height=45
            )

            st.markdown(f"""
                <div class="total-banner"><h3 style='margin:0; color:#1E3A8A;'>📊 총 승객수: {total_p:,}명</h3></div>
                <div class="carrier-banner">
                    <span class="carrier-item">KE: <span style="color:#1E3A8A;">{ke_s:,}</span>명</span>
                    <span class="carrier-item">OZ: <span style="color:#1E3A8A;">{oz_s:,}</span>명</span>
                    <span class="carrier-item">DL: <span style="color:#1E3A8A;">{dl_s:,}</span>명</span>
                </div>
                <hr style="margin: 2px 0 10px 0; border: 0; border-top: 1px solid #e5e7eb;">
            """, unsafe_allow_html=True)
            
            west_p = final[final['구역'] == '서편']['p_val'].sum()
            east_p = final[final['구역'] == '동편']['p_val'].sum()
            
            w_html = generate_table_html(final[final['구역'] == '서편'], "⬅️ 서편", west_p, "#DC2626", opt_airline, opt_peak)
            e_html = generate_table_html(final[final['구역'] == '동편'], "➡️ 동편", east_p, "#2563EB", opt_airline, opt_peak)
            
            st.markdown(f'<div class="print-row">{e_html}{w_html}</div>', unsafe_allow_html=True)
