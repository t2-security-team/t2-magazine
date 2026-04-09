import streamlit as st
import pandas as pd
import re

# 1. 페이지 설정
st.set_page_config(page_title="T2 보안검색 환승부 잡지", layout="wide")

# --- [💡 핵심: 튕김 현상 완벽 해결을 위한 저장소] ---
# 칸을 넉넉히 10개로 고정하여 지저분한 왼쪽 선택칸을 없앰
if 'saved_manual_df' not in st.session_state:
    st.session_state.saved_manual_df = pd.DataFrame([{"편명": "", "승객수": ""} for _ in range(10)])

# --- [사이드바 설정] ---
with st.sidebar:
    st.header("📂 데이터 업로드")
    pax_files = st.file_uploader("1. 승객수 파일 (.xlsx, .csv)", accept_multiple_files=True)
    gate_files = st.file_uploader("2. 게이트 파일 (.xlsx, .csv)", accept_multiple_files=True)
    
    st.divider()
    st.markdown("**※ 델타 엑셀파일은 없고 사진밖에 없을 때**")
    use_manual_input = st.checkbox("📸 독립된 수기 입력창 열기", help="체크하면 화면에 직접 입력할 수 있는 창이 나타납니다.")
    st.divider()
    
    st.markdown("### 🔠 글씨 크기 조절")
    font_offset = st.slider("출력용 글씨 크기 (기본: 0)", min_value=-3, max_value=7, value=0, step=1)
    
    st.divider()
    st.markdown("### 🎨 시각화 옵션")
    view_option = st.radio(
        "원하시는 표시 방식을 선택하세요:",
        ("⬜ 기본 (색상 없음)", "✈️ 1. 항공사별 색상 표시 (DL:연노랑, OZ:연분홍)", "⏰ 2. 첨두시간 색상 표시 (16~18시)"),
        index=0 
    )
    opt_airline = "1. 항공사별" in view_option
    opt_peak = "2. 첨두시간" in view_option
    
    st.divider()
    time_range = st.slider("조회 시간대 (시)", 0, 24, (0, 24))

# --- [글씨 크기 동적 계산] ---
tbl_fs = 13 + font_offset
sum_fs = 14 + font_offset
car_fs = 15 + font_offset
tit_fs = 17 + font_offset

# --- [디자인 및 PDF 압축 CSS] ---
st.markdown(f"""
    <style>
    .main .block-container {{ padding-top: 0px !important; padding-bottom: 0px !important; margin-top: -15px !important; }}
    div[data-testid="stVerticalBlock"] {{ gap: 0px !important; }}
    .element-container {{ margin-bottom: 0px !important; }}
    iframe {{ margin-bottom: 0px !important; min-height: 45px !important; }}
    
    .merged-table {{ width: 100%; border-collapse: collapse; font-size: {tbl_fs}px; text-align: center; font-family: sans-serif; margin-bottom: 0px !important; }}
    .merged-table tr {{ border: none !important; }} 
    .merged-table th {{ background-color: #f8f9fa !important; border: 1px solid #dee2e6 !important; padding: 4px; font-weight: bold; }}
    .merged-table td {{ border: 1px solid #dee2e6 !important; padding: 3px; vertical-align: middle; }}
    
    .sum-cell {{ background-color: #ffffff; font-weight: bold; color: #1E3A8A; font-size: {sum_fs}px; vertical-align: middle !important; position: relative; z-index: 10; }}
    
    .total-banner {{ background-color: #f0f7ff !important; padding: 8px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 2px; margin-top: 2px; }}
    .carrier-banner {{ background-color: #ffffff !important; padding: 4px; border-radius: 8px; text-align: center; border: 1px solid #3b82f6; margin-bottom: 4px; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap; }}
    .carrier-item {{ font-size: {car_fs}px; font-weight: bold; }}
    
    .table-title {{ text-align: center; font-size: {tit_fs}px; margin-top: 2px; margin-bottom: 5px; }}
    
    .print-row {{ display: flex; flex-direction: row; gap: 15px; width: 100%; }}
    .print-col {{ flex: 1; min-width: 0; margin-bottom: 0px !important; }}
    
    @media print {{
        .no-print, header, footer, [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"], iframe {{ display: none !important; }}
        [data-testid="stDataEditor"], .manual-box, button {{ display: none !important; }}
        [data-testid="stElementContainer"]:has([data-testid="stDataEditor"]), [data-testid="stElementContainer"]:has(button) {{ display: none !important; }}
        
        html, body {{ height: auto !important; min-height: auto !important; padding-bottom: 0 !important; margin-bottom: 0 !important; padding-top: 0 !important; }}
        .appview-container, .main, .block-container, .element-container {{ padding-top: 0 !important; margin-top: 0 !important; padding-bottom: 0 !important; margin-bottom: 0 !important; }}
        div[data-testid="stVerticalBlock"] {{ gap: 0 !important; }}
        body {{ zoom: 75%; }}
        .print-row {{ display: flex !important; flex-direction: row !important; }}
        table {{ page-break-inside: auto; margin-bottom: 0px !important; }}
        tr {{ page-break-inside: avoid; page-break-after: auto; }}
        thead {{ display: table-header-group; }}
        @page {{ size: A4; margin-top: 12mm !important; margin-bottom: 12mm !important; margin-left: 10mm !important; margin-right: 10mm !important; }}
        @page :first {{ margin-top: 0mm !important; }}
    }}
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
    for i, row in enumerate(all_rows):
        for cell in row:
            if str(cell).replace(" ", "").strip() == '환승객':
                pax_row_idx = i; pax_row_data = row; break
        if pax_row_idx != -1: break
    if pax_row_idx != -1:
        dl_data = []
        for col_idx, cell in enumerate(all_rows[0]):
            cell_str = str(cell)
            if 'DL' in cell_str.upper() and re.search(r'DL\s*\d+', cell_str, re.IGNORECASE):
                flt_no = clean_flight_no(re.search(r'(DL\s*\d+)', cell_str, re.IGNORECASE).group(1)) 
                if col_idx < len(pax_row_data):
                    try: dl_data.append({'편명': flt_no, '승객수': int(float(str(pax_row_data[col_idx]).replace(",", "").strip()))})
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

def clean_route(val):
    if pd.isna(val): return ""
    val = str(val).strip()
    match = re.search(r'(.*?)\s*(\([A-Za-z0-9]+\))', val)
    if match: return f"{match.group(1).split('/')[0].strip()}{match.group(2).strip()}"
    if '/' in val: return val.split('/')[0].strip()
    return val

def generate_table_html(df, title, count, color, opt_airline, opt_peak):
    display_title = f"{title} ({count:,}명)"
    html = f"<div class='print-col'><h3 class='table-title' style='color:{color};'>{display_title}</h3>"
    if df.empty: return html + "<div style='text-align:center; padding:20px; border:1px solid #ddd;'>데이터 없음</div></div>"
    
    df = df.sort_values('시간').reset_index(drop=True)
    html += '<table class="merged-table"><thead><tr><th style="width:14%;">예상시간</th><th style="width:12%;">시간</th><th>출발지</th><th style="width:14%;">편명</th><th style="width:11%;">게이트</th><th style="width:11%;">승객</th><th style="width:11%;">합계</th></tr></thead><tbody>'
    
    df['hour_val'] = df['시간'].astype(str).str.extract(r'(\d+)', expand=False).fillna(0).astype(int)
    hour_counts = df['hour_val'].value_counts().sort_index()
    hour_sums = df.groupby('hour_val')['p_val'].sum()
    processed_hours = set()
    
    for i, row in df.iterrows():
        curr_h = row['hour_val']; flt = str(row['편명']).upper()
        td_style = ""; sum_bg = "#ffffff" 
        
        if opt_peak:
            if curr_h == 16: td_style = ' style="background-color: #F4FAFD;"'; sum_bg = "#F4FAFD"
            elif curr_h == 17: td_style = ' style="background-color: #FFFDF0;"'; sum_bg = "#FFFDF0"
            elif curr_h == 18: td_style = ' style="background-color: #FFF5F8;"'; sum_bg = "#FFF5F8"

        if opt_airline:
            if flt.startswith("DL"): td_style = ' style="background-color: #FFFDE7;"' 
            elif flt.startswith("OZ"): td_style = ' style="background-color: #FDF4F7;"' 

        html += f'<tr><td{td_style}></td><td{td_style}>{row["시간"]}</td><td{td_style}>{row["출발지"]}</td><td{td_style}>{row["편명"]}</td><td{td_style}>{row["게이트"]}</td><td{td_style}>{row["p_val"]:,}</td>'
        
        if curr_h not in processed_hours:
            html += f'<td rowspan="{hour_counts[curr_h]}" class="sum-cell" style="background-color: {sum_bg};"><span style="position: relative; z-index: 20;">{hour_sums[curr_h]:,}</span></td>'
            processed_hours.add(curr_h)
        html += '</tr>'
    return html + '</tbody></table></div>'

# --- [메인 로직] ---

# 1. 독립된 창 (수기 입력 UI) - 입력 완료 후 저장 버튼 방식!
if use_manual_input:
    st.markdown("<div class='no-print manual-box' style='background-color: #FFFDE7; padding: 20px; border-radius: 12px; border: 2px solid #FBC02D; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>", unsafe_allow_html=True)
    st.markdown("<h3 style='margin-top:0; color: #1E3A8A;'>✍️ 델타항공 독립 수기 입력창</h3>", unsafe_allow_html=True)
    
    # 💡 요청하신 4번, 5번 안내 문구 추가
    st.markdown("""
        <div style='font-size:16px; margin-bottom:10px; line-height: 1.6;'>
        1️⃣ 사진을 보고 아래 표에 편명과 숫자를 모두 입력하세요.<br>
        2️⃣ 입력을 모두 마쳤으면 <b>아래의 [💾 데이터 저장] 버튼</b>을 꼭 눌러주세요.<br>
        <span style='color:#DC2626; font-weight:bold;'>3️⃣ 저장이 완료되면 좌측 메뉴의 체크박스를 한 번 더 눌러서 이 창을 꼭 닫아주세요!</span><br>
        <span style='color:#EAB308; font-weight:bold;'>4️⃣ 입력하던 내용이 반영 안되거나 자동으로 지워지면 다시 한번 입력해주세요.</span><br>
        <span style='color:#16A34A; font-weight:bold;'>5️⃣ 수기로 입력한 델타 승객수는 꼭 사진 총 승객수랑 비교해주세요.</span>
        </div>
    """, unsafe_allow_html=True)
    
    # num_rows="dynamic"을 제거하여 불필요한 왼쪽 선택칸 삭제 (대신 위에서 10칸 넉넉히 제공)
    temp_df = st.data_editor(st.session_state.saved_manual_df, key="manual_editor", use_container_width=True, hide_index=True)
    
    # 저장 버튼을 눌러야만 실제 데이터에 반영됨 (튕김 원천 차단)
    if st.button("💾 데이터 저장 (입력 완료 후 클릭!)", use_container_width=True):
        st.session_state.saved_manual_df = temp_df
        st.success("✅ 저장이 완료되었습니다! 이제 좌측의 체크박스를 한 번 더 눌러 창을 닫아주세요.")
    
    st.markdown("</div>", unsafe_allow_html=True)


# 데이터 유효성 검사
valid_manual_check = st.session_state.saved_manual_df[st.session_state.saved_manual_df["편명"].astype(str).str.strip() != ""]

# 아무 데이터도 없을 때 초기 안내 화면 문구 원복
if not (pax_files and gate_files) and valid_manual_check.empty and not use_manual_input:
    st.markdown("<h2 style='text-align: center;'>✈️ T2 보안검색 환승부 잡지 ✈️</h2>", unsafe_allow_html=True)
    with st.expander("💡 홈페이지 이용 방법 및 주의사항 (필독)", expanded=True):
        st.markdown("""
        ### 1. 파일 업로드 방법
        * **1번째 파일 업로드 (승객수 파일):** 이메일로 받은 승객수(T/S, Pax) 데이터 업로드
        * **2번째 파일 업로드 (게이트 파일):** 인천공항 게이트 및 도착시간 데이터 업로드
        * **- 인천공항 도착편 T2, 날짜, 시간대(00:00~23:59) 설정 후 검색 다운로드**

        ### 2. 중요: 파일 형식 필수 변환
        * 본 시스템은 **.xlsx 형식만 지원**합니다.
        * **인천공항 도착편** 다운로드 파일은 그대로 올리면 읽히지 않습니다.
        * **방법:** 파일을 열어 **[다른 이름으로 저장]** → 파일 형식을 **[Excel 통합 문서 (*.xlsx)]**로 선택하여 저장 후 업로드하세요.
        """)
else:
    p_all, g_all = [], []
    
    # 1. 저장된 수기 데이터 합치기
    if not valid_manual_check.empty:
        valid_manual_copy = valid_manual_check.copy()
        valid_manual_copy['편명'] = valid_manual_copy['편명'].astype(str).apply(clean_flight_no)
        p_all.append(valid_manual_copy)

    # 2. 승객수 파일 합치기
    for f in pax_files:
        df = smart_read(f)
        if df is not None:
            dl_df = parse_dl_pax(df)
            if dl_df is not None and not dl_df.empty:
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
                    if not tmp.empty:
                        p_all.append(tmp)

    # 3. 게이트 파일 합치기
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
                if not tmp.empty:
                    g_all.append(tmp)

    # 4. 데이터 병합 및 시각화
    if p_all and g_all:
        df_p = pd.concat(p_all).drop_duplicates('편명', keep='last')
        df_g = pd.concat(g_all).drop_duplicates('편명')
        
        df_p = df_p[df_p['편명'] != ""]
        df_g = df_g[df_g['편명'] != ""]
        
        if df_p.empty or df_g.empty:
            st.error("오류: 업로드된 파일이나 입력된 데이터에서 유효한 편명 또는 승객 정보를 찾을 수 없습니다.")
            st.stop()
            
        final = pd.merge(df_g, df_p, on='편명', how='inner', suffixes=('', '_p'))
        
        if not final.empty:
            # 빈 문자열("")도 여기서 알아서 0으로 변환 처리되므로 에러 없음!
            final['p_val'] = pd.to_numeric(final['승객수'], errors='coerce').fillna(0).astype(int)
            final['hour'] = final['시간'].astype(str).str.extract(r'(\d+)', expand=False).fillna(0).astype(int)
            final = final[(final['hour'] >= time_range[0]) & (final['hour'] <= time_range[1])]
            
            final['게이트'] = final['게이트'].astype(str).str.strip()
            final['g_num'] = pd.to_numeric(final['게이트'], errors='coerce').fillna(0)
            final['구역'] = final['g_num'].apply(lambda x: '서편' if 0 < x <= 250 else '동편')
            
            total_p = final['p_val'].sum()
            def c_sum(c): return final[final['편명'].str.startswith(c, na=False)]['p_val'].sum()
            ke_s, oz_s, dl_s = c_sum('KE'), c_sum('OZ'), c_sum('DL')

            # --- 결과 출력 및 자바스크립트 캡처/인쇄 버튼 ---
            st.components.v1.html(
                """
                <style>
                body { margin: 0; padding: 0; overflow: hidden; display: flex; gap: 10px; }
                .custom-btn { background-color: white; border: 1px solid #dcdcdc; color: #31333f; padding: 6px 15px; font-size: 14px; border-radius: 6px; cursor: pointer; font-family: sans-serif; box-shadow: 0px 1px 3px rgba(0,0,0,0.1); }
                .custom-btn:hover { border-color: #ff4b4b; color: #ff4b4b; }
                </style>
                <button class="custom-btn" onclick="prepareAndPrint()">📄 PDF 저장</button>
                <button class="custom-btn" onclick="takePic()" id="pic-btn">📸 전체 사진으로 저장</button>
                
                <script>
                function getHidableElements(doc) {
                    var elements = Array.from(doc.querySelectorAll('[data-testid="stSidebar"], header, iframe, .no-print, .manual-box'));
                    var editors = Array.from(doc.querySelectorAll('[data-testid="stDataEditor"]'));
                    editors.forEach(function(ed) {
                        var p = ed.closest('[data-testid="stElementContainer"]');
                        if(p) elements.push(p);
                        elements.push(ed);
                    });
                    var btns = Array.from(doc.querySelectorAll('button')).filter(b => b.innerText.includes('저장'));
                    btns.forEach(function(b) {
                        var p = b.closest('[data-testid="stElementContainer"]');
                        if(p) elements.push(p);
                        elements.push(b);
                    });
                    return elements;
                }

                function prepareAndPrint() {
                    var win = window.parent; var doc = win.document;
                    var hides = getHidableElements(doc);
                    hides.forEach(function(e){ e.dataset.oldPrint = e.style.display; e.style.display = 'none'; });
                    win.print();
                    setTimeout(function() { hides.forEach(function(e){ e.style.display = e.dataset.oldPrint || ''; }); }, 800);
                }

                function takePic() {
                    var btn = document.getElementById('pic-btn'); btn.innerText = "⏳ 캡처 중...";
                    try {
                        var win = window.parent; var doc = win.document;
                        if (!win.html2canvas) { var script = doc.createElement('script'); script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"; script.onload = function() { doCap(win, doc, btn); }; doc.head.appendChild(script); } else { doCap(win, doc, btn); }
                    } catch(e) { alert("캡처 오류"); btn.innerText = "📸 전체 사진으로 저장"; }
                }

                function doCap(win, doc, btn) {
                    var target = doc.querySelector('.block-container') || doc.querySelector('.main');
                    var hides = getHidableElements(doc);
                    var appView = doc.querySelector('.appview-container') || doc.querySelector('[data-testid="stAppViewContainer"]');
                    var mainView = doc.querySelector('.main');
                    var oldAppOverflow = appView ? appView.style.overflow : ''; var oldAppHeight = appView ? appView.style.height : ''; var oldMainOverflow = mainView ? mainView.style.overflow : ''; var oldMainHeight = mainView ? mainView.style.height : '';
                    if(appView) { appView.style.overflow = 'visible'; appView.style.height = 'auto'; }
                    if(mainView) { mainView.style.overflow = 'visible'; mainView.style.height = 'auto'; }
                    var oldPadding = target.style.paddingTop; var oldMargin = target.style.marginTop;
                    target.style.paddingTop = '0px'; target.style.marginTop = '0px';
                    hides.forEach(function(e){ e.dataset.old = e.style.display; e.style.display = 'none'; });
                    setTimeout(function() {
                        win.html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#ffffff', scrollY: 0, windowWidth: target.scrollWidth, windowHeight: target.scrollHeight }).then(function(canvas) {
                            var link = doc.createElement('a'); link.download = '보안검색_잡지_전체.png'; link.href = canvas.toDataURL('image/png'); link.click();
                        }).finally(function() {
                            if(appView) { appView.style.overflow = oldAppOverflow; appView.style.height = oldAppHeight; }
                            if(mainView) { mainView.style.overflow = oldMainOverflow; mainView.style.height = oldMainHeight; }
                            target.style.paddingTop = oldPadding; target.style.marginTop = oldMargin;
                            hides.forEach(function(e){ e.style.display = e.dataset.old || ''; }); btn.innerText = "📸 전체 사진으로 저장";
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
