import streamlit as st
import pandas as pd
import re
from io import BytesIO, StringIO
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="유창강건 마감 킬러", layout="wide")
st.title("📊 유창강건 세금계산서 누락 체크기")
st.info("(매출대장 + 일반경비) - (카드매출 + 현금영수증) = 세금계산서 발행 대상")

# ── 공통 함수 ────────────────────────────────────────────────
def clean(x):
    x = str(x)
    x = re.sub(r'[■▲▶●★☆□△◆◇]', '', x)
    x = x.replace("(주)", "").replace("(유)", "").replace("(株)", "")
    x = re.sub(r'\s+', '', x)
    return x.strip()

def is_number(s):
    s = str(s).strip().strip('"').replace(',','').replace('(','').replace(')','')
    try: float(s); return True
    except: return False

def to_num(x):
    try: return int(str(x).replace(',','').replace(' ',''))
    except: return 0

def fmt(n):
    try: return f"{int(n):,}"
    except: return "-"

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def find_similar(name_clean, tax_raw_list, threshold=0.72):
    results = []
    for row_idx, raw_name in tax_raw_list:
        score = similarity(name_clean, clean(raw_name))
        if score >= threshold and name_clean != clean(raw_name):
            results.append((raw_name, row_idx, score))
    results.sort(key=lambda x: -x[2])
    return results[:3]

def read_tsv_cp949(b):
    """CP949 인코딩 TSV 파일 읽기 (매출대장, 현금영수증)"""
    text = b.decode('cp949')
    df = pd.read_csv(StringIO(text), sep='\t', header=0, on_bad_lines='skip')
    return df

# ── session_state 초기화 ─────────────────────────────────────
for k in ['매출대장','일반경비','카드매출','현금영수증','계산서']:
    if k not in st.session_state:
        st.session_state[k] = None

# ── 파일 업로드 ─────────────────────────────────────────────
st.markdown("### 📁 파일 업로드")
c1, c2 = st.columns(2)

with c1:
    st.markdown("**발행 대상 ➕**")
    col1, col2 = st.columns(2)
    with col1:
        st.caption("1️⃣ 매출대장 (웍스 청구정보)")
        f1 = st.file_uploader("매출대장", type=['xls','xlsx','csv'], label_visibility="collapsed")
    with col2:
        st.caption("2️⃣ 일반경비 세금계산서 발행건")
        f2 = st.file_uploader("일반경비", type=['xls','xlsx','csv'], label_visibility="collapsed")

with c2:
    st.markdown("**제외 대상 ➖**")
    col3, col4 = st.columns(2)
    with col3:
        st.caption("3️⃣ 카드매출")
        f3 = st.file_uploader("카드매출", type=['xls','xlsx','csv'], label_visibility="collapsed")
    with col4:
        st.caption("4️⃣ 현금영수증")
        f4 = st.file_uploader("현금영수증", type=['xls','xlsx','csv'], label_visibility="collapsed")

st.markdown("**세금계산서 발행 목록 📋**")
f5 = st.file_uploader("5️⃣ 분기표 (세금계산서 발행목록)", type=['xls','xlsx','csv'], label_visibility="collapsed")

# bytes 저장
if f1: st.session_state['매출대장']  = f1.read()
if f2: st.session_state['일반경비']  = f2.read()
if f3: st.session_state['카드매출']  = f3.read()
if f4: st.session_state['현금영수증'] = f4.read()
if f5: st.session_state['계산서']    = f5.read()

# ── 카드매출 시트 선택 ───────────────────────────────────────
selected_sheet = None
if st.session_state['카드매출']:
    try:
        sheet_names = pd.ExcelFile(BytesIO(st.session_state['카드매출']), engine='openpyxl').sheet_names
        selected_sheet = st.selectbox("📅 카드매출 분석 월 선택", options=sheet_names, index=0)
        st.caption(f"선택된 시트: **{selected_sheet}**")
    except Exception as e:
        st.error(f"카드매출 파일 오류: {e}")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 미발행 업체 분석 시작", type="primary", use_container_width=True):
    missing = [k for k in ['매출대장','일반경비','카드매출','현금영수증','계산서']
               if not st.session_state[k]]
    if missing:
        st.warning(f"⚠️ 파일을 올려주세요: {', '.join(missing)}"); st.stop()
    if not selected_sheet:
        st.warning("⚠️ 카드매출 월을 선택해주세요."); st.stop()

    with st.spinner("분석 중..."):
        try:
            # ── 1. 매출대장 (CP949 TSV) ───────────────────────
            df_출고 = read_tsv_cp949(st.session_state['매출대장'])
            df_출고['_key']    = df_출고['거래처'].apply(lambda x: clean(str(x)))
            df_출고['_공급가액'] = df_출고['공급가액+운송'].apply(to_num)
            df_출고['_부가세'] = (df_출고['부가세_품목'].apply(to_num) + df_출고['운송비(부가세)'].apply(to_num))
            df_출고['_합계']     = df_출고['금액'].apply(to_num)

            출고_names = set()
            original_map = {}
            for v in df_출고['거래처'].dropna():
                v = str(v).strip().strip('"')
                if v and not is_number(v):
                    k = clean(v)
                    출고_names.add(k)
                    original_map[k] = v

            df_출고_v = df_출고[df_출고['_key'].isin(출고_names)]
            출고_공급 = df_출고_v.groupby('_key')['_공급가액'].sum().to_dict()
            출고_부가 = df_출고_v.groupby('_key')['_부가세'].sum().to_dict()
            출고_합계 = df_출고_v.groupby('_key')['_합계'].sum().to_dict()

            # ── 2. 일반경비 (xlsx) ────────────────────────────
            df_경비 = pd.read_excel(BytesIO(st.session_state['일반경비']),
                                    header=0, engine='openpyxl')
            df_경비['_key']    = df_경비['거래처명'].apply(lambda x: clean(str(x)))
            df_경비['_공급가액'] = df_경비['공급가액'].apply(to_num)
            df_경비['_부가세']   = df_경비['세액'].apply(to_num)
            df_경비['_합계']     = df_경비['계'].apply(to_num)

            경비_names = set()
            for v in df_경비['거래처명'].dropna():
                v = str(v).strip()
                if v and not is_number(v):
                    k = clean(v)
                    경비_names.add(k)
                    if k not in original_map:
                        original_map[k] = v

            df_경비_v = df_경비[df_경비['_key'].isin(경비_names)]
            경비_공급 = df_경비_v.groupby('_key')['_공급가액'].sum().to_dict()
            경비_부가 = df_경비_v.groupby('_key')['_부가세'].sum().to_dict()
            경비_합계 = df_경비_v.groupby('_key')['_합계'].sum().to_dict()

            # 발행 대상 업체 전체 = 매출대장 + 일반경비
            all_target_names = 출고_names | 경비_names

            # 통합 금액 (매출대장 + 일반경비 합산)
            all_공급 = {}
            all_부가 = {}
            all_합계 = {}
            for k in all_target_names:
                all_공급[k] = 출고_공급.get(k, 0) + 경비_공급.get(k, 0)
                all_부가[k] = 출고_부가.get(k, 0) + 경비_부가.get(k, 0)
                all_합계[k] = 출고_합계.get(k, 0) + 경비_합계.get(k, 0)

            # ── 3. 카드매출 (xlsx, 선택 시트) ─────────────────
            df_카드 = pd.read_excel(BytesIO(st.session_state['카드매출']),
                                    sheet_name=selected_sheet, header=None, engine='openpyxl')
            # 헤더 row2(index 2), 데이터 row3(index 3)부터
            card_names = set()
            for v in df_카드.iloc[3:, 0].dropna():
                k = clean(str(v))
                if k: card_names.add(k)

            # ── 4. 현금영수증 (CP949 TSV) ─────────────────────
            df_현금 = read_tsv_cp949(st.session_state['현금영수증'])
            cash_names = set()
            for v in df_현금['거래처'].dropna():
                v = str(v).strip().strip('"')
                if v and not is_number(v):
                    cash_names.add(clean(v))

            # 제외 업체 = 카드매출 + 현금영수증
            exclude_names = card_names | cash_names

            # ── 5. 세금계산서 (분기표) ─────────────────────────
            tax_b = st.session_state['계산서']
            is_xls = tax_b[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
            df_tax = pd.read_excel(BytesIO(tax_b), header=None,
                                   engine='xlrd' if is_xls else 'openpyxl')

            tax_names = set()
            tax_raw_list = []
            for i, v in df_tax.iloc[6:, 11].items():
                if pd.notna(v) and str(v).strip():
                    tax_names.add(clean(str(v)))
                    tax_raw_list.append((i - 5, str(v).strip()))

            df_tax_d = df_tax.iloc[6:].copy()
            df_tax_d['_key']    = df_tax_d[11].apply(lambda x: clean(str(x)))
            df_tax_d['_공급가액'] = df_tax_d[15].apply(to_num)
            df_tax_d['_세액']    = df_tax_d[16].apply(to_num)
            df_tax_d['_합계금액'] = df_tax_d[14].apply(to_num)
            df_tax_v = df_tax_d[df_tax_d['_key'].isin(tax_names)]
            tax_공급 = df_tax_v.groupby('_key')['_공급가액'].sum().to_dict()
            tax_세액 = df_tax_v.groupby('_key')['_세액'].sum().to_dict()
            tax_합계 = df_tax_v.groupby('_key')['_합계금액'].sum().to_dict()

            # ── 6. 미발행 계산 ────────────────────────────────
            # 발행대상 - 제외 - 이미발행 = 미발행
            target       = all_target_names - exclude_names
            missing_keys = target - tax_names
            skip_kw = ['기타거래처','현금영수증','카드매출','거래처명','거래처ID','유창']

            missing_rows = []
            for k in missing_keys:
                orig = original_map.get(k, k)
                if any(kw in orig for kw in skip_kw): continue
                similars = find_similar(k, tax_raw_list, 0.72)
                reason = (" / ".join([f"유사이름: '{s}' (계산서 {r}행, {sc:.0%})"
                                      for s,r,sc in similars])
                          if similars else "세금계산서 미발행")
                missing_rows.append({"업체명": orig, "비고": reason})
            missing_rows.sort(key=lambda x: x["업체명"])

            # ── 7. 금액 불일치 ────────────────────────────────
            issued_keys = (all_target_names - exclude_names) & tax_names
            amount_rows = []
            for k in issued_keys:
                orig  = original_map.get(k, k)
                e_공급 = all_공급.get(k, 0)
                e_부가 = all_부가.get(k, 0)
                e_합계 = all_합계.get(k, 0)
                t_공급 = tax_공급.get(k, 0)
                t_세액 = tax_세액.get(k, 0)
                t_합계 = tax_합계.get(k, 0)
                d_공급 = e_공급 - t_공급
                d_부가 = e_부가 - t_세액
                d_합계 = d_공급 + d_부가
                if d_공급 != 0 or d_부가 != 0:
                    amount_rows.append({
                        "업체명": orig,
                        "ERP_공급가액": e_공급, "ERP_부가세": e_부가, "ERP_합계": e_합계,
                        "세금계산서_공급가액": t_공급, "세금계산서_세액": t_세액, "세금계산서_합계": t_합계,
                        "차이_공급가액": d_공급, "차이_부가세": d_부가, "차이_합계": d_합계,
                    })
            amount_rows.sort(key=lambda x: -abs(x["차이_합계"]))

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback; st.code(traceback.format_exc()); st.stop()

    # ── 결과 ─────────────────────────────────────────────────
    st.divider()
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("매출대장 업체",           f"{len(출고_names)}개")
    c2.metric("일반경비 업체",           f"{len(경비_names)}개")
    c3.metric(f"카드+현금영수증 제외",    f"{len(exclude_names)}개")
    c4.metric("세금계산서 발행",         f"{len(tax_names)}개")
    c5.metric("⚠️ 미발행 업체",         f"{len(missing_rows)}개",
              delta=f"-{len(missing_rows)}", delta_color="inverse")

    tab1, tab2 = st.tabs([
        f"📋 미발행 업체 ({len(missing_rows)}개)",
        f"💰 금액 불일치 ({len(amount_rows)}개)"
    ])

    with tab1:
        st.subheader(f"[{selected_sheet}] 세금계산서 미발행 업체 ({len(missing_rows)}개)")
        if missing_rows:
            df_m = pd.DataFrame([{"No.":i+1,"업체명":r["업체명"],"비고":r["비고"]}
                                  for i,r in enumerate(missing_rows)])
            def hl_m(row):
                return (['background-color:#fff9c4']*len(row)
                        if "유사이름" in str(row["비고"]) else ['']*len(row))
            st.dataframe(df_m.style.apply(hl_m,axis=1),
                         use_container_width=True, hide_index=True)
            st.caption("🟡 노란색 = 세금계산서에 유사한 이름 있음 → 직접 확인 필요")
        else:
            st.success("🎉 미발행 업체 없음!")

    with tab2:
        st.subheader(f"[{selected_sheet}] 금액 불일치 ({len(amount_rows)}개)")
        if amount_rows:
            df_a = pd.DataFrame([{
                "No.": i+1, "업체명": r["업체명"],
                "ERP 공급가액":      fmt(r["ERP_공급가액"]),
                "ERP 부가세":        fmt(r["ERP_부가세"]),
                "ERP 합계":          fmt(r["ERP_합계"]),
                "세금계산서 공급가액": fmt(r["세금계산서_공급가액"]),
                "세금계산서 세액":    fmt(r["세금계산서_세액"]),
                "세금계산서 합계":    fmt(r["세금계산서_합계"]),
                "차이(공급가액)":     fmt(r["차이_공급가액"]),
                "차이(부가세)":       fmt(r["차이_부가세"]),
                "차이(합계)":         fmt(r["차이_합계"]),
            } for i,r in enumerate(amount_rows)])

            col_colors = {
                "No.":               "#F5F5F5",
                "업체명":             "#E8F4FD",
                "ERP 공급가액":       "#FFF3E0",
                "ERP 부가세":         "#FFF3E0",
                "ERP 합계":           "#FFE0B2",
                "세금계산서 공급가액": "#F3E5F5",
                "세금계산서 세액":     "#F3E5F5",
                "세금계산서 합계":     "#E1BEE7",
                "차이(공급가액)":      "#FCE4EC",
                "차이(부가세)":        "#FCE4EC",
                "차이(합계)":          "#EF9A9A",
            }
            def hl_a(row):
                return [f"background-color:{col_colors.get(c,'#FFF')}" for c in row.index]
            st.dataframe(df_a.style.apply(hl_a,axis=1),
                         use_container_width=True, hide_index=True)
            st.caption("🔵 업체명  🟠 ERP  🟣 세금계산서  🔴 차이  |  양수=ERP더 큼 | 음수=세금계산서더 큼")
        else:
            st.success("🎉 금액 모두 일치!")

    # ── 엑셀 다운로드 ────────────────────────────────────────
    wb = Workbook()
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    def mk_hdr(ws, titles, txt):
        ws['A1'] = txt
        ws['A1'].font = Font(bold=True, size=13, name="맑은 고딕")
        ws.merge_cells(f'A1:{chr(64+len(titles))}1')
        for ci, t in enumerate(titles, 1):
            c = ws.cell(2, ci, t)
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border

    # 시트1: 미발행
    ws1 = wb.active; ws1.title = "미발행업체"
    mk_hdr(ws1, ["No.","업체명","비고"], f"세금계산서 미발행 업체 ({selected_sheet})")
    yf = PatternFill("solid", fgColor="FFF9C4")
    af = PatternFill("solid", fgColor="FFF0F0")
    for i, r in enumerate(missing_rows, 1):
        ws1.cell(i+2,1,i); ws1.cell(i+2,2,r["업체명"]); ws1.cell(i+2,3,r["비고"])
        hl = yf if "유사이름" in r["비고"] else (af if i%2==0 else None)
        for ci in range(1,4):
            c = ws1.cell(i+2,ci)
            c.font = Font(name="맑은 고딕", size=10); c.border = border
            c.alignment = Alignment(horizontal='left' if ci>1 else 'center',
                                    vertical='center', wrap_text=(ci==3))
            if hl: c.fill = hl
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 32
    ws1.column_dimensions['C'].width = 55
    for r in range(2, len(missing_rows)+3): ws1.row_dimensions[r].height = 28

    # 시트2: 금액불일치
    ws2 = wb.create_sheet("금액불일치")
    h2 = ["No.","업체명","ERP 공급가액","ERP 부가세","ERP 합계",
          "세금계산서 공급가액","세금계산서 세액","세금계산서 합계",
          "차이(공급가액)","차이(부가세)","차이(합계)"]
    mk_hdr(ws2, h2, f"ERP ↔ 세금계산서 금액 불일치 ({selected_sheet})")
    xls_colors = ["F5F5F5","E8F4FD","FFF3E0","FFF3E0","FFE0B2",
                  "F3E5F5","F3E5F5","E1BEE7","FCE4EC","FCE4EC","EF9A9A"]
    nc = {3,4,5,6,7,8,9,10,11}
    for i, r in enumerate(amount_rows, 1):
        vals = [i, r["업체명"],
                r["ERP_공급가액"], r["ERP_부가세"], r["ERP_합계"],
                r["세금계산서_공급가액"], r["세금계산서_세액"], r["세금계산서_합계"],
                r["차이_공급가액"], r["차이_부가세"], r["차이_합계"]]
        for ci, val in enumerate(vals, 1):
            c = ws2.cell(i+2, ci, val)
            c.font = Font(name="맑은 고딕", size=10); c.border = border
            c.fill = PatternFill("solid", fgColor=xls_colors[ci-1])
            c.alignment = Alignment(
                horizontal='right' if ci in nc else ('center' if ci==1 else 'left'),
                vertical='center')
            if ci in nc: c.number_format = '#,##0'
    for ci, w in enumerate([6,28,16,14,16,18,16,16,16,14,14], 1):
        ws2.column_dimensions[chr(64+ci)].width = w
    for r in range(2, len(amount_rows)+3): ws2.row_dimensions[r].height = 22

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button(
        "⬇️ 전체 결과 엑셀 다운로드 (미발행 + 금액불일치)",
        data=buf,
        file_name=f"마감체크_{selected_sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
