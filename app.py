import streamlit as st
import pandas as pd
import re
from io import BytesIO, StringIO
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="유창강건 마감 킬러", layout="wide")
st.title("📊 유창강건 세금계산서 누락 체크기")
st.info("물품출고(ERP), 카드매출, 세금계산서 발행목록을 대조합니다.")

# ── 공통 함수 ────────────────────────────────────────────────
def clean(x):
    x = str(x)
    x = re.sub(r'[■▲▶●★☆□△◆◇]', '', x)
    x = x.replace("(주)", "").replace("(유)", "").replace("(株)", "")
    x = re.sub(r'\s+', '', x)
    return x.strip()

def is_number(s):
    s = str(s).strip().strip('"').replace(',', '').replace('(','').replace(')','')
    try:
        float(s)
        return True
    except:
        return False

def to_num(x):
    try:
        return int(str(x).replace(',','').replace(' ',''))
    except:
        return 0

def fmt(n):
    try:
        return f"{int(n):,}"
    except:
        return "-"

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

# ── 파일 업로드 ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("#### 1️⃣ 물품출고(ERP)")
    file_out = st.file_uploader("ERP 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
    st.caption("유창강건 ERP에서 내보낸 xls")
with col2:
    st.markdown("#### 2️⃣ 카드매출 비교")
    file_card = st.file_uploader("카드매출 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
    st.caption("월별 시트가 있는 xlsx")
with col3:
    st.markdown("#### 3️⃣ 세금계산서 발행목록")
    file_tax = st.file_uploader("세금계산서 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
    st.caption("국세청 전자세금계산서 목록 xls")

# ── 카드매출 시트 선택 ───────────────────────────────────────
card_bytes = None
selected_sheet = None
if file_card:
    try:
        card_bytes = file_card.read()
        sheet_names = pd.ExcelFile(BytesIO(card_bytes), engine='openpyxl').sheet_names
        selected_sheet = st.selectbox("📅 카드매출 분석 월 선택", options=sheet_names, index=0)
        st.caption(f"선택된 시트: **{selected_sheet}**")
    except Exception as e:
        st.error(f"카드매출 파일 시트 읽기 오류: {e}")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 미발행 업체 분석 시작", type="primary", use_container_width=True):
    if not (file_out and card_bytes and file_tax):
        st.warning("⚠️ 파일 3개를 모두 올려주세요.")
        st.stop()
    if not selected_sheet:
        st.warning("⚠️ 카드매출 월을 선택해주세요.")
        st.stop()

    with st.spinner("분석 중..."):
        try:
            # ── 1. ERP ───────────────────────────────────────
            erp_bytes = file_out.read()
            text = erp_bytes.decode('cp949')
            df_erp = pd.read_csv(StringIO(text), sep='\t', header=0, on_bad_lines='skip')
            col_name = '거래처명' if '거래처명' in df_erp.columns else df_erp.columns[3]

            df_erp['_key'] = df_erp[col_name].apply(lambda x: clean(str(x)))
            df_erp['_공급가액'] = df_erp['공급가액'].apply(to_num)
            df_erp['_부가세']   = df_erp['부가세'].apply(to_num)
            df_erp['_합계금액'] = df_erp['합계금액'].apply(to_num)

            erp_names = set()
            original_map = {}
            for v in df_erp[col_name].dropna():
                v = str(v).strip().strip('"')
                if v and not is_number(v):
                    k = clean(v)
                    erp_names.add(k)
                    original_map[k] = v

            # 거래처별 ERP 금액 합산
            df_erp_valid = df_erp[df_erp['_key'].isin(erp_names)]
            erp_공급가액 = df_erp_valid.groupby('_key')['_공급가액'].sum().to_dict()
            erp_부가세   = df_erp_valid.groupby('_key')['_부가세'].sum().to_dict()
            erp_합계금액 = df_erp_valid.groupby('_key')['_합계금액'].sum().to_dict()

            # ── 2. 카드매출 ───────────────────────────────────
            df_card = pd.read_excel(BytesIO(card_bytes), sheet_name=selected_sheet,
                                    header=None, engine='openpyxl')
            card_names = set()
            for v in df_card.iloc[3:, 0].dropna():
                k = clean(str(v))
                if k:
                    card_names.add(k)

            # ── 3. 세금계산서 ─────────────────────────────────
            tax_bytes = file_tax.read()
            is_xls = tax_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
            df_tax = pd.read_excel(BytesIO(tax_bytes), header=None,
                                   engine='xlrd' if is_xls else 'openpyxl')

            # col11=상호, col14=합계금액, col15=공급가액, col16=세액
            tax_names = set()
            tax_raw_list = []
            for i, v in df_tax.iloc[6:, 11].items():
                if pd.notna(v) and str(v).strip():
                    tax_names.add(clean(str(v)))
                    tax_raw_list.append((i - 5, str(v).strip()))

            df_tax_data = df_tax.iloc[6:].copy()
            df_tax_data['_key']    = df_tax_data[11].apply(lambda x: clean(str(x)))
            df_tax_data['_공급가액'] = df_tax_data[15].apply(to_num)
            df_tax_data['_세액']    = df_tax_data[16].apply(to_num)
            df_tax_data['_합계금액'] = df_tax_data[14].apply(to_num)

            df_tax_valid = df_tax_data[df_tax_data['_key'].isin(tax_names)]
            tax_공급가액 = df_tax_valid.groupby('_key')['_공급가액'].sum().to_dict()
            tax_세액     = df_tax_valid.groupby('_key')['_세액'].sum().to_dict()
            tax_합계금액 = df_tax_valid.groupby('_key')['_합계금액'].sum().to_dict()

            # ── 4. 미발행 업체 ────────────────────────────────
            target = erp_names - card_names
            missing_keys = target - tax_names
            skip_keywords = ['기타거래처', '현금영수증', '카드매출', '거래처명', '거래처ID']

            missing_rows = []
            for k in missing_keys:
                orig = original_map.get(k, k)
                if any(kw in orig for kw in skip_keywords):
                    continue
                similars = find_similar(k, tax_raw_list, threshold=0.72)
                if similars:
                    notes = [f"유사이름: '{s}' (세금계산서 {r}행, {sc:.0%})" for s, r, sc in similars]
                    reason = " / ".join(notes)
                else:
                    reason = "세금계산서 미발행"
                missing_rows.append({"업체명": orig, "비고": reason})
            missing_rows.sort(key=lambda x: x["업체명"])

            # ── 5. 금액 불일치 ────────────────────────────────
            issued_keys = erp_names & tax_names
            amount_rows = []
            for k in issued_keys:
                orig = original_map.get(k, k)
                e_공급 = erp_공급가액.get(k, 0)
                e_부가 = erp_부가세.get(k, 0)
                e_합계 = erp_합계금액.get(k, 0)
                t_공급 = tax_공급가액.get(k, 0)
                t_세액 = tax_세액.get(k, 0)
                t_합계 = tax_합계금액.get(k, 0)

                diff_공급 = e_공급 - t_공급
                diff_부가 = e_부가 - t_세액
                diff_합계 = e_합계 - t_합계

                # 하나라도 차이나면 표시
                if diff_공급 != 0 or diff_부가 != 0 or diff_합계 != 0:
                    amount_rows.append({
                        "업체명":          orig,
                        "ERP_공급가액":    e_공급,
                        "ERP_부가세":      e_부가,
                        "ERP_합계":        e_합계,
                        "세금계산서_공급가액": t_공급,
                        "세금계산서_세액":    t_세액,
                        "세금계산서_합계":    t_합계,
                        "차이_공급가액":   diff_공급,
                        "차이_부가세":     diff_부가,
                        "차이_합계":       diff_합계,
                    })
            amount_rows.sort(key=lambda x: -abs(x["차이_합계"]))

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

    # ── 결과 ─────────────────────────────────────────────────
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 매출 업체",               f"{len(erp_names)}개")
    m2.metric(f"카드매출 ({selected_sheet})",  f"{len(card_names)}개")
    m3.metric("세금계산서 발행",               f"{len(tax_names)}개")
    m4.metric("⚠️ 미발행 업체",               f"{len(missing_rows)}개",
              delta=f"-{len(missing_rows)}", delta_color="inverse")

    tab1, tab2 = st.tabs([
        f"📋 미발행 업체 ({len(missing_rows)}개)",
        f"💰 금액 불일치 ({len(amount_rows)}개)"
    ])

    # ── 탭1: 미발행 ──────────────────────────────────────────
    with tab1:
        st.subheader(f"[{selected_sheet}] 세금계산서 미발행 업체 ({len(missing_rows)}개)")
        if missing_rows:
            df_miss = pd.DataFrame([
                {"No.": i+1, "업체명": r["업체명"], "비고": r["비고"]}
                for i, r in enumerate(missing_rows)
            ])
            def hl_miss(row):
                if "유사이름" in str(row["비고"]):
                    return ['background-color: #fff9c4'] * len(row)
                return [''] * len(row)
            st.dataframe(df_miss.style.apply(hl_miss, axis=1),
                         use_container_width=True, hide_index=True)
            st.caption("🟡 노란색 = 세금계산서에 유사한 이름 있음 → 직접 확인 필요")
        else:
            st.success("🎉 미발행 업체가 없습니다!")

    # ── 탭2: 금액 불일치 ─────────────────────────────────────
    with tab2:
        st.subheader(f"[{selected_sheet}] ERP ↔ 세금계산서 금액 불일치 ({len(amount_rows)}개)")
        if amount_rows:
            df_amt = pd.DataFrame([
                {
                    "No.":             i+1,
                    "업체명":           r["업체명"],
                    "ERP 공급가액":     fmt(r["ERP_공급가액"]),
                    "ERP 부가세":       fmt(r["ERP_부가세"]),
                    "ERP 합계":         fmt(r["ERP_합계"]),
                    "세금계산서 공급가액": fmt(r["세금계산서_공급가액"]),
                    "세금계산서 세액":    fmt(r["세금계산서_세액"]),
                    "세금계산서 합계":    fmt(r["세금계산서_합계"]),
                    "차이(공급가액)":    fmt(r["차이_공급가액"]),
                    "차이(부가세)":      fmt(r["차이_부가세"]),
                    "차이(합계)":        fmt(r["차이_합계"]),
                }
                for i, r in enumerate(amount_rows)
            ])
            def hl_amt(row):
                return ['background-color: #ffe0e0'] * len(row)
            st.dataframe(df_amt.style.apply(hl_amt, axis=1),
                         use_container_width=True, hide_index=True)
            st.caption("차이 양수 = ERP가 더 큼 (세금계산서 추가 발행 필요)  |  음수 = 세금계산서가 더 큼")
        else:
            st.success("🎉 금액이 모두 일치합니다!")

    # ── 엑셀 다운로드 ────────────────────────────────────────
    wb = Workbook()
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    def make_header(ws, titles, title_text):
        ws['A1'] = title_text
        ws['A1'].font = Font(bold=True, size=13, name="맑은 고딕")
        ws.merge_cells(f'A1:{chr(64+len(titles))}1')
        for ci, title in enumerate(titles, 1):
            c = ws.cell(row=2, column=ci, value=title)
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border

    # 시트1: 미발행
    ws1 = wb.active
    ws1.title = "미발행업체"
    make_header(ws1, ["No.", "업체명", "비고"], f"세금계산서 미발행 업체 ({selected_sheet})")
    yellow = PatternFill("solid", fgColor="FFF9C4")
    alt    = PatternFill("solid", fgColor="FFF0F0")
    for i, r in enumerate(missing_rows, 1):
        ws1.cell(i+2, 1, i)
        ws1.cell(i+2, 2, r["업체명"])
        ws1.cell(i+2, 3, r["비고"])
        hl = yellow if "유사이름" in r["비고"] else (alt if i%2==0 else None)
        for ci in range(1, 4):
            c = ws1.cell(i+2, ci)
            c.font = Font(name="맑은 고딕", size=10)
            c.border = border
            c.alignment = Alignment(horizontal='left' if ci>1 else 'center',
                                    vertical='center', wrap_text=(ci==3))
            if hl: c.fill = hl
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 32
    ws1.column_dimensions['C'].width = 55
    for r in range(2, len(missing_rows)+3):
        ws1.row_dimensions[r].height = 28

    # 시트2: 금액불일치
    ws2 = wb.create_sheet("금액불일치")
    h2 = ["No.", "업체명",
          "ERP 공급가액", "ERP 부가세", "ERP 합계",
          "세금계산서 공급가액", "세금계산서 세액", "세금계산서 합계",
          "차이(공급가액)", "차이(부가세)", "차이(합계)"]
    make_header(ws2, h2, f"ERP ↔ 세금계산서 금액 불일치 ({selected_sheet})")
    red  = PatternFill("solid", fgColor="FFE0E0")
    red2 = PatternFill("solid", fgColor="FFF5F5")
    num_cols = {3,4,5,6,7,8,9,10,11}
    for i, r in enumerate(amount_rows, 1):
        vals = [i, r["업체명"],
                r["ERP_공급가액"], r["ERP_부가세"], r["ERP_합계"],
                r["세금계산서_공급가액"], r["세금계산서_세액"], r["세금계산서_합계"],
                r["차이_공급가액"], r["차이_부가세"], r["차이_합계"]]
        for ci, val in enumerate(vals, 1):
            c = ws2.cell(i+2, ci, val)
            c.font = Font(name="맑은 고딕", size=10)
            c.border = border
            c.fill = red if i%2==0 else red2
            c.alignment = Alignment(
                horizontal='right' if ci in num_cols else ('center' if ci==1 else 'left'),
                vertical='center')
            if ci in num_cols:
                c.number_format = '#,##0'
    widths2 = [6, 28, 16, 14, 16, 18, 16, 16, 16, 14, 14]
    for ci, w in enumerate(widths2, 1):
        ws2.column_dimensions[chr(64+ci)].width = w
    for r in range(2, len(amount_rows)+3):
        ws2.row_dimensions[r].height = 22

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button(
        "⬇️ 전체 결과 엑셀 다운로드 (미발행 + 금액불일치)",
        data=buf,
        file_name=f"마감체크_{selected_sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
