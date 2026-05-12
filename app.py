import streamlit as st
import pandas as pd
import re
from io import BytesIO, StringIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="외상매출 vs 청구 비교", layout="wide", page_icon="📊")
st.title("📊 외상매출 vs 청구 차이 검출기")
st.info("거래처별 외상매출(금월매출) vs 청구금액을 비교해서 차이 나는 곳을 찾습니다.")

# ── 공통 함수 ────────────────────────────────────────────────
def clean(x):
    """거래처명 정규화"""
    x = str(x).strip().strip('"')
    if x in ('', 'nan', 'NaN', 'None'):
        return ''
    x = re.sub(r'[■□▣▪▫▲△▼▽●○◆◇★☆▶◀♠♣♥♦]', '', x)
    x = x.replace("(주)", "").replace("(유)", "").replace("(株)", "").replace("㈜", "")
    x = x.replace("주식회사", "")
    x = re.sub(r'\([^)]*\)', '', x)
    x = re.sub(r'\[[^\]]*\]', '', x)
    x = re.sub(r'\s+\S+?(공장|영업소|지점|지사|사무소|물류센터)\s*$', '', x)
    x = re.sub(r'\s+', '', x)
    x = re.sub(r'[·•・]', '', x)
    x = re.sub(r'(서울영업소|부산영업소|영업소|지점|지사|본사|본점)$', '', x)
    if re.search(r'[가-힣]', x):
        x = re.sub(r'[a-zA-Z]$', '', x)
    return x.strip().lower()

def to_num(x):
    try:
        s = str(x).replace(',','').replace(' ','').strip('"').strip()
        if not s or s in ('nan','NaN'): return 0
        return int(float(s))
    except: return 0

def fmt(n):
    try: return f"{int(n):,}"
    except: return "-"

def read_tsv_cp949(b):
    """CP949 TSV 읽기 (외상매출/청구 둘 다 이 포맷)"""
    text = b.decode('cp949', errors='replace')
    return pd.read_csv(StringIO(text), sep='\t', quotechar='"', on_bad_lines='skip')

# ── session state ───────────────────────────────────────────
for k in ['외상매출', '청구']:
    if k not in st.session_state:
        st.session_state[k] = None

# ── 파일 업로드 ──────────────────────────────────────────────
st.markdown("### 📁 파일 업로드")
c1, c2 = st.columns(2)
with c1:
    st.caption("1️⃣ **외상매출**.XLS (거래처별 잔액)")
    f1 = st.file_uploader("외상매출 파일", type=['xls','XLS','xlsx','csv'], label_visibility="collapsed")
with c2:
    st.caption("2️⃣ **청구**.XLS (청구서별 상세)")
    f2 = st.file_uploader("청구 파일", type=['xls','XLS','xlsx','csv'], label_visibility="collapsed")

if f1: st.session_state['외상매출'] = f1.read()
if f2: st.session_state['청구'] = f2.read()

# ── 옵션 ────────────────────────────────────────────────────
st.markdown("### ⚙️ 옵션")
opt1, opt2 = st.columns(2)
with opt1:
    target_ym = st.text_input("📅 청구일 필터 (예: 26/04)", value="26/04",
                              help="청구.XLS에서 이 prefix로 시작하는 청구일만 합산")
with opt2:
    threshold = st.number_input("⚖️ 차이 임계값 (원)", value=1000, step=100,
                                help="이 금액 미만 차이는 '일치'로 처리")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 비교 시작", type="primary", use_container_width=True):
    if not st.session_state['외상매출'] or not st.session_state['청구']:
        st.warning("⚠️ 두 파일 다 업로드해주세요")
        st.stop()

    with st.spinner("분석 중..."):
        try:
            # 1. 외상매출
            df_ar = read_tsv_cp949(st.session_state['외상매출'])
            df_ar = df_ar.dropna(subset=['거래처명'])
            df_ar['금월매출_n'] = df_ar['금월매출'].apply(to_num)
            df_ar['norm'] = df_ar['거래처명'].apply(clean)
            df_ar = df_ar[df_ar['norm'] != '']
            ar_agg = df_ar.groupby('norm').agg(
                거래처=('거래처명','first'),
                담당=('영업담당','first'),
                구분=('거래처구분','first'),
                금월매출=('금월매출_n','sum'),
            ).reset_index()

            # 2. 청구 (4월만)
            df_bill = read_tsv_cp949(st.session_state['청구'])
            df_bill = df_bill.dropna(subset=['거래처'])
            df_bill['금액_n'] = df_bill['금액'].apply(to_num)
            df_bill['norm'] = df_bill['거래처'].apply(clean)
            df_bill = df_bill[df_bill['norm'] != '']
            df_bill = df_bill[df_bill['청구일'].astype(str).str.startswith(target_ym)]
            bill_agg = df_bill.groupby('norm').agg(
                청구거래처=('거래처','first'),
                청구금액=('금액_n','sum'),
                건수=('금액_n','size'),
            ).reset_index()

            # 3. 머지
            merged = ar_agg.merge(bill_agg, on='norm', how='outer')
            for col, default in [('거래처',''),('담당',''),('구분',''),
                                 ('청구거래처',''),('금월매출',0),('청구금액',0),('건수',0)]:
                merged[col] = merged[col].fillna(default)
            merged['차이'] = merged['금월매출'] - merged['청구금액']

            # 표시 거래처명
            def disp_name(r):
                a, b = str(r['거래처']).strip(), str(r['청구거래처']).strip()
                return a if a and a != '0' else b
            merged['표시명'] = merged.apply(disp_name, axis=1)

            # 분류
            def classify(r):
                ar = r['금월매출']; bl = r['청구금액']; d = r['차이']
                if ar != 0 and bl == 0: return "🚨 외상만 있음 (청구 누락 의심)"
                if ar == 0 and bl != 0: return "⚠️ 청구만 있음 (외상X)"
                if d > threshold: return "🚨 외상 > 청구"
                if d < -threshold: return "⚠️ 청구 > 외상"
                return "✅ 일치"
            merged['상태'] = merged.apply(classify, axis=1)

            # 이슈만
            issues = merged[merged['상태'] != '✅ 일치'].copy()
            issues['절대차이'] = issues['차이'].abs()
            issues = issues.sort_values('절대차이', ascending=False).reset_index(drop=True)

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback; st.code(traceback.format_exc()); st.stop()

    # ── 결과 ─────────────────────────────────────────────────
    st.divider()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📚 외상매출 거래처", f"{len(ar_agg)}곳", f"{ar_agg['금월매출'].sum():,.0f}")
    c2.metric("📄 청구 거래처", f"{len(bill_agg)}곳", f"{bill_agg['청구금액'].sum():,.0f}")
    c3.metric("🔴 차이 거래처", f"{len(issues)}곳", delta=f"-{len(issues)}", delta_color="inverse")
    c4.metric("💰 차이 합계", f"{issues['차이'].sum():,.0f}원")

    st.subheader(f"🔍 차이 발견 거래처 ({len(issues)}곳)")

    if len(issues) > 0:
        # 상태별 카운트
        status_counts = issues['상태'].value_counts()
        cols = st.columns(len(status_counts))
        for col, (status, count) in zip(cols, status_counts.items()):
            col.metric(status, f"{count}곳")

        df_show = pd.DataFrame([{
            "No.": i+1,
            "거래처": r["표시명"],
            "담당": str(r["담당"]) if r["담당"] else "-",
            "구분": str(r["구분"]) if r["구분"] else "-",
            "외상매출(금월)": fmt(r["금월매출"]),
            "청구금액": fmt(r["청구금액"]),
            "청구건수": int(r["건수"]) if r["건수"] else 0,
            "차이": fmt(r["차이"]),
            "상태": r["상태"],
        } for i, r in issues.iterrows()])

        def hl(row):
            color = ""
            if "🚨" in str(row['상태']): color = "#FFCDD2"
            elif "⚠️" in str(row['상태']): color = "#FFF3E0"
            return [f'background-color:{color}' if color else ''] * len(row)

        st.dataframe(df_show.style.apply(hl, axis=1),
                    use_container_width=True, hide_index=True, height=600)

        st.markdown("""
        **🚨 빨강 = 외상매출엔 잡혔는데 청구가 안 됐거나 적음 → 청구 누락 의심**  
        **⚠️ 노랑 = 청구는 됐는데 외상매출에 안 잡힘 → 카드매출/현금영수증/선매출 가능성**
        """)
    else:
        st.success("🎉 차이 있는 거래처 없음!")

    # ── Excel 다운로드 ───────────────────────────────────────
    if len(issues) > 0:
        wb = Workbook()
        ws = wb.active; ws.title = "차이거래처"
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

        headers = ["No.","거래처","담당","구분","외상매출(금월)","청구금액","청구건수","차이","상태"]
        ws['A1'] = f"외상매출 vs 청구 차이 검출 ({target_ym})"
        ws['A1'].font = Font(bold=True, size=13, name="맑은 고딕")
        ws.merge_cells('A1:I1')

        for ci, h in enumerate(headers, 1):
            c = ws.cell(2, ci, h)
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border

        num_cols = {5, 6, 7, 8}
        for i, r in enumerate(issues.itertuples(), 1):
            vals = [i, r.표시명, str(r.담당) or '-', str(r.구분) or '-',
                    int(r.금월매출), int(r.청구금액), int(r.건수) if r.건수 else 0,
                    int(r.차이), r.상태]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(i+2, ci, val)
                c.font = Font(name="맑은 고딕", size=10)
                c.border = border
                c.alignment = Alignment(
                    horizontal='right' if ci in num_cols else ('center' if ci in (1,7,9) else 'left'),
                    vertical='center')
                if ci in num_cols: c.number_format = '#,##0'
                if '🚨' in str(r.상태):
                    c.fill = PatternFill("solid", fgColor="FFCDD2")
                elif '⚠️' in str(r.상태):
                    c.fill = PatternFill("solid", fgColor="FFF3E0")

        widths = [6, 32, 10, 14, 16, 16, 10, 14, 28]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64+ci)].width = w
        for r in range(2, len(issues)+3):
            ws.row_dimensions[r].height = 22

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button(
            "⬇️ 결과 Excel 다운로드",
            data=buf,
            file_name=f"외상매출_청구_차이_{target_ym.replace('/','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    # ── 전체 보기 (옵션) ─────────────────────────────────────
    with st.expander("📋 전체 거래처 보기 (일치 포함)"):
        merged_show = merged.sort_values('차이', key=abs, ascending=False)
        df_all = pd.DataFrame([{
            "거래처": r["표시명"],
            "담당": str(r["담당"]) if r["담당"] else "-",
            "외상매출": fmt(r["금월매출"]),
            "청구": fmt(r["청구금액"]),
            "차이": fmt(r["차이"]),
            "상태": r["상태"],
        } for _, r in merged_show.iterrows()])
        st.dataframe(df_all, use_container_width=True, hide_index=True, height=500)
