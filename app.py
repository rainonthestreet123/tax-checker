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

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def find_similar(name_clean, tax_raw_list, threshold=0.75):
    """세금계산서에서 유사한 이름 찾기 → (원본이름, 행번호, 유사도)"""
    results = []
    for row_idx, raw_name in tax_raw_list:
        raw_clean = clean(raw_name)
        score = similarity(name_clean, raw_clean)
        if score >= threshold and name_clean != raw_clean:
            results.append((raw_name, row_idx, score))
    results.sort(key=lambda x: -x[2])
    return results[:3]  # 상위 3개만

# ── 파일 업로드 ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("#### 1️⃣ 물품출고(ERP)")
    st.caption("유창강건 ERP에서 내보낸 xls")
    file_out = st.file_uploader("ERP 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
with col2:
    st.markdown("#### 2️⃣ 카드매출 비교")
    st.caption("월별 시트가 있는 xlsx")
    file_card = st.file_uploader("카드매출 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
with col3:
    st.markdown("#### 3️⃣ 세금계산서 발행목록")
    st.caption("국세청 전자세금계산서 목록 xls")
    file_tax = st.file_uploader("세금계산서 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")

# ── 카드매출 시트 선택 ───────────────────────────────────────
selected_sheet = None
if file_card:
    try:
        file_card.seek(0)
        sheet_names = pd.ExcelFile(file_card, engine='openpyxl').sheet_names
        selected_sheet = st.selectbox(
            "📅 카드매출 분석 월 선택",
            options=sheet_names,
            index=0,
            help="세금계산서와 대조할 월을 선택하세요"
        )
        st.caption(f"선택된 시트: **{selected_sheet}**")
    except Exception as e:
        st.error(f"카드매출 파일 시트 읽기 오류: {e}")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 미발행 업체 분석 시작", type="primary", use_container_width=True):
    if not (file_out and file_card and file_tax):
        st.warning("⚠️ 파일 3개를 모두 올려주세요.")
        st.stop()
    if not selected_sheet:
        st.warning("⚠️ 카드매출 월을 선택해주세요.")
        st.stop()

    with st.spinner("분석 중..."):
        try:
            # ── 1. ERP ───────────────────────────────────────
            file_out.seek(0)
            raw = file_out.read()
            text = raw.decode('cp949')
            df_erp = pd.read_csv(StringIO(text), sep='\t', header=0, on_bad_lines='skip')
            col_name = '거래처명' if '거래처명' in df_erp.columns else df_erp.columns[3]
            erp_names = set()
            original_map = {}
            for v in df_erp[col_name].dropna():
                v = str(v).strip().strip('"')
                if v and not is_number(v):
                    k = clean(v)
                    erp_names.add(k)
                    original_map[k] = v

            # ── 2. 카드매출 ───────────────────────────────────
            file_card.seek(0)
            df_card = pd.read_excel(file_card, sheet_name=selected_sheet, header=None, engine='openpyxl')
            card_names = set()
            for v in df_card.iloc[3:, 0].dropna():
                k = clean(str(v))
                if k:
                    card_names.add(k)

            # ── 3. 세금계산서 ─────────────────────────────────
            file_tax.seek(0)
            raw_tax = file_tax.read()
            is_xls = raw_tax[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
            file_tax.seek(0)
            df_tax = pd.read_excel(file_tax, header=None,
                                   engine='xlrd' if is_xls else 'openpyxl')

            tax_names = set()
            # 행번호 포함해서 저장 (유사도 검색용)
            tax_raw_list = []   # [(행번호, 원본이름), ...]
            for i, v in df_tax.iloc[6:, 11].items():
                if pd.notna(v) and str(v).strip():
                    tax_names.add(clean(str(v)))
                    tax_raw_list.append((i - 5, str(v).strip()))  # 행번호는 헤더 제외 기준

            # ── 4. 비교 + 유사도 검색 ─────────────────────────
            target = erp_names - card_names
            missing_keys = target - tax_names
            skip_keywords = ['기타거래처', '현금영수증', '카드매출', '거래처명', '거래처ID']

            rows = []
            for k in missing_keys:
                orig = original_map.get(k, k)
                if any(kw in orig for kw in skip_keywords):
                    continue

                # 유사 이름 검색
                similars = find_similar(k, tax_raw_list, threshold=0.72)
                if similars:
                    notes = []
                    for sim_name, sim_row, score in similars:
                        notes.append(f"유사이름: '{sim_name}' (세금계산서 {sim_row}행, 유사도 {score:.0%})")
                    reason = " / ".join(notes)
                else:
                    reason = "세금계산서 미발행"

                rows.append({"업체명": orig, "비고": reason, "_key": k})

            rows.sort(key=lambda x: x["업체명"])

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

    # ── 결과 ─────────────────────────────────────────────────
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 매출 업체",              f"{len(erp_names)}개")
    m2.metric(f"카드매출 ({selected_sheet})", f"{len(card_names)}개")
    m3.metric("세금계산서 발행",              f"{len(tax_names)}개")
    m4.metric("⚠️ 미발행 업체",              f"{len(rows)}개",
              delta=f"-{len(rows)}", delta_color="inverse")

    st.subheader(f"✅ [{selected_sheet}] 세금계산서 미발행 업체 ({len(rows)}개)")

    if rows:
        df_result = pd.DataFrame([
            {"No.": i+1, "업체명": r["업체명"], "비고": r["비고"]}
            for i, r in enumerate(rows)
        ])

        # 유사이름 있는 행은 노란색 강조
        def highlight_row(row):
            if "유사이름" in str(row["비고"]):
                return ['background-color: #fff9c4'] * len(row)
            return [''] * len(row)

        st.dataframe(
            df_result.style.apply(highlight_row, axis=1),
            use_container_width=True,
            hide_index=True
        )

        st.caption("🟡 노란색 = 세금계산서에 유사한 이름이 있음 (오타/띄어쓰기 차이일 수 있음, 직접 확인 필요)")

        # ── 엑셀 다운로드 ─────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = "미발행업체"
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'),  bottom=Side(style='thin'))

        ws['A1'] = f'세금계산서 미발행 업체 목록 ({selected_sheet})'
        ws['A1'].font = Font(bold=True, size=14, name="맑은 고딕")
        ws.merge_cells('A1:C1')
        ws['A2'] = f'미발행: {len(rows)}개  |  🟡 노란색: 세금계산서에 유사이름 있음 (직접 확인 필요)'
        ws['A2'].font = Font(size=10, name="맑은 고딕", color="555555")
        ws.merge_cells('A2:C2')

        for col, title in [('A','No.'), ('B','업체명'), ('C','비고')]:
            c = ws[f'{col}3']
            c.value = title
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border

        yellow_fill = PatternFill("solid", fgColor="FFF9C4")
        alt_fill    = PatternFill("solid", fgColor="FFF0F0")

        for i, r in enumerate(rows, 1):
            ws[f'A{i+3}'] = i
            ws[f'B{i+3}'] = r["업체명"]
            ws[f'C{i+3}'] = r["비고"]
            has_similar = "유사이름" in r["비고"]
            for col in ['A','B','C']:
                c = ws[f'{col}{i+3}']
                c.font = Font(name="맑은 고딕", size=10)
                c.border = border
                c.alignment = Alignment(horizontal='left' if col != 'A' else 'center',
                                        vertical='center', wrap_text=(col=='C'))
                if has_similar:
                    c.fill = yellow_fill
                elif i % 2 == 0:
                    c.fill = alt_fill

        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 60
        for row in range(3, len(rows)+4):
            ws.row_dimensions[row].height = 30

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            "⬇️ 미발행 업체 엑셀 다운로드",
            data=buf,
            file_name=f"미발행업체_{selected_sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    else:
        st.success("🎉 미발행 업체가 없습니다!")
