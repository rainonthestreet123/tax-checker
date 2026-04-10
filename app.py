import streamlit as st
import pandas as pd
import re
from io import BytesIO, StringIO
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

def read_as_xls_ole(file):
    """OLE2 .xls 파일 읽기 (xlrd)"""
    file.seek(0)
    return pd.read_excel(file, header=None, engine='xlrd')

def read_as_xlsx(file):
    """xlsx 파일 읽기 (openpyxl)"""
    file.seek(0)
    return pd.read_excel(file, header=None, engine='openpyxl')

def read_erp(file):
    """ERP 물품출고 파일: CP949 인코딩 TSV"""
    file.seek(0)
    raw = file.read()
    text = raw.decode('cp949')
    df = pd.read_csv(StringIO(text), sep='\t', header=0, on_bad_lines='skip')
    return df

# ── 파일 업로드 ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("#### 1️⃣ 물품출고(ERP)")
    st.caption("유창강건 ERP에서 내보낸 xls")
    file_out = st.file_uploader("ERP 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
with col2:
    st.markdown("#### 2️⃣ 카드매출 비교")
    st.caption("A열에 거래처명이 있는 xlsx")
    file_card = st.file_uploader("카드매출 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")
with col3:
    st.markdown("#### 3️⃣ 세금계산서 발행목록")
    st.caption("국세청 전자세금계산서 목록 xls")
    file_tax = st.file_uploader("세금계산서 파일", type=['xlsx','xls','csv'], label_visibility="collapsed")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 미발행 업체 분석 시작", type="primary", use_container_width=True):
    if not (file_out and file_card and file_tax):
        st.warning("⚠️ 파일 3개를 모두 올려주세요.")
        st.stop()

    with st.spinner("분석 중..."):
        try:
            # ── 1. ERP 물품출고 (CP949 TSV) ──────────────────
            df_erp = read_erp(file_out)

            # '거래처명' 컬럼 찾기
            if '거래처명' in df_erp.columns:
                col_name = '거래처명'
            else:
                # 혹시 컬럼명이 다르면 3번째 열(index 3) 사용
                col_name = df_erp.columns[3]

            erp_names = set()
            original_map = {}
            for v in df_erp[col_name].dropna():
                v = str(v).strip().strip('"')
                if v and not is_number(v):
                    k = clean(v)
                    erp_names.add(k)
                    original_map[k] = v

            # ── 2. 카드매출 (xlsx) ────────────────────────────
            df_card = read_as_xlsx(file_card)
            card_names = set()
            for v in df_card.iloc[3:, 0].dropna():   # row 0~2 = 헤더
                k = clean(str(v))
                if k:
                    card_names.add(k)

            # ── 3. 세금계산서 (OLE2 xls) ─────────────────────
            # xlrd로 읽기 시도, 안되면 LibreOffice 변환 후 openpyxl
            try:
                df_tax = read_as_xls_ole(file_tax)
            except Exception:
                # xlrd 실패시 openpyxl 시도 (xlsx로 저장된 경우)
                df_tax = read_as_xlsx(file_tax)

            tax_names = set()
            for v in df_tax.iloc[6:, 11].dropna():   # 공급받는자 상호 = col 11
                tax_names.add(clean(str(v)))

            # ── 4. 비교 ───────────────────────────────────────
            target = erp_names - card_names
            missing_keys = target - tax_names
            skip_keywords = ['기타거래처', '현금영수증', '카드매출', '거래처명', '거래처ID']
            missing_display = sorted([
                original_map.get(k, k) for k in missing_keys
                if not any(kw in original_map.get(k, k) for kw in skip_keywords)
            ])

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

    # ── 결과 ─────────────────────────────────────────────────
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 매출 업체",    f"{len(erp_names)}개")
    m2.metric("카드매출 업체",     f"{len(card_names)}개")
    m3.metric("세금계산서 발행",   f"{len(tax_names)}개")
    m4.metric("⚠️ 미발행 업체",   f"{len(missing_display)}개",
              delta=f"-{len(missing_display)}", delta_color="inverse")

    st.subheader(f"✅ 세금계산서 미발행 업체 ({len(missing_display)}개)")

    if missing_display:
        df_result = pd.DataFrame({
            "No.": range(1, len(missing_display)+1),
            "업체명": missing_display
        })
        st.dataframe(df_result, use_container_width=True, hide_index=True)

        # 엑셀 다운로드
        wb = Workbook()
        ws = wb.active
        ws.title = "미발행업체"
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'),  bottom=Side(style='thin'))
        ws['A1'] = '세금계산서 미발행 업체 목록'
        ws['A1'].font = Font(bold=True, size=14, name="맑은 고딕")
        ws.merge_cells('A1:B1')
        ws['A2'] = f'미발행: {len(missing_display)}개'
        ws['A2'].font = Font(size=10, name="맑은 고딕", color="555555")
        ws.merge_cells('A2:B2')
        for col, title in [('A','No.'), ('B','업체명')]:
            c = ws[f'{col}3']
            c.value = title
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border
        for i, name in enumerate(missing_display, 1):
            ws[f'A{i+3}'] = i
            ws[f'B{i+3}'] = name
            for col in ['A','B']:
                c = ws[f'{col}{i+3}']
                c.font = Font(name="맑은 고딕")
                c.border = border
                c.alignment = Alignment(horizontal='left' if col=='B' else 'center')
                if i % 2 == 0:
                    c.fill = PatternFill("solid", fgColor="FFF0F0")
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 45

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            "⬇️ 미발행 업체 엑셀 다운로드",
            data=buf,
            file_name="미발행업체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    else:
        st.success("🎉 미발행 업체가 없습니다!")
