import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="유창강건 마감 킬러", layout="wide")
st.title("📊 유창강건 세금계산서 누락 체크기")

def read_excel_any(file):
    """xls/xlsx/csv 어떤 형식이든 읽기"""
    file.seek(0)
    raw = file.read()
    # 1) EUC-KR 텍스트(xls 위장 CSV) 시도
    try:
        text = raw.decode('euc-kr')
        lines = text.strip().replace('\r\n', '\n').split('\n')
        return pd.DataFrame({0: lines})
    except:
        pass
    # 2) xlrd (구형 xls)
    try:
        file.seek(0)
        return pd.read_excel(file, header=None, engine='xlrd')
    except:
        pass
    # 3) openpyxl (xlsx)
    try:
        file.seek(0)
        return pd.read_excel(file, header=None, engine='openpyxl')
    except:
        pass
    raise ValueError("파일을 읽을 수 없습니다.")
st.info("물품출고(ERP), 카드매출, 세금계산서 발행목록을 대조합니다.")

# ── 이름 정제 함수 ──────────────────────────────────────────
def clean(x):
    x = str(x)
    x = re.sub(r'[■▲▶●★☆□△◆◇]', '', x)   # 특수문자 제거
    x = x.replace("(주)", "").replace("(유)", "").replace("(株)", "")
    x = re.sub(r'\s+', '', x)               # 공백 전부 제거
    return x.strip()

def is_number(s):
    s = str(s).strip().strip('"').replace(',', '').replace('(', '').replace(')', '')
    try:
        float(s)
        return True
    except:
        return False

# ── 파일 업로드 ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("#### 1️⃣ 물품출고(ERP)")
    st.caption("거래처명이 한 열에 있는 xls/xlsx/csv")
    file_out = st.file_uploader("ERP 파일", type=['xlsx', 'xls', 'csv'], label_visibility="collapsed")

with col2:
    st.markdown("#### 2️⃣ 카드매출 비교")
    st.caption("A열에 거래처명이 있는 xlsx")
    file_card = st.file_uploader("카드매출 파일", type=['xlsx', 'xls', 'csv'], label_visibility="collapsed")

with col3:
    st.markdown("#### 3️⃣ 세금계산서 발행목록")
    st.caption("국세청 전자세금계산서 목록")
    file_tax = st.file_uploader("세금계산서 파일", type=['xlsx', 'xls', 'csv'], label_visibility="collapsed")

# ── 분석 ────────────────────────────────────────────────────
if st.button("🚀 미발행 업체 분석 시작", type="primary", use_container_width=True):
    if not (file_out and file_card and file_tax):
        st.warning("⚠️ 파일 3개를 모두 올려주세요.")
        st.stop()

    with st.spinner("분석 중..."):
        try:
            # ── 1. ERP 물품출고 파일 ──────────────────────────
            # EUC-KR CSV 형태(xls 위장) or 정상 xlsx 모두 처리
            try:
                raw = file_out.read()
                text = raw.decode('euc-kr')
                lines = text.strip().split('\r\n') if '\r\n' in text else text.strip().split('\n')
                erp_names = set()
                for line in lines[1:]:  # 헤더 제외
                    val = line.strip().strip('"')
                    if val and not is_number(val):
                        erp_names.add(clean(val))
            except Exception:
                file_out.seek(0)
                df_erp = pd.read_excel(file_out, header=None)
                erp_names = set()
                for val in df_erp.iloc[1:, 0].dropna():
                    if not is_number(str(val)):
                        erp_names.add(clean(str(val)))

            # ── 2. 카드매출 파일 ─────────────────────────────
            file_card.seek(0)
            df_card = pd.read_excel(file_card, header=None)
            card_names = set()
            for val in df_card.iloc[3:, 0].dropna():   # row 0~2 = 헤더
                v = clean(str(val))
                if v:
                    card_names.add(v)

            # ── 3. 세금계산서 파일 ───────────────────────────
            # 국세청 양식: 5번째 행(index=5)이 헤더, 공급받는자 상호 = 12번째 열(index=11)
            file_tax.seek(0)
            df_tax = pd.read_excel(file_tax, header=None)
            tax_names = set()
            for val in df_tax.iloc[6:, 11].dropna():   # 데이터는 row 6부터
                tax_names.add(clean(str(val)))

            # ── 4. 비교 로직 ─────────────────────────────────
            # 카드매출 제외한 발행 대상
            target = erp_names - card_names
            # 실제 미발행
            missing_keys = target - tax_names
            # 원본 이름 복원 (clean 전 이름 보여주기)
            # ERP 원본이름 dict 만들기
            try:
                raw2 = file_out.read()
                text2 = raw2.decode('euc-kr')
                lines2 = text2.strip().split('\r\n') if '\r\n' in text2 else text2.strip().split('\n')
                original_map = {}
                for line in lines2[1:]:
                    val = line.strip().strip('"')
                    if val and not is_number(val):
                        original_map[clean(val)] = val
            except Exception:
                original_map = {k: k for k in erp_names}

            missing_display = sorted([
                original_map.get(k, k) for k in missing_keys
            ])

            # 특수 항목 필터 (기타거래처, 현금영수증 등)
            skip_keywords = ['기타거래처', '현금영수증', '카드매출']
            missing_display = [
                m for m in missing_display
                if not any(kw in m for kw in skip_keywords)
            ]

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            st.stop()

    # ── 결과 표시 ────────────────────────────────────────────
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체 매출 업체", f"{len(erp_names)}개")
    m2.metric("카드매출 업체", f"{len(card_names)}개")
    m3.metric("세금계산서 발행", f"{len(tax_names)}개")
    m4.metric("⚠️ 미발행 업체", f"{len(missing_display)}개", delta=f"-{len(missing_display)}", delta_color="inverse")

    st.subheader(f"✅ 세금계산서 미발행 업체 ({len(missing_display)}개)")

    if missing_display:
        df_result = pd.DataFrame({
            "No.": range(1, len(missing_display)+1),
            "업체명": missing_display
        })
        st.dataframe(df_result, use_container_width=True, hide_index=True)

        # ── 엑셀 다운로드 ────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = "미발행업체"
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        ws['A1'] = '세금계산서 미발행 업체 목록'
        ws['A1'].font = Font(bold=True, size=14, name="맑은 고딕")
        ws.merge_cells('A1:B1')
        ws['A2'] = f'미발행: {len(missing_display)}개'
        ws['A2'].font = Font(size=10, name="맑은 고딕", color="555555")
        ws.merge_cells('A2:B2')
        for col, title in [('A', 'No.'), ('B', '업체명')]:
            c = ws[f'{col}3']
            c.value = title
            c.fill = PatternFill("solid", fgColor="CC2222")
            c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕")
            c.alignment = Alignment(horizontal='center')
            c.border = border
        for i, name in enumerate(missing_display, 1):
            ws[f'A{i+3}'] = i
            ws[f'B{i+3}'] = name
            for col in ['A', 'B']:
                c = ws[f'{col}{i+3}']
                c.font = Font(name="맑은 고딕")
                c.border = border
                c.alignment = Alignment(horizontal='left' if col == 'B' else 'center')
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
