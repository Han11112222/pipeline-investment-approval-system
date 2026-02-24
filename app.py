import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

# --------------------------------------------------------------------------
# [설정] 페이지 기본
# --------------------------------------------------------------------------
st.set_page_config(page_title="공식 배관 투자 결재 시스템 (Pipeline Approval)", page_icon="🏗️", layout="wide")

# --------------------------------------------------------------------------
# [함수] 금융 계산 로직 (전산 시스템과 대조할 순수 오리지널 로직)
# --------------------------------------------------------------------------
def manual_npv(rate, values):
    return sum(v / ((1 + rate) ** i) for i, v in enumerate(values))

def calculate_simulation(sim_len, sim_inv, sim_contrib, sim_other, sim_vol, sim_rev, sim_cost, 
                         sim_jeon, sim_basic_rev, rate, tax, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m):
    
    # 1. 초기 순투자액
    net_inv = sim_inv - sim_contrib - sim_other
    
    # 2. 고정 수익/비용 항목 계산
    margin_total = (sim_rev - sim_cost) + sim_basic_rev 
    
    cost_sga = (sim_len * c_maint) + (sim_len * c_adm_m) + (sim_jeon * c_adm_jeon)
    annual_depreciation = sim_inv / dep_period if dep_period > 0 else 0
    
    # 3. 세후 현금흐름(OCF) 산출
    ebit = margin_total - cost_sga - annual_depreciation
    net_income = ebit * (1 - tax) 
    fixed_ocf = net_income + annual_depreciation
    
    flows = [-net_inv]
    ocfs = []
    
    for year in range(1, int(analysis_period) + 1):
        flows.append(fixed_ocf)
        ocfs.append(fixed_ocf)

    npv_val = manual_npv(rate, flows)
    irr_val = None
    irr_reason = ""
    
    if net_inv <= 0:
        irr_reason = "초기 투자비 ≤ 0"
    elif all(f <= 0 for f in ocfs): 
        irr_reason = "운영 적자 지속"
    else:
        try:
            irr_val = npf.irr(flows)
        except:
            irr_reason = "계산 오류"
            
    return npv_val, irr_val, irr_reason, flows

def get_col_idx(df, keywords, exact=False):
    for col_idx in range(df.shape[1]):
        for row_idx in range(min(20, df.shape[0])):
            val = str(df.iloc[row_idx, col_idx]).replace(" ", "").replace("\n", "")
            for kw in keywords:
                if exact:
                    if val == kw: return col_idx
                else:
                    if kw in val: return col_idx
    return None

# --------------------------------------------------------------------------
# [UI] 좌측 사이드바 탭 메뉴 구성
# --------------------------------------------------------------------------
st.sidebar.title("메뉴 네비게이션")
menu_choice = st.sidebar.radio(
    "이동할 페이지를 선택하세요:",
    ('1. 배관 투자 경제성 결재 대시보드', '2. 배관 투자 승인 내역')
)
st.sidebar.divider()


# ==========================================================================
# 탭 1: 기존 배관 투자 경제성 결재 대시보드
# ==========================================================================
if menu_choice == '1. 배관 투자 경제성 결재 대시보드':

    # [UI] 탭 1 전용 사이드바 변수 설정
    with st.sidebar:
        st.header("⚙️ 분석 변수 설정")
        st.info("💡 전산 시스템의 NPV와 일치하도록 아래 단가와 세율을 시스템 세팅값과 동일하게 맞춰주세요.")
        
        rate_pct = st.number_input("할인율 (%)", value=6.15, step=0.01, format="%.2f")
        tax_pct = st.number_input("법인세율+주민세율 (%)", value=22.0, step=0.1, format="%.1f")
        dep_period = st.number_input("감가상각 연수 (년)", value=30, step=1)
        analysis_period = st.number_input("경제성 분석 연수 (년)", value=30, step=1)
        
        st.subheader("💰 비용 단가")
        c_maint = st.number_input("유지비 (원/m)", value=8222, format="%d")
        c_adm_jeon = st.number_input("관리비 (원/전)", value=6209, format="%d")
        c_adm_m = st.number_input("관리비 (원/m)", value=13605, format="%d")
        sim_basic_price = st.number_input("주택용 월 기본요금 단가 (원)", value=900, step=10, format="%d")

        RATE = rate_pct / 100
        TAX = tax_pct / 100

    # [UI] 탭 1 메인 화면
    st.title("🏗️ 배관 투자 경제성 결재 대시보드")
    st.markdown("전산 시스템 Raw 데이터를 업로드하여 경제성을 시뮬레이션합니다. **분석에서 제외할 항목은 체크 해제**하세요.")

    uploaded_file = st.file_uploader("📂 전산 Raw 데이터 파일 업로드 (Excel 또는 CSV)", type=['xlsx', 'xls', 'csv'])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, header=None) 
            else:
                df = pd.read_excel(uploaded_file, header=None)
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            st.stop()

        idx_usage = get_col_idx(df, ["용도", "가스용도"], exact=True)
        idx_name = get_col_idx(df, ["구간명"], exact=True)
        idx_len = get_col_idx(df, ["길이(m)", "배관길이"], exact=False)
        idx_inv = get_col_idx(df, ["배관투자금액", "총공사비"], exact=False)
        idx_contrib = get_col_idx(df, ["시설분담금"], exact=False)
        idx_other = get_col_idx(df, ["기타이익", "보조금"], exact=False)
        idx_jeon = get_col_idx(df, ["수요전수계", "총전수"], exact=False)
        idx_jeon_apt = get_col_idx(df, ["공동주택전수"], exact=False)
        idx_jeon_single = get_col_idx(df, ["단독주택전수"], exact=False)
        idx_vol = get_col_idx(df, ["계(MJ)"], exact=False) 
        idx_rev = get_col_idx(df, ["연간판매액", "판매액"], exact=False)
        idx_cost = get_col_idx(df, ["연간판매원가", "판매원가"], exact=False)

        if idx_name is None:
            st.error("❌ 데이터에서 '구간명'의 위치를 찾을 수 없습니다.")
            st.stop()

        mapped_data = {}
        mapped_data['용도'] = df.iloc[:, idx_usage] if idx_usage is not None else '미분류'
        mapped_data['구간명'] = df.iloc[:, idx_name]
        mapped_data['길이'] = df.iloc[:, idx_len] if idx_len is not None else 0
        mapped_data['투자비'] = df.iloc[:, idx_inv] if idx_inv is not None else 0
        mapped_data['분담금'] = df.iloc[:, idx_contrib] if idx_contrib is not None else 0
        mapped_data['기타이익'] = df.iloc[:, idx_other] if idx_other is not None else 0
        mapped_data['총전수'] = df.iloc[:, idx_jeon] if idx_jeon is not None else 0
        mapped_data['공동주택전수'] = df.iloc[:, idx_jeon_apt] if idx_jeon_apt is not None else 0
        mapped_data['단독주택전수'] = df.iloc[:, idx_jeon_single] if idx_jeon_single is not None else 0
        mapped_data['판매량'] = df.iloc[:, idx_vol] if idx_vol is not None else 0
        mapped_data['판매액'] = df.iloc[:, idx_rev] if idx_rev is not None else 0
        mapped_data['판매원가'] = df.iloc[:, idx_cost] if idx_cost is not None else 0

        clean_df = pd.DataFrame(mapped_data)

        clean_df['구간명'] = clean_df['구간명'].astype(str).str.strip()
        invalid_names = ['', '0', 'nan', 'None', '구간명']
        clean_df = clean_df[~clean_df['구간명'].isin(invalid_names)]
        clean_df['용도'] = clean_df['용도'].astype(str).str.strip().ffill().fillna('미분류')

        num_cols_base = ['길이', '투자비', '분담금', '기타이익', '총전수', '공동주택전수', '단독주택전수', '판매량', '판매액', '판매원가']
        for c in num_cols_base:
            if clean_df[c].dtype == object:
                clean_df[c] = clean_df[c].astype(str).str.replace(',', '', regex=False)
            clean_df[c] = pd.to_numeric(clean_df[c], errors='coerce').fillna(0)

        clean_df['총전수'] = np.maximum(clean_df['총전수'], clean_df['공동주택전수'] + clean_df['단독주택전수'])
        
        clean_df['기본요금수익'] = 0.0
        temp_usage = clean_df['용도'].astype(str).str.replace(' ', '', regex=False)
        is_home = temp_usage.str.contains('주택|가정|공동') & ~temp_usage.str.contains('외')
        clean_df.loc[is_home, '기본요금수익'] = clean_df.loc[is_home, '총전수'] * sim_basic_price * 12
        
        num_cols = ['길이', '투자비', '분담금', '기타이익', '총전수', '판매량', '판매액', '판매원가', '기본요금수익']

        st.success("✅ 전산 파일 업로드 완료! 사이드바의 변수를 전산 시스템과 동일하게 세팅해 보세요.")

        def get_analysis_result(row):
            npv, irr, irr_msg, _ = calculate_simulation(
                row['길이'], row['투자비'], row['분담금'], row['기타이익'], row['판매량'], row['판매액'], row['판매원가'], 
                row['총전수'], row['기본요금수익'], RATE, TAX, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m
            )
            return row['길이'], row['투자비'] - row['분담금'] - row['기타이익'], row['판매량'], npv, irr, irr_msg

        # ---------------------------------------------------------
        # UI Section 1: 요약 표
        # ---------------------------------------------------------
        st.subheader("1. 📁 용도별 경제성 요약 (분석 대상 선택)")
        
        usage_results = []
        for u in clean_df['용도'].unique():
            u_df = clean_df[clean_df['용도'] == u]
            u_len, u_net_inv, u_vol, u_npv, u_irr, u_irr_msg = get_analysis_result(u_df[num_cols].sum())
            is_selected = False if 'ROE' in str(u).upper() else True
            
            usage_results.append({
                "선택": is_selected,
                "용도": u,
                "총 투자길이(m)": float(u_len),
                "총 순투자액(원)": float(u_net_inv),
                "연간판매량(MJ)": float(u_vol),
                "NPV(원)": float(u_npv),
                "IRR(%)": float(u_irr*100) if u_irr is not None else None
            })
            
        df_usage_summary = pd.DataFrame(usage_results)

        edited_df = st.data_editor(
            df_usage_summary,
            column_config={
                "선택": st.column_config.CheckboxColumn("선택"),
                "총 투자길이(m)": st.column_config.NumberColumn(format="%.1f"),
                "총 순투자액(원)": st.column_config.NumberColumn(format="%.0f"),
                "연간판매량(MJ)": st.column_config.NumberColumn(format="%.0f"),
                "NPV(원)": st.column_config.NumberColumn(format="%.0f"),
                "IRR(%)": st.column_config.NumberColumn(format="%.2f")
            },
            disabled=["용도", "총 투자길이(m)", "총 순투자액(원)", "연간판매량(MJ)", "NPV(원)", "IRR(%)"],
            hide_index=True,
            use_container_width=True
        )

        # ---------------------------------------------------------
        # 필터링 및 소계/상세 계산
        # ---------------------------------------------------------
        selected_usages = edited_df[edited_df['선택'] == True]['용도'].tolist()

        if selected_usages:
            filtered_df = clean_df[clean_df['용도'].isin(selected_usages)]
            t_len, t_net_inv, t_vol, tot_npv, tot_irr, tot_irr_msg = get_analysis_result(filtered_df[num_cols].sum())

            st.subheader("2. 📊 선택 항목 합산 소계 (Subtotal)")
            m1, m2 = st.columns(2)
            m1.metric("최종 합산 NPV", f"{tot_npv:,.0f} 원")
            m2.metric("최종 합산 IRR", f"{tot_irr*100:.2f} %" if tot_irr is not None else tot_irr_msg)
            
            subtotal_df = pd.DataFrame([{
                "항목명": "☑️ 선택 용도 총합계",
                "총 투자길이(m)": t_len, "총 순투자액(원)": t_net_inv, "연간판매량(MJ)": t_vol, "NPV(원)": tot_npv
            }])
            st.dataframe(subtotal_df.style.format({"{:,.1f}": "총 투자길이(m)", "총 순투자액(원)": "{:,.0f}", "연간판매량(MJ)": "{:,.0f}", "NPV(원)": "{:,.0f}"}), hide_index=True)

            st.divider()

            st.subheader("3. 📑 구간별 경제성 상세 명세서")
            df_detail = filtered_df.groupby(['용도', '구간명'])[num_cols].sum().reset_index()
            detail_results = []
            for _, row in df_detail.iterrows():
                d_len, d_net_inv, d_vol, d_npv, d_irr, d_irr_msg = get_analysis_result(row)
                detail_results.append({
                    "용도": row['용도'], "구간명": row['구간명'], "투자길이(m)": d_len,
                    "공급전수(전)": row['총전수'], "기본요금수익(원)": row['기본요금수익'], 
                    "순투자액(원)": d_net_inv, "연간판매량(MJ)": d_vol, "NPV(원)": d_npv
                })
                
            st.dataframe(pd.DataFrame(detail_results).style.format({
                "투자길이(m)": "{:,.1f}", "공급전수(전)": "{:,.0f}", "기본요금수익(원)": "{:,.0f}",
                "순투자액(원)": "{:,.0f}", "연간판매량(MJ)": "{:,.0f}", "NPV(원)": "{:,.0f}"
            }), use_container_width=True, hide_index=True)
    else:
        st.info("👆 분석을 시작하려면 전산 Raw 파일을 올려주세요.")


# ==========================================================================
# 탭 2: 신규 배관 투자 승인 내역 자동화
# ==========================================================================
elif menu_choice == '2. 배관 투자 승인 내역':
    st.title("📋 배관 투자 승인 내역 자동 생성기")
    st.markdown("기초자료(1차, 2차) 엑셀/CSV 파일을 업로드하면 품의서 양식에 맞게 데이터를 자동 병합합니다.")

    uploaded_files = st.file_uploader("기초자료 파일을 여러 개 업로드 해주세요", accept_multiple_files=True, type=['csv', 'xlsx', 'xls'])

    if uploaded_files:
        all_data = []
        for file in uploaded_files:
            try:
                # 엑셀 상단의 병합된 불필요한 행을 건너뛰고 3번째 행(인덱스 2)부터 읽기
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, skiprows=2, encoding='utf-8-sig') 
                else:
                    df = pd.read_excel(file, skiprows=2)

                # 컬럼명에 섞여 있는 공백 제거 (유연한 매칭을 위함)
                df.columns = df.columns.astype(str).str.replace(" ", "")

                # 포함할 주요 키워드로 컬럼 찾기
                def find_col(keyword):
                    for col in df.columns:
                        if keyword in col:
                            return col
                    return None

                extracted = pd.DataFrame()
                
                col_name = find_col('구간명')
                col_inv = find_col('배관투자금액')
                col_home = find_col('가정용')
                col_general = find_col('일반용')
                col_total_vol = find_col('계(MJ)')
                col_npv = find_col('NPV')
                col_irr = find_col('IRR')

                if col_name:
                    extracted['공사명'] = df[col_name]
                    extracted['투자비(원)'] = df[col_inv] if col_inv else 0
                    extracted['가정용 판매량(MJ)'] = df[col_home] if col_home else 0
                    extracted['일반용 판매량(MJ)'] = df[col_general] if col_general else 0
                    extracted['합계 판매량(MJ)'] = df[col_total_vol] if col_total_vol else 0
                    extracted['NPV(원)'] = df[col_npv] if col_npv else 0
                    extracted['IRR(%)'] = df[col_irr] if col_irr else 0

                    # 공사명이 없는 빈 데이터 제거 및 헤더 잔재 정리
                    extracted = extracted.dropna(subset=['공사명'])
                    extracted = extracted[~extracted['공사명'].isin(['구간명', 'nan', ''])]
                    
                    all_data.append(extracted)

            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다 ({file.name}): {e}")
        
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            final_df.index = final_df.index + 1  # No. 용도로 1부터 시작
            
            st.success("✅ 여러 기초자료 데이터의 병합이 완료되었습니다!")
            
            # 숫자 포맷 깔끔하게 지정해서 출력
            st.dataframe(final_df.style.format({
                "투자비(원)": "{:,.0f}",
                "가정용 판매량(MJ)": "{:,.0f}",
                "일반용 판매량(MJ)": "{:,.0f}",
                "합계 판매량(MJ)": "{:,.0f}",
                "NPV(원)": "{:,.0f}",
                "IRR(%)": "{:,.2f}"
            }), use_container_width=True)
            
            # CSV 다운로드 버튼
            csv_data = final_df.to_csv(index=True).encode('utf-8-sig')
            st.download_button(
                label="📥 취합된 승인내역 CSV 다운로드",
                data=csv_data,
                file_name="배관투자_승인내역_최종.csv",
                mime="text/csv"
            )
