import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import re

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
    
    net_inv = sim_inv - sim_contrib - sim_other
    margin_total = (sim_rev - sim_cost) + sim_basic_rev 
    
    cost_sga = (sim_len * c_maint) + (sim_len * c_adm_m) + (sim_jeon * c_adm_jeon)
    annual_depreciation = sim_inv / dep_period if dep_period > 0 else 0
    
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
# [UI] 공통 사이드바 (파일 업로드 및 메뉴)
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("📂 데이터 업로드 (공통)")
    st.markdown("여기서 업로드한 파일은 양쪽 탭 모두에 적용됩니다.")
    uploaded_files = st.file_uploader("기초자료 파일 업로드 (*차 다중 선택)", accept_multiple_files=True, type=['csv', 'xlsx', 'xls'])
    
    st.divider()

    st.title("메뉴 네비게이션")
    menu_choice = st.radio(
        "이동할 페이지를 선택하세요:",
        ('1. 배관 투자 경제성 결재 대시보드', '2. 배관 투자 승인 내역')
    )
    st.divider()

# ==========================================================================
# 탭 1: 기존 배관 투자 경제성 결재 대시보드
# ==========================================================================
if menu_choice == '1. 배관 투자 경제성 결재 대시보드':

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

    st.title("🏗️ 배관 투자 경제성 결재 대시보드")
    st.markdown("전산 시스템 Raw 데이터를 업로드하여 경제성을 시뮬레이션합니다. **분석에서 제외할 항목은 체크 해제**하세요.")

    if uploaded_files:
        clean_df_list = []
        for file in uploaded_files:
            try:
                match = re.search(r'(\d+)차', file.name)
                cha_num = int(match.group(1)) if match else 1

                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, header=None) 
                else:
                    df = pd.read_excel(file, header=None)

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
                    continue

                mapped_data = {}
                mapped_data['차수'] = cha_num
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

                temp_clean_df = pd.DataFrame(mapped_data)
                clean_df_list.append(temp_clean_df)

            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다 ({file.name}): {e}")
        
        if not clean_df_list:
            st.error("분석 가능한 유효 데이터가 없습니다.")
            st.stop()

        clean_df = pd.concat(clean_df_list, ignore_index=True)

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

        st.success("✅ 파일 업로드 완료! 아래에서 분석할 데이터의 범위를 선택해 주세요.")
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        available_chas = sorted(clean_df['차수'].unique())
        
        with col1:
            selected_cha = st.selectbox("📌 기준 차수 선택", available_chas, index=len(available_chas)-1, format_func=lambda x: f"{x}차")
        with col2:
            view_mode = st.radio("보기 옵션 (데이터 조회 범위)", ["1. 당해차수 데이터", "2. 1차~현재까지 데이터"])

        if view_mode == "1. 당해차수 데이터":
            filtered_clean_df = clean_df[clean_df['차수'] == selected_cha]
            st.info(f"선택됨: **{selected_cha}차** 당해차수 데이터만 분석합니다.")
        else:
            filtered_clean_df = clean_df[clean_df['차수'] <= selected_cha]
            st.info(f"선택됨: **1차 부터 {selected_cha}차 까지의 누적** 데이터를 분석합니다.")
            
        st.markdown("---")

        def get_analysis_result(row):
            npv, irr, irr_msg, _ = calculate_simulation(
                row['길이'], row['투자비'], row['분담금'], row['기타이익'], row['판매량'], row['판매액'], row['판매원가'], 
                row['총전수'], row['기본요금수익'], RATE, TAX, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m
            )
            return row['길이'], row['투자비'] - row['분담금'] - row['기타이익'], row['판매량'], npv, irr, irr_msg

        st.subheader("1. 📁 용도별 경제성 요약 (분석 대상 선택)")
        
        usage_results = []
        for u in filtered_clean_df['용도'].unique():
            u_df = filtered_clean_df[filtered_clean_df['용도'] == u]
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

        selected_usages = edited_df[edited_df['선택'] == True]['용도'].tolist()

        if selected_usages:
            final_filtered_df = filtered_clean_df[filtered_clean_df['용도'].isin(selected_usages)]
            t_len, t_net_inv, t_vol, tot_npv, tot_irr, tot_irr_msg = get_analysis_result(final_filtered_df[num_cols].sum())

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
            df_detail = final_filtered_df.groupby(['용도', '구간명'])[num_cols].sum().reset_index()
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
        st.info("👆 좌측 메뉴바 상단에서 분석을 시작할 전산 Raw 파일(*차)들을 업로드 해주세요.")


# ==========================================================================
# 탭 2: 신규 배관 투자 승인 내역 자동화
# ==========================================================================
elif menu_choice == '2. 배관 투자 승인 내역':
    
    # --- 탭 2 전용 좌측 사이드바 하단 UI 구성 ---
    with st.sidebar:
        st.header("💰 2026년 사업계획 투자한도액")
        st.markdown("아래 표에 항목별 **[규모]**와 **[금액]**을 입력하세요.")
        
        st.subheader("🔹 수요개발배관")
        df_sd_base = pd.DataFrame({
            "항목": ["공공택지", "공동주택", "산업용", "업무용", "영업용", "연료전지용", "주택용(지자체)", "투자보수율가산"],
            "규모": [1556, 906, 325, 498, 275, 735, 0, 3004], 
            "금액": [1055430560, 851196752, 287568274, 439429508, 182956113, 610435480, 0, 1695844012]  
        })
        edited_sd = st.data_editor(df_sd_base, key="sd_editor", hide_index=True, use_container_width=True)
        # 수요개발배관 소계 계산 및 천단위 콤마 표출
        sd_scale_sub = edited_sd['규모'].sum()
        sd_amt_sub = edited_sd['금액'].sum()
        st.markdown(f"<div style='text-align: right; color: #1E88E5;'><b>소계 ➔ 규모: {sd_scale_sub:,.0f} / 금액: {sd_amt_sub:,.0f}</b></div>", unsafe_allow_html=True)
        
        st.divider()

        st.subheader("🔹 기본계획배관")
        df_bp_base = pd.DataFrame({
            "항목": ["계획배관", "Loop", "이설배관", "지역정압기", "인입배관", "공급시설물 개선"],
            "규모": [2828, 749, 624, 3, 857, 95], 
            "금액": [2031952014, 626987840, 766452499, 338045023, 3230129038, 2749999724]  
        })
        edited_bp = st.data_editor(df_bp_base, key="bp_editor", hide_index=True, use_container_width=True)
        # 기본계획배관 소계 계산 및 천단위 콤마 표출
        bp_scale_sub = edited_bp['규모'].sum()
        bp_amt_sub = edited_bp['금액'].sum()
        st.markdown(f"<div style='text-align: right; color: #1E88E5;'><b>소계 ➔ 규모: {bp_scale_sub:,.0f} / 금액: {bp_amt_sub:,.0f}</b></div>", unsafe_allow_html=True)
        
        # 합산 예산 계산 (총 투자한도액)
        budget_2026 = int(sd_amt_sub + bp_amt_sub)

    # --- 메인 화면 ---
    st.title("📋 2026년도 배관 투자 승인 내역")
    st.markdown("기초자료 엑셀/CSV 파일을 바탕으로 2026년 투자 누계액과 승인 내역을 자동 산출합니다.")

    if uploaded_files:
        # ==== 신규 기능: 1번째 탭처럼 차수 필터링 UI (메인 화면 최상단) ====
        st.success("✅ 파일 업로드 완료! 아래에서 분석할 데이터의 범위를 선택해 주세요.")
        st.markdown("---")
        
        chas = []
        for f in uploaded_files:
            m = re.search(r'(\d+)차', f.name)
            chas.append(int(m.group(1)) if m else 1)
        available_chas_t2 = sorted(list(set(chas)))
        
        col1, col2 = st.columns(2)
        with col1:
            selected_cha_t2 = st.selectbox("📌 기준 차수 선택 (승인 내역용)", available_chas_t2, index=len(available_chas_t2)-1, format_func=lambda x: f"{x}차")
        with col2:
            view_mode_t2 = st.radio("보기 옵션 (데이터 조회 범위)", ["1. 당해차수 데이터", "2. 1차~현재까지 데이터"], key="view_mode_t2")

        if view_mode_t2 == "1. 당해차수 데이터":
            st.info(f"선택됨: **{selected_cha_t2}차** 당해차수 데이터만 상세 내역에 출력합니다.")
        else:
            st.info(f"선택됨: **1차 부터 {selected_cha_t2}차 까지의 누적** 데이터를 상세 내역에 출력합니다.")
            
        st.markdown("---")

        # ==== 우측 메인 화면 상단: 디폴트값 기반 2026년 사업계획 투자한도액 세부 구성 표 표출 ====
        st.subheader("📌 2026년 사업계획 투자한도액 세부 구성")
        
        # 1. 수요개발배관 원본 + 소계 병합
        df_sd_display = edited_sd.copy()
        df_sd_display.insert(0, '구분', '수요개발배관')
        df_sd_display.loc[len(df_sd_display)] = ['수요개발배관', '소계', sd_scale_sub, sd_amt_sub]
        
        # 2. 기본계획배관 원본 + 소계 병합
        df_bp_display = edited_bp.copy()
        df_bp_display.insert(0, '구분', '기본계획배관')
        df_bp_display.loc[len(df_bp_display)] = ['기본계획배관', '소계', bp_scale_sub, bp_amt_sub]
        
        # 3. 전체 병합 및 총계 산출 (엑셀 원본 구조 100% 반영)
        df_budget_detail = pd.concat([df_sd_display, df_bp_display], ignore_index=True)
        df_budget_detail.loc[len(df_budget_detail)] = ['합계', '총계', sd_scale_sub + bp_scale_sub, budget_2026]
        
        # 표 렌더링 (천 단위 콤마 포맷팅)
        st.dataframe(df_budget_detail.style.format({"규모": "{:,.0f}", "금액": "{:,.0f}"}), hide_index=True, use_container_width=True)
        st.divider()

        # 데이터 추출 로직
        all_data = []
        current_inv_amount = 0 # 선택된 차수(당해)의 합
        total_inv_amount = 0   # 누계 합
        
        for file in uploaded_files:
            try:
                match = re.search(r'(\d+)차', file.name)
                file_cha = int(match.group(1)) if match else 1
                
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, skiprows=2, encoding='utf-8-sig') 
                else:
                    df = pd.read_excel(file, skiprows=2)

                df.columns = df.columns.astype(str).str.replace(" ", "")

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
                    extracted['차수'] = file_cha
                    extracted['공사명'] = df[col_name]
                    extracted['투자비(원)'] = pd.to_numeric(df[col_inv], errors='coerce').fillna(0) if col_inv else 0
                    extracted['가정용 판매량(MJ)'] = pd.to_numeric(df[col_home], errors='coerce').fillna(0) if col_home else 0
                    extracted['일반용 판매량(MJ)'] = pd.to_numeric(df[col_general], errors='coerce').fillna(0) if col_general else 0
                    extracted['합계 판매량(MJ)'] = pd.to_numeric(df[col_total_vol], errors='coerce').fillna(0) if col_total_vol else 0
                    extracted['NPV(원)'] = pd.to_numeric(df[col_npv], errors='coerce').fillna(0) if col_npv else 0
                    extracted['IRR(%)'] = pd.to_numeric(df[col_irr], errors='coerce').fillna(0) if col_irr else 0

                    extracted = extracted.dropna(subset=['공사명'])
                    extracted = extracted[~extracted['공사명'].isin(['구간명', 'nan', ''])]
                    
                    # 기준 차수 누계액 계산
                    file_total_inv = extracted['투자비(원)'].sum()
                    if file_cha <= selected_cha_t2:
                        total_inv_amount += file_total_inv
                        if file_cha == selected_cha_t2:
                            current_inv_amount += file_total_inv

                    # 보기 옵션(필터링) 적용 후 데이터 저장
                    if view_mode_t2 == "1. 당해차수 데이터":
                        if file_cha == selected_cha_t2:
                            all_data.append(extracted)
                    else:
                        if file_cha <= selected_cha_t2:
                            all_data.append(extracted)

            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다 ({file.name}): {e}")
        
        if all_data:
            st.subheader("📊 2026년 배관 투자 승인 요약")
            summary_df = pd.DataFrame({
                "구분": ["2026년 사업계획 투자한도액", f"금회 신청 내역 ({selected_cha_t2}차)", "2026년 본부투자 누계", "잔여 한도액"],
                "금액(원)": [
                    budget_2026, 
                    current_inv_amount, 
                    total_inv_amount, 
                    budget_2026 - total_inv_amount
                ]
            })
            
            st.dataframe(summary_df.style.format({"금액(원)": "{:,.0f}"}), hide_index=True, use_container_width=True)
            
            st.divider()
            
            detail_title_prefix = f"{selected_cha_t2}차 당해" if view_mode_t2 == "1. 당해차수 데이터" else f"{selected_cha_t2}차 누계"
            st.subheader(f"📝 {detail_title_prefix} 상세 승인 내역")
            
            final_df = pd.concat(all_data, ignore_index=True)
            final_df.index = final_df.index + 1 
            
            # 화면 출력에서 '차수' 컬럼은 숨기고(필요시) 표출
            display_cols = ['공사명', '투자비(원)', '가정용 판매량(MJ)', '일반용 판매량(MJ)', '합계 판매량(MJ)', 'NPV(원)', 'IRR(%)']
            st.dataframe(final_df[display_cols].style.format({
                "투자비(원)": "{:,.0f}",
                "가정용 판매량(MJ)": "{:,.0f}",
                "일반용 판매량(MJ)": "{:,.0f}",
                "합계 판매량(MJ)": "{:,.0f}",
                "NPV(원)": "{:,.0f}",
                "IRR(%)": "{:,.2f}"
            }), use_container_width=True)
            
            csv_data = final_df[display_cols].to_csv(index=True).encode('utf-8-sig')
            st.download_button(
                label="📥 상세 승인내역 CSV 다운로드",
                data=csv_data,
                file_name=f"2026년도_배관투자_승인내역_{detail_title_prefix}.csv",
                mime="text/csv"
            )
        else:
            st.info("조건에 맞는 데이터가 없습니다.")
    else:
        st.info("👆 좌측 메뉴바 상단에서 분석을 시작할 전산 Raw 파일(*차)들을 업로드 해주세요.")
