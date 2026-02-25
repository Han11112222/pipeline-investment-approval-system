import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import re
import plotly.express as px

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
    st.markdown("수요개발(기초자료) 및 기본계획/인입(투자현황정리) 엑셀을 모두 드래그 앤 드롭 하세요.")
    uploaded_files = st.file_uploader("기초자료 및 현황정리 파일 업로드 (*다중 선택 가능)", accept_multiple_files=True, type=['csv', 'xlsx', 'xls'])
    
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
                # [안전장치] 투자현황 파일은 1탭에서 무시하도록 건너뜀
                if '현황정리' in file.name or '투자현황' in file.name:
                    continue
                    
                file.seek(0)
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

        st.success("✅ 기초자료 파일 업로드 완료! 아래에서 분석할 데이터의 범위를 선택해 주세요.")
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
    
    # [핵심 추가] 공무팀 현황정리 파일 파싱 로직
    extracted_bp_data = {}
    if uploaded_files:
        for file in uploaded_files:
            if '현황정리' in file.name or '투자현황' in file.name:
                try:
                    file.seek(0)
                    if file.name.endswith('.csv'):
                        df_status = pd.read_csv(file, header=None)
                    else:
                        # 엑셀 파일인 경우 '투자현황' 시트가 존재하면 그걸 읽음
                        try:
                            df_status = pd.read_excel(file, sheet_name='투자현황', header=None)
                        except:
                            df_status = pd.read_excel(file, header=None)

                    # 엑셀 파일 내에서 '기본계획'과 '인입배관' 등의 값을 스캔해서 딕셔너리에 저장
                    for row_idx in range(df_status.shape[0]):
                        col1_val = str(df_status.iloc[row_idx, 1]).strip()
                        
                        target_items = ['계획배관', 'Loop', '이설배관', '지역정압기', '지역\n정압기', '공급시설물 개선', '공급시설물\n개선', '인입배관']
                        
                        if any(item in col1_val for item in target_items):
                            # 보통 12, 13열에 '기승인(2026년 본부투자 승인내역)' 규모/금액, 15, 16열에 '금회(3차)' 규모/금액이 있음
                            # (엑셀 양식에 맞춰 인덱스 강제 지정: 12=기승인규모, 13=기승인금액, 15=금회규모, 16=금회금액)
                            try:
                                p_scale = pd.to_numeric(str(df_status.iloc[row_idx, 11]).replace(',', ''), errors='coerce')
                                p_amt = pd.to_numeric(str(df_status.iloc[row_idx, 12]).replace(',', ''), errors='coerce')
                                c_scale = pd.to_numeric(str(df_status.iloc[row_idx, 14]).replace(',', ''), errors='coerce')
                                c_amt = pd.to_numeric(str(df_status.iloc[row_idx, 15]).replace(',', ''), errors='coerce')
                                
                                # 공백 제거나 줄바꿈 정규화
                                clean_key = col1_val.replace('\n', '')
                                extracted_bp_data[clean_key] = {
                                    '기승인_규모': p_scale if not pd.isna(p_scale) else 0,
                                    '기승인_금액': p_amt if not pd.isna(p_amt) else 0,
                                    '금회_규모': c_scale if not pd.isna(c_scale) else 0,
                                    '금회_금액': c_amt if not pd.isna(c_amt) else 0
                                }
                            except:
                                pass
                except Exception as e:
                    pass

    # --- 탭 2 전용 좌측 사이드바 하단 UI 구성 ---
    with st.sidebar:
        st.header("💰 2026년 사업계획 투자한도액 세팅")
        
        st.subheader("🔹 1. 수요개발배관 (실적은 기초자료 추출)")
        df_sd_base = pd.DataFrame({
            "항목": ["공공택지", "공동주택", "산업용", "업무용", "영업용", "연료전지용", "주택용(지자체)", "투자보수율가산"],
            "한도_규모": [1556, 906, 325, 498, 275, 735, 0, 3004], 
            "한도_금액": [1055430560, 851196752, 287568274, 439429508, 182956113, 610435480, 0, 1695844012]  
        })
        edited_sd = st.data_editor(df_sd_base, key="sd_editor", hide_index=True, use_container_width=True)
        
        st.divider()

        st.subheader("🔹 2. 기본계획배관 (현황정리 엑셀 연동 + 수기수정)")
        st.markdown("현황정리 파일을 업로드하면 값이 자동 채워집니다. 아래 표에서 직접 수정도 가능합니다.")
        
        # 기본계획배관 베이스 데이터 프레임 생성
        bp_items = ["계획배관", "Loop", "이설배관", "지역정압기", "공급시설물 개선"]
        bp_limit_s = [2828, 749, 624, 3, 95]
        bp_limit_a = [2031952014, 626987840, 766452499, 338045023, 2749999724]
        
        bp_p_s, bp_p_a, bp_c_s, bp_c_a = [], [], [], []
        
        for item in bp_items:
            clean_item = item.replace('\n', '')
            if clean_item in extracted_bp_data:
                bp_p_s.append(extracted_bp_data[clean_item]['기승인_규모'])
                bp_p_a.append(extracted_bp_data[clean_item]['기승인_금액'])
                bp_c_s.append(extracted_bp_data[clean_item]['금회_규모'])
                bp_c_a.append(extracted_bp_data[clean_item]['금회_금액'])
            else:
                bp_p_s.append(0); bp_p_a.append(0); bp_c_s.append(0); bp_c_a.append(0)

        df_bp_base = pd.DataFrame({
            "항목": bp_items,
            "한도_규모": bp_limit_s, 
            "한도_금액": bp_limit_a,
            "기승인_규모": bp_p_s,
            "기승인_금액": bp_p_a,
            "금회_규모": bp_c_s,
            "금회_금액": bp_c_a
        })
        edited_bp = st.data_editor(df_bp_base, key="bp_editor", hide_index=True, use_container_width=True)

        st.divider()

        st.subheader("🔹 3. 65A미만 인입 (현황정리 엑셀 연동 + 수기수정)")
        
        in_p_s, in_p_a, in_c_s, in_c_a = 0, 0, 0, 0
        if '인입배관' in extracted_bp_data:
            in_p_s = extracted_bp_data['인입배관']['기승인_규모']
            in_p_a = extracted_bp_data['인입배관']['기승인_금액']
            in_c_s = extracted_bp_data['인입배관']['금회_규모']
            in_c_a = extracted_bp_data['인입배관']['금회_금액']

        df_in_base = pd.DataFrame({
            "항목": ["65A미만 인입"],
            "한도_규모": [857], 
            "한도_금액": [3230129038],
            "기승인_규모": [in_p_s],
            "기승인_금액": [in_p_a],
            "금회_규모": [in_c_s],
            "금회_금액": [in_c_a]
        })
        edited_in = st.data_editor(df_in_base, key="in_editor", hide_index=True, use_container_width=True)

    # --- 메인 화면 ---
    st.title("📋 2026년도 배관 투자 승인 내역")
    st.markdown("기초자료 파일을 바탕으로 **수요개발배관은 자동 계산**하고, **기본계획배관 및 65A미만 인입은 공무팀 실적 파일과 연동**하여 산출합니다.")

    if uploaded_files:
        st.success("✅ 파일 업로드 완료! 아래에서 분석할 데이터의 범위를 선택해 주세요.")
        st.markdown("---")
        
        chas = []
        for f in uploaded_files:
            m = re.search(r'(\d+)차', f.name)
            if m: chas.append(int(m.group(1)))
        
        if not chas:
            available_chas_t2 = [1]
        else:
            available_chas_t2 = sorted(list(set(chas)))
        
        col1, col2 = st.columns(2)
        with col1:
            selected_cha_t2 = st.selectbox("📌 기준 차수 선택 (승인 내역용)", available_chas_t2, index=len(available_chas_t2)-1, format_func=lambda x: f"{x}차")
        with col2:
            view_mode_t2 = st.radio("보기 옵션 (데이터 조회 범위)", ["1. 당해차수 데이터", "2. 1차~현재까지 데이터"], key="view_mode_t2")
            
        st.markdown("---")

        # ==== 1. 데이터 추출 ====
        all_data_unfiltered = []
        
        def get_multi_col_idx(df_temp, keywords):
            found_cols = []
            for col_idx in range(df_temp.shape[1]):
                for row_idx in range(min(20, df_temp.shape[0])):
                    val = str(df_temp.iloc[row_idx, col_idx]).replace(" ", "").replace("\n", "")
                    for kw in keywords:
                        if kw in val:
                            found_cols.append(col_idx)
            return list(set(found_cols))
        
        for file in uploaded_files:
            # 1탭에서 사용하는 기초자료 파일만 분석
            if '현황정리' in file.name or '투자현황' in file.name:
                continue

            try:
                file.seek(0)
                match = re.search(r'(\d+)차', file.name)
                file_cha = int(match.group(1)) if match else 1
                
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, header=None, encoding='utf-8-sig') 
                else:
                    df = pd.read_excel(file, header=None)

                extracted = pd.DataFrame()
                extracted['항목'] = df.iloc[:, 0].astype(str).str.strip().replace(['nan', 'None', ''], np.nan).ffill().fillna('미분류')
                
                if df.shape[1] > 6:
                    extracted['규모(m)'] = df.iloc[:, 5]
                    extracted['투자비(원)'] = df.iloc[:, 6]
                else:
                    extracted['규모(m)'] = 0
                    extracted['투자비(원)'] = 0

                idx_name = get_col_idx(df, ["구간명"], exact=True)
                extracted['공사명'] = df.iloc[:, idx_name] if idx_name is not None else df.iloc[:, 1]
                extracted['차수'] = file_cha

                home_cols = get_multi_col_idx(df, ["취사용(MJ)", "개별난방용(MJ)", "중앙난방용(MJ)"])
                gen_cols = get_multi_col_idx(df, ["일반용(영업1)(MJ)", "일반용(영업2)(MJ)"])
                idx_total_vol = get_col_idx(df, ["계(MJ)"], exact=False)
                idx_npv = get_col_idx(df, ["NPV"], exact=False)
                idx_irr = get_col_idx(df, ["IRR"], exact=False)

                def get_clean_series(c_idx):
                    return pd.to_numeric(df.iloc[:, c_idx].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)

                extracted['가정용 판매량(MJ)'] = 0
                for c_idx in home_cols:
                    extracted['가정용 판매량(MJ)'] += get_clean_series(c_idx)

                extracted['일반용 판매량(MJ)'] = 0
                for c_idx in gen_cols:
                    extracted['일반용 판매량(MJ)'] += get_clean_series(c_idx)

                extracted['합계 판매량(MJ)'] = get_clean_series(idx_total_vol) if idx_total_vol is not None else 0
                extracted['NPV(원)'] = get_clean_series(idx_npv) if idx_npv is not None else 0
                extracted['IRR(%)'] = get_clean_series(idx_irr) if idx_irr is not None else 0

                extracted['공사명'] = extracted['공사명'].astype(str).str.strip()
                invalid_names = ['', '0', 'nan', 'None', '구간명', '소계', '합계', '총계']
                extracted = extracted[~extracted['공사명'].isin(invalid_names)]
                
                def clean_numeric(x):
                    s = str(x).replace(',', '')
                    s = re.sub(r'[^\d.-]', '', s)
                    try:
                        return float(s) if s else 0.0
                    except:
                        return 0.0

                num_cols_ext = ['규모(m)', '투자비(원)', '가정용 판매량(MJ)', '일반용 판매량(MJ)', '합계 판매량(MJ)', 'NPV(원)', 'IRR(%)']
                for c in num_cols_ext:
                    extracted[c] = extracted[c].apply(clean_numeric)
                    
                all_data_unfiltered.append(extracted)

            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다 ({file.name}): {e}")

        # ==== 2. 수요개발배관 파일 자동 매핑을 위한 그룹핑 ====
        if all_data_unfiltered:
            all_parsed_df = pd.concat(all_data_unfiltered, ignore_index=True)
            all_parsed_df['항목_clean'] = all_parsed_df['항목'].astype(str).str.replace(r'\s+', '', regex=True)
            
            prev_df = all_parsed_df[all_parsed_df['차수'] < selected_cha_t2]
            curr_df = all_parsed_df[all_parsed_df['차수'] == selected_cha_t2]
            
            prev_agg = prev_df.groupby('항목_clean')[['규모(m)', '투자비(원)']].sum()
            curr_agg = curr_df.groupby('항목_clean')[['규모(m)', '투자비(원)']].sum()
        else:
            all_parsed_df = pd.DataFrame()
            prev_agg = pd.DataFrame()
            curr_agg = pd.DataFrame()

        def fill_sd_metrics(df_base):
            for col in ['기승인_규모', '기승인_금액', '금회_규모', '금회_금액', '누계_규모', '누계_금액']:
                df_base[col] = 0.0
                
            for i in range(len(df_base)):
                item_raw = df_base.at[i, '항목']
                if item_raw == '소계': continue
                item_clean = str(item_raw).replace('\n', '').replace(' ', '')
                
                if not prev_agg.empty:
                    matched_idx = [idx for idx in prev_agg.index if item_clean in str(idx) or str(idx) in item_clean]
                    if matched_idx:
                        df_base.at[i, '기승인_규모'] = prev_agg.loc[matched_idx, '규모(m)'].sum()
                        df_base.at[i, '기승인_금액'] = prev_agg.loc[matched_idx, '투자비(원)'].sum()
                
                if not curr_agg.empty:
                    matched_idx = [idx for idx in curr_agg.index if item_clean in str(idx) or str(idx) in item_clean]
                    if matched_idx:
                        df_base.at[i, '금회_규모'] = curr_agg.loc[matched_idx, '규모(m)'].sum()
                        df_base.at[i, '금회_금액'] = curr_agg.loc[matched_idx, '투자비(원)'].sum()
                
                df_base.at[i, '누계_규모'] = df_base.at[i, '기승인_규모'] + df_base.at[i, '금회_규모']
                df_base.at[i, '누계_금액'] = df_base.at[i, '기승인_금액'] + df_base.at[i, '금회_금액']
                df_base.at[i, '잔여_금액'] = df_base.at[i, '한도_금액'] - df_base.at[i, '누계_금액']
                
            sub_idx = len(df_base) - 1
            for col in ['기승인_규모', '기승인_금액', '금회_규모', '금회_금액', '누계_규모', '누계_금액', '잔여_금액']:
                df_base.at[sub_idx, col] = df_base.iloc[:-1][col].sum()
                
            return df_base


        # ==== 3. 프레임 구성 및 총계 산출 ====
        # 1) 수요개발배관
        df_sd_display = edited_sd.copy()
        df_sd_display.rename(columns={'규모': '한도_규모', '금액': '한도_금액'}, inplace=True)
        df_sd_display.insert(0, '구분', '수요개발배관')
        df_sd_display.loc[len(df_sd_display)] = ['수요개발배관', '소계', df_sd_display['한도_규모'].sum(), df_sd_display['한도_금액'].sum()] 
        df_sd_display = fill_sd_metrics(df_sd_display)
        
        # 2) 기본계획배관
        df_bp_display = edited_bp.copy()
        df_bp_display.insert(0, '구분', '기본계획배관')
        df_bp_display['누계_규모'] = df_bp_display['기승인_규모'] + df_bp_display['금회_규모']
        df_bp_display['누계_금액'] = df_bp_display['기승인_금액'] + df_bp_display['금회_금액']
        df_bp_display['잔여_금액'] = df_bp_display['한도_금액'] - df_bp_display['누계_금액']
        bp_sub_row = {
            '구분': '기본계획배관', '항목': '소계',
            '한도_규모': df_bp_display['한도_규모'].sum(), '한도_금액': df_bp_display['한도_금액'].sum(),
            '기승인_규모': df_bp_display['기승인_규모'].sum(), '기승인_금액': df_bp_display['기승인_금액'].sum(),
            '금회_규모': df_bp_display['금회_규모'].sum(), '금회_금액': df_bp_display['금회_금액'].sum(),
            '누계_규모': df_bp_display['누계_규모'].sum(), '누계_금액': df_bp_display['누계_금액'].sum(),
            '잔여_금액': df_bp_display['잔여_금액'].sum()
        }
        df_bp_display = pd.concat([df_bp_display, pd.DataFrame([bp_sub_row])], ignore_index=True)

        # 3) 65A미만 인입
        df_in_display = edited_in.copy()
        df_in_display.insert(0, '구분', '65A미만 인입')
        df_in_display['누계_규모'] = df_in_display['기승인_규모'] + df_in_display['금회_규모']
        df_in_display['누계_금액'] = df_in_display['기승인_금액'] + df_in_display['금회_금액']
        df_in_display['잔여_금액'] = df_in_display['한도_금액'] - df_in_display['누계_금액']
        in_sub_row = {
            '구분': '65A미만 인입', '항목': '소계',
            '한도_규모': df_in_display['한도_규모'].sum(), '한도_금액': df_in_display['한도_금액'].sum(),
            '기승인_규모': df_in_display['기승인_규모'].sum(), '기승인_금액': df_in_display['기승인_금액'].sum(),
            '금회_규모': df_in_display['금회_규모'].sum(), '금회_금액': df_in_display['금회_금액'].sum(),
            '누계_규모': df_in_display['누계_규모'].sum(), '누계_금액': df_in_display['누계_금액'].sum(),
            '잔여_금액': df_in_display['잔여_금액'].sum()
        }
        df_in_display = pd.concat([df_in_display, pd.DataFrame([in_sub_row])], ignore_index=True)
        
        # 합체 및 총계 산출
        df_budget_detail = pd.concat([df_sd_display, df_bp_display, df_in_display], ignore_index=True)
        
        sd_sub_idx = len(df_sd_display) - 1
        bp_sub_idx = len(df_bp_display) - 1
        in_sub_idx = len(df_in_display) - 1
        
        tot_limit_scale = df_sd_display.at[sd_sub_idx, '한도_규모'] + df_bp_display.at[bp_sub_idx, '한도_규모'] + df_in_display.at[in_sub_idx, '한도_규모']
        tot_limit_amt = df_sd_display.at[sd_sub_idx, '한도_금액'] + df_bp_display.at[bp_sub_idx, '한도_금액'] + df_in_display.at[in_sub_idx, '한도_금액']
        tot_prev_scale = df_sd_display.at[sd_sub_idx, '기승인_규모'] + df_bp_display.at[bp_sub_idx, '기승인_규모'] + df_in_display.at[in_sub_idx, '기승인_규모']
        tot_prev_amt = df_sd_display.at[sd_sub_idx, '기승인_금액'] + df_bp_display.at[bp_sub_idx, '기승인_금액'] + df_in_display.at[in_sub_idx, '기승인_금액']
        tot_curr_scale = df_sd_display.at[sd_sub_idx, '금회_규모'] + df_bp_display.at[bp_sub_idx, '금회_규모'] + df_in_display.at[in_sub_idx, '금회_규모']
        tot_curr_amt = df_sd_display.at[sd_sub_idx, '금회_금액'] + df_bp_display.at[bp_sub_idx, '금회_금액'] + df_in_display.at[in_sub_idx, '금회_금액']
        tot_cum_scale = df_sd_display.at[sd_sub_idx, '누계_규모'] + df_bp_display.at[bp_sub_idx, '누계_규모'] + df_in_display.at[in_sub_idx, '누계_규모']
        tot_cum_amt = df_sd_display.at[sd_sub_idx, '누계_금액'] + df_bp_display.at[bp_sub_idx, '누계_금액'] + df_in_display.at[in_sub_idx, '누계_금액']
        tot_remain_amt = df_sd_display.at[sd_sub_idx, '잔여_금액'] + df_bp_display.at[bp_sub_idx, '잔여_금액'] + df_in_display.at[in_sub_idx, '잔여_금액']

        df_budget_detail.loc[len(df_budget_detail)] = [
            '합계', '총계', 
            tot_limit_scale, tot_limit_amt,          
            tot_prev_scale, tot_prev_amt,                      
            tot_curr_scale, tot_curr_amt,             
            tot_cum_scale, tot_cum_amt,                 
            tot_remain_amt                     
        ]


        # =====================================================================
        # [신규 추가] 상단 전체 요약 표 & 그래프
        # =====================================================================
        st.subheader("📈 2026년 배관 투자 전체 요약 및 진척도")
        
        sum_df = pd.DataFrame({
            "구분": ["수요개발배관", "기본계획배관", "65A미만 인입", "총계"],
            "총 투자한도액": [df_sd_display.at[sd_sub_idx, '한도_금액'], df_bp_display.at[bp_sub_idx, '한도_금액'], df_in_display.at[in_sub_idx, '한도_금액'], tot_limit_amt],
            "기승인 누계": [df_sd_display.at[sd_sub_idx, '기승인_금액'], df_bp_display.at[bp_sub_idx, '기승인_금액'], df_in_display.at[in_sub_idx, '기승인_금액'], tot_prev_amt],
            f"금회 신청 ({selected_cha_t2}차)": [df_sd_display.at[sd_sub_idx, '금회_금액'], df_bp_display.at[bp_sub_idx, '금회_금액'], df_in_display.at[in_sub_idx, '금회_금액'], tot_curr_amt],
            "현재 본부 누계": [df_sd_display.at[sd_sub_idx, '누계_금액'], df_bp_display.at[bp_sub_idx, '누계_금액'], df_in_display.at[in_sub_idx, '누계_금액'], tot_cum_amt],
            "잔여 한도액": [df_sd_display.at[sd_sub_idx, '잔여_금액'], df_bp_display.at[bp_sub_idx, '잔여_금액'], df_in_display.at[in_sub_idx, '잔여_금액'], tot_remain_amt]
        })
        sum_df['집행률(%)'] = np.where(sum_df['총 투자한도액'] > 0, (sum_df['현재 본부 누계'] / sum_df['총 투자한도액']) * 100, 0)
        
        st.dataframe(sum_df.style.format({
            "총 투자한도액": "{:,.0f}", "기승인 누계": "{:,.0f}", f"금회 신청 ({selected_cha_t2}차)": "{:,.0f}",
            "현재 본부 누계": "{:,.0f}", "잔여 한도액": "{:,.0f}", "집행률(%)": "{:,.1f}%"
        }).apply(lambda x: ['background-color: #FFE6E6; font-weight: bold'] * len(x) if x['구분'] == '총계' else [''] * len(x), axis=1), 
        hide_index=True, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        col_c1, col_c2 = st.columns([2, 1.5])
        
        with col_c1:
            st.markdown("#### 📊 구분별 투자 예산 및 실적 현황")
            chart_df = sum_df.iloc[:3].copy()
            bar_data = pd.DataFrame({
                "구분": chart_df["구분"].tolist() * 2,
                "금액(원)": chart_df["총 투자한도액"].tolist() + chart_df["현재 본부 누계"].tolist(),
                "유형": ["총 투자한도액"] * 3 + ["현재 본부 누계"] * 3
            })
            fig_bar = px.bar(bar_data, x="구분", y="금액(원)", color="유형", barmode="group",
                             color_discrete_map={"총 투자한도액": "#1E88E5", "현재 본부 누계": "#FFB300"})
            fig_bar.update_layout(margin=dict(t=20, b=0, l=0, r=0), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_c2:
            st.markdown("#### 🍩 전체 투자계획 대비 실적")
            donut_rem = tot_remain_amt if tot_remain_amt > 0 else 0
            donut_df = pd.DataFrame({
                "상태": ["실적 (누계)", "잔여 한도액"],
                "금액": [tot_cum_amt, donut_rem]
            })
            fig_donut = px.pie(donut_df, values='금액', names='상태', hole=0.5, 
                               color='상태', color_discrete_map={"실적 (누계)": "#FFB300", "잔여 한도액": "#E0E0E0"})
            fig_donut.update_traces(textinfo='percent+label', textposition='inside')
            
            exec_rate = (tot_cum_amt / tot_limit_amt * 100) if tot_limit_amt > 0 else 0
            fig_donut.update_layout(showlegend=False, margin=dict(t=20, b=0, l=0, r=0), 
                                    annotations=[dict(text=f"집행률<br><b>{exec_rate:.1f}%</b>", x=0.5, y=0.5, font_size=18, showarrow=False)])
            st.plotly_chart(fig_donut, use_container_width=True)

        st.divider()


        # ==== 4. 엑셀 스타일 상세 표 렌더링 ====
        st.subheader("📌 2026년 배관 투자 승인 요약 (Excel 양식)")

        df_budget_detail['승인비율'] = np.where(df_budget_detail['한도_금액'] > 0, (df_budget_detail['누계_금액'] / df_budget_detail['한도_금액']) * 100, 0)
        df_budget_detail['구분'] = df_budget_detail['구분'].mask(df_budget_detail['구분'].duplicated(), '')

        df_budget_detail.columns = pd.MultiIndex.from_tuples([
            ('구분', ''), ('항목', ''),
            ('2026년 사업계획 투자한도액', '규모(m)'), ('2026년 사업계획 투자한도액', '금액'),
            ('2026년 본부투자 승인내역 (기승인)', '규모(m)'), ('2026년 본부투자 승인내역 (기승인)', '금액'),
            (f'금회 신청 내역 ({selected_cha_t2}차)', '규모(m)'), (f'금회 신청 내역 ({selected_cha_t2}차)', '금액'),
            ('2026년 본부투자 누계', '규모(m)'), ('2026년 본부투자 누계', '금액'),
            ('투자한도 잔액', '금액'), ('승인비율', '(%)')
        ])

        def style_rows(row):
            if row[('항목', '')] == '소계':
                return ['background-color: #E6F3FF; font-weight: bold'] * len(row)
            elif row[('항목', '')] == '총계':
                return ['background-color: #FFE6E6; font-weight: bold; color: #D32F2F'] * len(row) 
            return [''] * len(row)

        format_dict = {col: "{:,.0f}" for col in df_budget_detail.columns if col[0] not in ['구분', '항목', '승인비율']}
        format_dict[('승인비율', '(%)')] = "{:,.1f}%"

        styled_df = df_budget_detail.style.format(format_dict).apply(style_rows, axis=1)
        st.dataframe(styled_df, hide_index=True, use_container_width=True)
        st.divider()
        
        
        # ==== 하단: 필터링 조건에 맞는 상세 데이터 표출 ====
        if not all_parsed_df.empty:
            if view_mode_t2 == "1. 당해차수 데이터":
                detail_df = all_parsed_df[all_parsed_df['차수'] == selected_cha_t2]
            else:
                detail_df = all_parsed_df[all_parsed_df['차수'] <= selected_cha_t2]
                
            if not detail_df.empty:
                detail_title_prefix = f"{selected_cha_t2}차 당해" if view_mode_t2 == "1. 당해차수 데이터" else f"{selected_cha_t2}차 누계"
                
                # =====================================================================
                # [신규 추가] 용도별 요약 표 (건수 추가)
                # =====================================================================
                st.subheader(f"📊 {detail_title_prefix} 용도별 실적 요약")
                
                usage_counts = detail_df.groupby('항목').size().reset_index(name='건수')
                usage_sums = detail_df.groupby('항목')[['규모(m)', '투자비(원)', '가정용 판매량(MJ)', '일반용 판매량(MJ)', '합계 판매량(MJ)']].sum().reset_index()
                
                usage_summary = pd.merge(usage_counts, usage_sums, on='항목')
                
                usage_summary.loc[len(usage_summary)] = [
                    '총계', 
                    usage_summary['건수'].sum(),
                    usage_summary['규모(m)'].sum(), 
                    usage_summary['투자비(원)'].sum(), 
                    usage_summary['가정용 판매량(MJ)'].sum(), 
                    usage_summary['일반용 판매량(MJ)'].sum(), 
                    usage_summary['합계 판매량(MJ)'].sum()
                ]
                
                st.dataframe(usage_summary.style.format({
                    "건수": "{:,.0f} 건",
                    "규모(m)": "{:,.0f}",
                    "투자비(원)": "{:,.0f}",
                    "가정용 판매량(MJ)": "{:,.0f}",
                    "일반용 판매량(MJ)": "{:,.0f}",
                    "합계 판매량(MJ)": "{:,.0f}"
                }).apply(lambda x: ['background-color: #F5F5F5; font-weight: bold'] * len(x) if x['항목'] == '총계' else [''] * len(x), axis=1), 
                hide_index=True, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # =====================================================================
                # 기존 상세 승인 내역 리스트
                # =====================================================================
                st.subheader(f"📝 {detail_title_prefix} 상세 승인 내역 리스트")
                
                detail_df = detail_df.reset_index(drop=True)
                display_cols = ['항목', '공사명', '규모(m)', '투자비(원)', '가정용 판매량(MJ)', '일반용 판매량(MJ)', '합계 판매량(MJ)', 'NPV(원)', 'IRR(%)']
                
                st.dataframe(detail_df[display_cols].style.format({
                    "규모(m)": "{:,.0f}",
                    "투자비(원)": "{:,.0f}",
                    "가정용 판매량(MJ)": "{:,.0f}",
                    "일반용 판매량(MJ)": "{:,.0f}",
                    "합계 판매량(MJ)": "{:,.0f}",
                    "NPV(원)": "{:,.0f}",
                    "IRR(%)": "{:,.2f}"
                }), use_container_width=True, hide_index=True)
                
                csv_data = detail_df[display_cols].to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 상세 승인내역 CSV 다운로드",
                    data=csv_data,
                    file_name=f"2026년도_배관투자_승인내역_{detail_title_prefix}.csv",
                    mime="text/csv"
                )
            else:
                st.info("해당 차수에 추출된 세부 공사 내역이 없습니다.")
        else:
            st.info("조건에 맞는 데이터가 없습니다.")
    else:
        st.info("👆 좌측 메뉴바 상단에서 분석을 시작할 전산 Raw 파일(*차)들을 업로드 해주세요.")
