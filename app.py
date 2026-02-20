import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

# --------------------------------------------------------------------------
# [설정] 페이지 기본
# --------------------------------------------------------------------------
st.set_page_config(page_title="공식 배관 투자 결재 시스템 (Pipeline Approval)", page_icon="🏗️", layout="wide")

# --------------------------------------------------------------------------
# [함수] 금융 계산 로직 (★ 엑셀 방식 100% 동기화 ★)
# --------------------------------------------------------------------------
def manual_npv(rate, values):
    # 엑셀의 NPV 함수 타임라인과 동일하게 맞춤 (Year 0을 어떻게 처리할지 조정)
    return sum(v / ((1 + rate) ** i) for i, v in enumerate(values))

def calculate_simulation(sim_len, sim_inv, sim_contrib, sim_other, sim_vol, sim_rev, sim_cost, 
                         sim_jeon, sim_basic_rev, rate, tax, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m):
    
    # 1. 초기 순투자액
    net_inv = round(sim_inv - sim_contrib - sim_other)
    
    # 2. 고정 수익/비용 항목 계산 (엑셀처럼 셀 단위로 반올림 적용)
    margin_total = round((sim_rev - sim_cost) + sim_basic_rev)
    
    cost_sga = round((sim_len * c_maint) + (sim_len * c_adm_m) + (sim_jeon * c_adm_jeon))
    annual_depreciation = round(sim_inv / dep_period) if dep_period > 0 else 0
    
    # 3. 엑셀과 동일한 세후 현금흐름(OCF) 산출
    ebit = margin_total - cost_sga - annual_depreciation
    net_income = round(ebit * (1 - tax)) # 세금 계산 후 엑셀처럼 정수 반올림
    fixed_ocf = net_income + annual_depreciation
    
    flows = [-net_inv]
    ocfs = []
    
    for year in range(1, int(analysis_period) + 1):
        flows.append(fixed_ocf)
        ocfs.append(fixed_ocf)

    # 4. 지표 산출
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

# --------------------------------------------------------------------------
# [함수] 2차원 좌표 스캔
# --------------------------------------------------------------------------
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
# [UI] 좌측 사이드바
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 분석 변수 설정")
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

# --------------------------------------------------------------------------
# [UI] 메인 화면
# --------------------------------------------------------------------------
st.title("🏗️ 배관 투자 경제성 결재 대시보드")
st.markdown("결재용 엑셀을 업로드하면 데이터를 자동 추출합니다. **분석에서 제외할 항목은 체크 해제**하세요.")

uploaded_file = st.file_uploader("📂 결재용 Raw 데이터 파일 업로드 (Excel 또는 CSV)", type=['xlsx', 'xls', 'csv'])

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

    clean_df['용도'] = clean_df['용도'].astype(str).str.strip()
    clean_df['용도'] = clean_df['용도'].replace(['', '0', 'nan', 'None', '용도'], np.nan)
    clean_df['용도'] = clean_df['용도'].ffill().fillna('미분류')

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

    st.success("✅ 파일 업로드 완료! 데이터 스캔 및 결재용 경제성 엔진 준비 완료.")

    def get_analysis_result(row):
        s_len = row['길이']
        s_inv = row['투자비']
        s_contrib = row['분담금']
        s_other = row['기타이익']
        s_jeon = row['총전수']
        s_vol = row['판매량']
        s_rev = row['판매액']
        s_cost = row['판매원가']
        s_basic_rev = row['기본요금수익'] 

        npv, irr, irr_msg, _ = calculate_simulation(
            s_len, s_inv, s_contrib, s_other, s_vol, s_rev, s_cost, 
            s_jeon, s_basic_rev, RATE, TAX, dep_period, analysis_period, 
            c_maint, c_adm_jeon, c_adm_m
        )
        return s_len, s_inv - s_contrib - s_other, s_vol, npv, irr, irr_msg

    # ---------------------------------------------------------
    # UI Section 1: 용도별 요약 & 체크박스
    # ---------------------------------------------------------
    st.subheader("1. 📁 용도별 경제성 요약 (분석 대상 선택)")
    
    usage_results = []
    unique_usages = clean_df['용도'].unique()
    
    for u in unique_usages:
        u_df = clean_df[clean_df['용도'] == u]
        u_len, u_net_inv, u_vol, u_npv, u_irr, u_irr_msg = get_analysis_result(u_df[num_cols].sum())
        
        is_selected = False if 'ROE' in str(u).upper() else True
        
        usage_results.append({
            "선택": is_selected,
            "용도": u,
            "총 투자길이(m)": f"{u_len:,.1f}",
            "총 순투자액(원)": f"{u_net_inv:,.0f}",
            "연간판매량(MJ)": f"{u_vol:,.0f}",
            "NPV(원)": f"{u_npv:,.0f}",
            "IRR(%)": f"{u_irr*100:.2f}%" if u_irr is not None else u_irr_msg
        })
        
    df_usage_summary = pd.DataFrame(usage_results)

    edited_df = st.data_editor(
        df_usage_summary,
        column_config={
            "선택": st.column_config.CheckboxColumn("선택", help="분석에 포함하려면 체크하세요"),
        },
        disabled=["용도", "총 투자길이(m)", "총 순투자액(원)", "연간판매량(MJ)", "NPV(원)", "IRR(%)"],
        hide_index=True,
        use_container_width=True
    )

    # ---------------------------------------------------------
    # UI Section 2 & 3: 필터링 및 소계/상세 계산
    # ---------------------------------------------------------
    selected_usages = edited_df[edited_df['선택'] == True]['용도'].tolist()

    if not selected_usages:
        st.warning("⚠️ 선택된 용도가 없습니다. 표에서 하나 이상의 용도를 체크해 주세요.")
    else:
        filtered_df = clean_df[clean_df['용도'].isin(selected_usages)]
        
        t_len, t_net_inv, t_vol, tot_npv, tot_irr, tot_irr_msg = get_analysis_result(filtered_df[num_cols].sum())

        st.subheader("2. 📊 선택 항목 합산 소계 (Subtotal)")
        
        m1, m2 = st.columns(2)
        m1.metric("최종 합산 NPV", f"{tot_npv:,.0f} 원")
        m2.metric("최종 합산 IRR", f"{tot_irr*100:.2f} %" if tot_irr is not None else tot_irr_msg)
        
        subtotal_df = pd.DataFrame([{
            "항목명": "☑️ 선택 용도 총합계",
            "총 투자길이(m)": t_len,
            "총 순투자액(원)": t_net_inv,
            "연간판매량(MJ)": t_vol,
            "NPV(원)": tot_npv,
            "IRR(%)": f"{tot_irr*100:.2f}%" if tot_irr is not None else tot_irr_msg
        }])
        st.dataframe(
            subtotal_df.style.format({"총 투자길이(m)": "{:,.1f}", "총 순투자액(원)": "{:,.0f}", "연간판매량(MJ)": "{:,.0f}", "NPV(원)": "{:,.0f}"}), 
            use_container_width=True, hide_index=True
        )

        st.divider()

        st.subheader("3. 📑 구간별 경제성 상세 명세서 (기본요금 검증 포함)")
        df_detail = filtered_df.groupby(['용도', '구간명'])[num_cols].sum().reset_index()
        
        detail_results = []
        for _, row in df_detail.iterrows():
            d_len, d_net_inv, d_vol, d_npv, d_irr, d_irr_msg = get_analysis_result(row)
            detail_results.append({
                "용도": row['용도'],
                "구간명": row['구간명'],
                "투자길이(m)": d_len,
                "공급전수(전)": row['총전수'],          
                "기본요금수익(원)": row['기본요금수익'], 
                "순투자액(원)": d_net_inv,
                "연간판매량(MJ)": d_vol,
                "NPV(원)": d_npv,
                "IRR(%)": f"{d_irr*100:.2f}%" if d_irr is not None else d_irr_msg
            })
            
        st.dataframe(pd.DataFrame(detail_results).style.format({
            "투자길이(m)": "{:,.1f}", 
            "공급전수(전)": "{:,.0f}",
            "기본요금수익(원)": "{:,.0f}",
            "순투자액(원)": "{:,.0f}", 
            "연간판매량(MJ)": "{:,.0f}", 
            "NPV(원)": "{:,.0f}"
        }), use_container_width=True, hide_index=True)

else:
    st.info("👆 분석을 시작하려면 결재용 Raw 파일을 올려주세요.")
