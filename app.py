import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

# --------------------------------------------------------------------------
# [설정] 페이지 기본
# --------------------------------------------------------------------------
st.set_page_config(page_title="신규배관 경제성 분석 Simulation", page_icon="🏗️", layout="wide")

# --------------------------------------------------------------------------
# [함수] 금융 계산 로직
# --------------------------------------------------------------------------
def manual_npv(rate, values):
    return sum(v / ((1 + rate) ** i) for i, v in enumerate(values))

def calculate_simulation(sim_len, sim_inv, sim_contrib, sim_other, sim_vol, sim_rev, sim_cost, 
                         sim_jeon, sim_basic_rev, rate, tax, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m):
    
    # 1. 초기 순투자액 (Year 0)
    net_inv = sim_inv - sim_contrib - sim_other
    
    # 2. 고정 수익/비용 항목 계산
    margin_total = (sim_rev - sim_cost) + sim_basic_rev 
    unit_margin = margin_total / sim_vol if sim_vol > 0 else 0
    cost_sga = (sim_len * c_maint) + (sim_len * c_adm_m) + (sim_jeon * c_adm_jeon)
    annual_depreciation = sim_inv / dep_period if dep_period > 0 else 0
    
    # 3. 세후 현금흐름(OCF) 산출 (분석기간 내 고정값)
    ebit = margin_total - cost_sga - annual_depreciation
    net_income = ebit * (1 - tax)
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
# [함수] 데이터 컬럼 유연 매핑 (띄어쓰기/특수문자 무시)
# --------------------------------------------------------------------------
def find_col(df, keywords):
    for col in df.columns:
        col_str = str(col).replace(" ", "").replace("\n", "")
        for kw in keywords:
            if kw in col_str:
                return col
    return None

# --------------------------------------------------------------------------
# [UI] 좌측 사이드바
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 분석 변수")
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
st.title("🏗️ 신규배관 경제성 분석 다중 시뮬레이션")
st.markdown("Raw 데이터(엑셀/CSV)를 업로드하면 **구간명** 기준으로 전체/용도별 경제성(NPV, IRR)을 일괄 분석합니다.")

uploaded_file = st.file_uploader("📂 Raw 데이터 파일 업로드 (Excel 또는 CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    # 1. 파일 읽기
    try:
        if uploaded_file.name.endswith('.csv'):
            # 도시가스 raw 파일 특성상 앞부분 로우가 지저분할 수 있으므로, 
            # 필요에 따라 skiprows=1 또는 2를 적용하실 수 있습니다.
            df = pd.read_csv(uploaded_file, header=1) 
        else:
            df = pd.read_excel(uploaded_file, header=1)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    # 2. 필수 컬럼 자동 찾기
    col_name = find_col(df, ["구간명"])
    col_len = find_col(df, ["길이", "배관길이"])
    col_inv = find_col(df, ["배관투자금액", "총공사비"])
    col_contrib = find_col(df, ["시설분담금"])
    col_other = find_col(df, ["기타이익", "보조금"])
    col_jeon = find_col(df, ["수요전수계", "총전수"])
    col_jeon_home = find_col(df, ["가정용", "주택용전수", "공동주택전수"]) # 주택용 전수 파악용
    col_vol = find_col(df, ["연간판매량", "판매량"])
    col_rev = find_col(df, ["연간판매액", "판매액"])
    col_cost = find_col(df, ["연간판매원가", "판매원가"])

    if not col_name:
        st.error("❌ 데이터에서 '구간명' 컬럼을 찾을 수 없습니다. 원본 파일의 헤더 위치를 확인해 주세요.")
        st.stop()

    st.success("✅ 파일 업로드 및 컬럼 매핑 완료! 데이터를 분석합니다.")
    
    # 결측치 0으로 처리
    df.fillna(0, inplace=True)
    
    # 3. 구간명 기준으로 그룹화 (Numeric 컬럼만 합산)
    numeric_cols = [c for c in [col_len, col_inv, col_contrib, col_other, col_jeon, col_jeon_home, col_vol, col_rev, col_cost] if c is not None]
    df_grouped = df.groupby(col_name)[numeric_cols].sum().reset_index()

    # 4. 분석 결과 저장을 위한 리스트
    results = []

    for index, row in df_grouped.iterrows():
        section_name = row[col_name]
        if str(section_name).strip() == "0" or str(section_name).strip() == "":
            continue
            
        s_len = row[col_len] if col_len else 0
        s_inv = row[col_inv] if col_inv else 0
        s_contrib = row[col_contrib] if col_contrib else 0
        s_other = row[col_other] if col_other else 0
        s_jeon = row[col_jeon] if col_jeon else 0
        s_jeon_home = row[col_jeon_home] if col_jeon_home else 0
        s_vol = row[col_vol] if col_vol else 0
        s_rev = row[col_rev] if col_rev else 0
        s_cost = row[col_cost] if col_cost else 0
        
        # 주택용 기본요금 계산 (가정용 전수 기준)
        s_basic_rev = sim_basic_price * s_jeon_home * 12

        # 계산 실행
        npv, irr, irr_msg, flows = calculate_simulation(
            s_len, s_inv, s_contrib, s_other, s_vol, s_rev, s_cost, 
            s_jeon, s_basic_rev, RATE, TAX, dep_period, analysis_period, 
            c_maint, c_adm_jeon, c_adm_m
        )
        
        results.append({
            "구간명": section_name,
            "투자길이(m)": s_len,
            "순투자액(원)": s_inv - s_contrib - s_other,
            "연간판매량(MJ)": s_vol,
            "NPV(원)": npv,
            "IRR(%)": f"{irr*100:.2f}%" if irr is not None else irr_msg
        })

    # 전체 합산용 데이터 (Total)
    t_len = df_grouped[col_len].sum() if col_len else 0
    t_inv = df_grouped[col_inv].sum() if col_inv else 0
    t_contrib = df_grouped[col_contrib].sum() if col_contrib else 0
    t_other = df_grouped[col_other].sum() if col_other else 0
    t_jeon = df_grouped[col_jeon].sum() if col_jeon else 0
    t_jeon_home = df_grouped[col_jeon_home].sum() if col_jeon_home else 0
    t_vol = df_grouped[col_vol].sum() if col_vol else 0
    t_rev = df_grouped[col_rev].sum() if col_rev else 0
    t_cost = df_grouped[col_cost].sum() if col_cost else 0
    t_basic_rev = sim_basic_price * t_jeon_home * 12

    tot_npv, tot_irr, tot_irr_msg, _ = calculate_simulation(
        t_len, t_inv, t_contrib, t_other, t_vol, t_rev, t_cost, 
        t_jeon, t_basic_rev, RATE, TAX, dep_period, analysis_period, 
        c_maint, c_adm_jeon, c_adm_m
    )

    # 5. 결과 출력
    st.subheader("📊 전체 합산 경제성 결과")
    m1, m2 = st.columns(2)
    m1.metric("총 순현재가치 (Total NPV)", f"{tot_npv:,.0f} 원")
    m2.metric("총 내부수익률 (Total IRR)", f"{tot_irr*100:.2f} %" if tot_irr is not None else tot_irr_msg)

    st.divider()

    st.subheader("📑 구간명별 세부 분석 결과")
    df_results = pd.DataFrame(results)
    
    # 숫자 포맷팅을 위해 Styler 적용
    st.dataframe(
        df_results.style.format({
            "투자길이(m)": "{:,.1f}",
            "순투자액(원)": "{:,.0f}",
            "연간판매량(MJ)": "{:,.0f}",
            "NPV(원)": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )
    
else:
    st.info("👆 분석을 시작하려면 좌측 또는 상단의 업로드 영역에 Raw 파일을 올려주세요.")
