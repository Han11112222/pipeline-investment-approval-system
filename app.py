import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf

# --------------------------------------------------------------------------
# [설정] 페이지 기본
# --------------------------------------------------------------------------
st.set_page_config(page_title="공식 배관 투자 결재 시스템 (Pipeline Approval)", page_icon="🏗️", layout="wide")

# --------------------------------------------------------------------------
# [함수] 금융 계산 로직 (파이썬이 직접 계산!)
# --------------------------------------------------------------------------
def manual_npv(rate, values):
    return sum(v / ((1 + rate) ** i) for i, v in enumerate(values))

def calculate_simulation(sim_len, sim_inv, sim_contrib, sim_other, sim_vol, sim_rev, sim_cost, 
                         sim_jeon, sim_basic_rev, rate, tax, dep_period, analysis_period, c_maint, c_adm_jeon, c_adm_m):
    
    net_inv = sim_inv - sim_contrib - sim_other
    margin_total = (sim_rev - sim_cost) + sim_basic_rev 
    unit_margin = margin_total / sim_vol if sim_vol > 0 else 0
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

# --------------------------------------------------------------------------
# [함수] 데이터 컬럼 유연 매핑
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
st.markdown("Raw 데이터를 기반으로 파이썬 엔진이 직접 세후 현금흐름을 시뮬레이션하여 **용도별/구간별 NPV와 IRR**을 계산합니다.")

uploaded_file = st.file_uploader("📂 결재용 Raw 데이터 파일 업로드 (Excel 또는 CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=None) 
        else:
            df = pd.read_excel(uploaded_file, header=None)
            
        new_cols = []
        for i in range(len(df.columns)):
            val0 = str(df.iloc[0, i]).replace("nan", "").strip() if pd.notna(df.iloc[0, i]) else ""
            val1 = str(df.iloc[1, i]).replace("nan", "").strip() if pd.notna(df.iloc[1, i]) else ""
            new_cols.append(f"{val0}_{val1}") 
        df.columns = new_cols
        df = df.iloc[2:].reset_index(drop=True)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    # 필수 컬럼 찾기 (용도 추가)
    col_usage = find_col(df, ["용도", "가스용도"]) 
    col_name = find_col(df, ["구간명"])
    col_len = find_col(df, ["길이", "배관길이"])
    col_inv = find_col(df, ["배관투자금액", "총공사비"])
    col_contrib = find_col(df, ["시설분담금"])
    col_other = find_col(df, ["기타이익", "보조금"])
    col_jeon = find_col(df, ["수요전수계", "총전수"])
    col_jeon_home = find_col(df, ["가정용", "주택용전수", "공동주택전수"]) 
    col_vol = find_col(df, ["연간판매량", "판매량", "계(MJ)"]) 
    col_rev = find_col(df, ["연간판매액", "판매액"])
    col_cost = find_col(df, ["연간판매원가", "판매원가"])

    if not col_name:
        st.error("❌ 데이터에서 '구간명' 관련 컬럼을 찾을 수 없습니다.")
        st.stop()
        
    # 용도 컬럼이 없을 경우 임시 부여
    if not col_usage:
        df['임시_용도'] = '미분류(용도없음)'
        col_usage = '임시_용도'
    else:
        # 용도 데이터의 빈칸 채우기 (엑셀 병합셀 고려)
        df[col_usage] = df[col_usage].ffill()

    numeric_cols = [c for c in [col_len, col_inv, col_contrib, col_other, col_jeon, col_jeon_home, col_vol, col_rev, col_cost] if c is not None]
    
    for c in numeric_cols:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.replace(',', '', regex=False)
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
    # 1. 구간별 상세 (용도 + 구간명)
    df_detail = df.groupby([col_usage, col_name])[numeric_cols].sum().reset_index()
    
    # 2. 용도별 합산 (용도만)
    df_usage = df.groupby(col_usage)[numeric_cols].sum().reset_index()

    st.success("✅ 파이썬 시뮬레이션 엔진 구동 완료! 결과를 출력합니다.")
    
    # 계산 실행용 헬퍼 함수
    def get_analysis_result(row):
        s_len = row[col_len] if col_len else 0
        s_inv = row[col_inv] if col_inv else 0
        s_contrib = row[col_contrib] if col_contrib else 0
        s_other = row[col_other] if col_other else 0
        s_jeon = row[col_jeon] if col_jeon else 0
        s_jeon_home = row[col_jeon_home] if col_jeon_home else 0
        s_vol = row[col_vol] if col_vol else 0
        s_rev = row[col_rev] if col_rev else 0
        s_cost = row[col_cost] if col_cost else 0
        
        s_basic_rev = sim_basic_price * s_jeon_home * 12

        npv, irr, irr_msg, _ = calculate_simulation(
            s_len, s_inv, s_contrib, s_other, s_vol, s_rev, s_cost, 
            s_jeon, s_basic_rev, RATE, TAX, dep_period, analysis_period, 
            c_maint, c_adm_jeon, c_adm_m
        )
        return s_len, s_inv - s_contrib - s_other, s_vol, npv, irr, irr_msg

    # --- 전체 합산 결과 ---
    t_len, t_net_inv, t_vol, tot_npv, tot_irr, tot_irr_msg = get_analysis_result(df[numeric_cols].sum())
    
    st.subheader("1. 📊 전체 통합 합산 결과 (Grand Total)")
    m1, m2 = st.columns(2)
    m1.metric("총 순현재가치 (Total NPV)", f"{tot_npv:,.0f} 원")
    m2.metric("총 내부수익률 (Total IRR)", f"{tot_irr*100:.2f} %" if tot_irr is not None else tot_irr_msg)
    st.divider()

    # --- 용도별 요약 ---
    st.subheader("2. 📁 용도별 경제성 요약")
    usage_results = []
    for _, row in df_usage.iterrows():
        usage_name = str(row[col_usage]).replace("nan", "미분류").strip()
        if not usage_name or usage_name == "0": continue
        
        u_len, u_net_inv, u_vol, u_npv, u_irr, u_irr_msg = get_analysis_result(row)
        usage_results.append({
            "용도": usage_name,
            "총 투자길이(m)": u_len,
            "총 순투자액(원)": u_net_inv,
            "연간판매량(MJ)": u_vol,
            "NPV(원)": u_npv,
            "IRR(%)": f"{u_irr*100:.2f}%" if u_irr is not None else u_irr_msg
        })
    
    st.dataframe(pd.DataFrame(usage_results).style.format({"총 투자길이(m)": "{:,.1f}", "총 순투자액(원)": "{:,.0f}", "연간판매량(MJ)": "{:,.0f}", "NPV(원)": "{:,.0f}"}), use_container_width=True, hide_index=True)
    st.divider()

    # --- 구간별 상세 ---
    st.subheader("3. 📑 용도-구간별 경제성 상세 명세서")
    detail_results = []
    for _, row in df_detail.iterrows():
        usage_name = str(row[col_usage]).replace("nan", "미분류").strip()
        section_name = str(row[col_name]).replace("nan", "").strip()
        
        if not section_name or section_name == "0": continue
        
        d_len, d_net_inv, d_vol, d_npv, d_irr, d_irr_msg = get_analysis_result(row)
        detail_results.append({
            "용도": usage_name,
            "구간명": section_name,
            "투자길이(m)": d_len,
            "순투자액(원)": d_net_inv,
            "연간판매량(MJ)": d_vol,
            "NPV(원)": d_npv,
            "IRR(%)": f"{d_irr*100:.2f}%" if d_irr is not None else d_irr_msg
        })
        
    st.dataframe(pd.DataFrame(detail_results).style.format({"투자길이(m)": "{:,.1f}", "순투자액(원)": "{:,.0f}", "연간판매량(MJ)": "{:,.0f}", "NPV(원)": "{:,.0f}"}), use_container_width=True, hide_index=True)

else:
    st.info("👆 분석을 시작하려면 결재용 Raw 파일을 올려주세요.")
