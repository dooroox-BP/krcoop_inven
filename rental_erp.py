import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime

# 1. 웹 페이지 레이아웃 설정
st.set_page_config(page_title="한국렌탈판매협동조합 통합 재고관리 ERP", layout="wide")

DB_FILE = "rental_erp_db.xlsx"

# 2. 엑셀 데이터베이스 로드 및 초기화
def init_database():
    if not os.path.exists(DB_FILE):
        init_products = pd.DataFrame([
            {"구분": "기기", "카테고리": "공기청정기", "제조사": "코웨이", "모델명": "AP-1019C", "타입": "해당없음", "기본입고가": 150000, "판매가": 250000},
            {"구분": "기기", "카테고리": "정수기", "제조사": "코웨이", "모델명": "CHP-264L", "타입": "해당없음", "기본입고가": 200000, "판매가": 350000},
            {"구분": "기기", "카테고리": "비데", "제조사": "노비타", "모델명": "BID-01A", "타입": "해당없음", "기본입고가": 100000, "판매가": 180000},
            {"구분": "필터", "카테고리": "공기청정기", "제조사": "코웨이", "모델명": "FT-AP10-ORG", "타입": "정품", "기본입고가": 15000, "판매가": 30000},
            {"구분": "필터", "카테고리": "공기청정기", "제조사": "중소기업", "모델명": "FT-AP10-COM", "타입": "호환", "기본입고가": 7000, "판매가": 15000},
        ])
        init_inventory = pd.DataFrame([
            {"모델명": "AP-1019C", "보관위치": "제1창고", "현재재고": 50},
            {"모델명": "CHP-264L", "보관위치": "제1창고", "현재재고": 30},
            {"모델명": "BID-01A", "보관위치": "제2창고", "현재재고": 20},
            {"모델명": "FT-AP10-ORG", "보관위치": "제1창고", "현재재고": 200},
            {"모델명": "FT-AP10-COM", "보관위치": "제2창고", "현재재고": 150},
        ])
        init_history = pd.DataFrame([
            {"날짜": "2025-06-15", "모델명": "AP-1019C", "구분": "입고", "수량": 10, "보관위치": "제1창고", "상태/흐름": "일반 입고 (+)", "변동입고가": 145000, "비고": "작년 특가 매입"},
            {"날짜": "2026-01-10", "모델명": "AP-1019C", "구분": "출고", "수량": 5, "보관위치": "제1창고", "상태/흐름": "일반 출고 (-)", "변동입고가": 150000, "비고": "연초 대리점 출고"},
            {"날짜": "2026-04-15", "모델명": "CHP-264L", "구분": "입고", "수량": 15, "보관위치": "제1창고", "상태/흐름": "일반 입고 (+)", "변동입고가": 200000, "비고": "정기 입고"},
            {"날짜": "2026-06-28", "모델명": "AP-1019C", "구분": "입고", "수량": 5, "보관위치": "제1창고", "상태/흐름": "일반 입고 (+)", "변동입고가": 155000, "비고": "원자재 상승 원가 반영"},
        ])
        
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            init_products.to_excel(writer, sheet_name="Products", index=False)
            init_inventory.to_excel(writer, sheet_name="Inventory", index=False)
            init_history.to_excel(writer, sheet_name="History", index=False)

    with pd.ExcelFile(DB_FILE, engine='openpyxl') as xls:
        st.session_state.products = pd.read_excel(xls, "Products")
        st.session_state.inventory = pd.read_excel(xls, "Inventory")
        st.session_state.history = pd.read_excel(xls, "History")

def save_database():
    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
        st.session_state.products.to_excel(writer, sheet_name="Products", index=False)
        st.session_state.inventory.to_excel(writer, sheet_name="Inventory", index=False)
        st.session_state.history.to_excel(writer, sheet_name="History", index=False)

if 'products' not in st.session_state:
    init_database()

# 데이터 융합 정형화
report_df = pd.merge(st.session_state.inventory, st.session_state.products, on="모델명", how="left")

def get_current_cost(row):
    hist_match = st.session_state.history[(st.session_state.history["모델명"] == row["모델명"]) & (st.session_state.history["구분"] == "입고")]
    if not hist_match.empty:
        return hist_match.sort_values(by="날짜").iloc[-1]["변동입고가"]
    return row["기본입고가"]

report_df["실제입고가"] = report_df.apply(get_current_cost, axis=1)
report_df["총입고가치"] = report_df["현재재고"] * report_df["실제입고가"]

# 메인 헤더
st.title("💼 스마트 렌탈·케어 통합 재고관리 ERP 시스템")
st.caption("한국렌탈판매협동조합 전용 솔루션 | 전체 선택 마스터 스위치 및 타이핑 검색 필터 탑재")
st.markdown("---")

# 대시보드 KPI 카드 요약
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📦 총 보유 재고 수량", f"{int(st.session_state.inventory['현재재고'].sum()):,} 개")
with col2:
    st.metric("💰 실시간 재고 자산 총액", f"{int(report_df['총입고가치'].sum()):,}")
with col3:
    pending_orders = len(st.session_state.history[st.session_state.history["상태/흐름"].str.contains("발주|수주", na=False)])
    st.metric("⏳ 진행 중인 발주/수주", f"{pending_orders} 건")
with col4:
    shortage_items = len(st.session_state.inventory[st.session_state.inventory["현재재고"] <= 15])
    st.metric("⚠️ 안전재고 부족 품목", f"{shortage_items} 종")

st.markdown("---")

# 5단계 기능별 확장 탭 메뉴 구성
tabs = st.tabs(["📊 실시간 재고 보고서 조회", "📈 기간별 변동 통계 분석", "🔄 입출고 및 발·수주 등록", "📑 품목 마스터(등록/삭제) 관리", "📜 누적 통합 장부 기록"])

# --- TAB 1: 실시간 재고 보고서 조회 ---
with tabs[0]:
    st.subheader("🔍 실시간 창고별 재고 현황판")
    
    # 🔍 통합 검색창 기능
    search_query = st.text_input("🔍 품목 통합 검색 (모델명, 제조사, 카테고리 입력)", "")
    
    # ⚙️ 엑셀식 타이핑 기반 입력형 고급 필터 시스템
    st.markdown("⚙️ **검색창 필터링 (글자를 입력하면 관련 항목만 자동으로 걸러집니다)**")
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        loc_input = st.text_input("📍 보관위치 필터 검색 (예: 제1창고)", "")
    with col_f2:
        cat_input = st.text_input("🗂️ 카테고리 필터 검색 (예: 공기청정기)", "")
    with col_f3:
        sep_input = st.text_input("📦 상품구분 필터 검색 (예: 기기 / 필터)", "")
    with col_f4:
        type_input = st.text_input("🧪 타입 필터 검색 (예: 정품 / 호환)", "")

    # 검색어 및 입력형 필터 적용
    filtered_df = report_df.copy()
    if search_query:
        filtered_df = filtered_df[
            filtered_df["모델명"].str.contains(search_query, case=False, na=False) |
            filtered_df["제조사"].str.contains(search_query, case=False, na=False) |
            filtered_df["카테고리"].str.contains(search_query, case=False, na=False)
        ]
    if loc_input:
        filtered_df = filtered_df[filtered_df["보관위치"].str.contains(loc_input, case=False, na=False)]
    if cat_input:
        filtered_df = filtered_df[filtered_df["카테고리"].str.contains(cat_input, case=False, na=False)]
    if sep_input:
        filtered_df = filtered_df[filtered_df["구분"].str.contains(sep_input, case=False, na=False)]
    if type_input:
        filtered_df = filtered_df[filtered_df["타입"].str.contains(type_input, case=False, na=False)]

    if not filtered_df.empty:
        base_cols = ["구분", "카테고리", "제조사", "모델명", "타입", "보관위치", "현재재고", "실제입고가", "판매가", "총입고가치"]
        select_df = filtered_df[base_cols].copy()
        
        # 🔳 [요청사항 반영] 상단 전체 선택 / 해제 마스터 체크박스 스위치
        st.markdown("---")
        select_all = st.checkbox("☑️ 현재 필터링된 모든 품목 전체 선택 / 해제", value=False)
        
        # 선택 기본값 처리
        select_df.insert(0, "선택", select_all)
        
        # 📊 테이블 크기 최적화 및 천단위 콤마 서식화 설정
        edited_df = st.data_editor(
            select_df,
            hide_index=True,
            use_container_width=False, # 휑하게 커지는 것을 방지하기 위해 고정 폭 스타일 결합
            column_config={
                "선택": st.column_config.CheckboxColumn("선택", default=False),
                "구분": st.column_config.Column("구분", width="medium"),
                "카테고리": st.column_config.Column("카테고리", width="medium"),
                "제조사": st.column_config.Column("제조사", width="medium"),
                "모델명": st.column_config.Column("모델명", width="medium"),
                "타입": st.column_config.Column("타입", width="medium"),
                "보관위치": st.column_config.Column("보관위치", width="medium"),
                "현재재고": st.column_config.NumberColumn("현재재고", format="%d"),
                "실제입고가": st.column_config.NumberColumn("실제입고가", format="%,d"),
                "판매가": st.column_config.NumberColumn("판매가", format="%,d"),
                "총입고가치": st.column_config.NumberColumn("총입고가치", format="%,d")
            }
        )
        
        # 🎨 [요청사항 반영] 고도화된 CSS 주입 구조
        # 1. 일반 문자열은 왼쪽을 넉넉히 띄워 중앙 정렬 효과(Padding) 유도
        # 2. 현재재고는 가운데 정렬을 유지하되 글자 끝 위치를 정밀 정렬 수직 매칭
        # 3. 금액은 완벽한 우측 회계 정렬 처리
        st.markdown("""
            <style>
                /* 문자열 컬럼 구역 (2번째 열부터 6번째 열까지) 왼쪽 여백 부여 및 슬림화 */
                div[data-testid="stDataFrame"] td:nth-child(2), 
                div[data-testid="stDataFrame"] td:nth-child(3), 
                div[data-testid="stDataFrame"] td:nth-child(4), 
                div[data-testid="stDataFrame"] td:nth-child(5), 
                div[data-testid="stDataFrame"] td:nth-child(6),
                div[data-testid="stDataFrame"] td:nth-child(7) { 
                    text-align: left !important; 
                    justify-content: flex-start !important; 
                    padding-left: 25px !important; 
                }
                
                /* 현재재고 컬럼 (7번째 열) - 가운데 정렬을 하되 끝숫자 수직 줄맞춤 강제 최적화 */
                div[data-testid="stDataFrame"] td:nth-child(8) { 
                    text-align: center !important; 
                    justify-content: center !important;
                    font-variant-numeric: tabular-nums !important;
                    padding-right: 20px !important;
                }
                
                /* 금액 컬럼 구역 (실제입고가, 판매가, 총입고가치) 우측 밀착 정렬 */
                div[data-testid="stDataFrame"] td:nth-child(9), 
                div[data-testid="stDataFrame"] td:nth-child(10), 
                div[data-testid="stDataFrame"] td:nth-child(11) { 
                    text-align: right !important; 
                    justify-content: flex-end !important; 
                    padding-right: 25px !important;
                    font-variant-numeric: tabular-nums !important;
                }
            </style>
        """, unsafe_allow_html=True)

        # 🧮 선택 항목 실시간 연산 로직
        checked_rows = edited_df[edited_df["선택"] == True]
        
        st.markdown("### 📊 선택 품목 실시간 합계 행 (Total)")
        
        total_qty = checked_rows["현재재고"].sum() if not checked_rows.empty else 0
        total_value = checked_rows["총입고가치"].sum() if not checked_rows.empty else 0
        
        t_col1, t_col2, t_col3 = st.columns([4, 4, 4])
        with t_col1:
            st.markdown(f"🏁 **선택된 품목 개수:** {len(checked_rows)} 종")
        with t_col2:
            st.markdown(f"🔢 **현재재고 합계 (가운데 줄맞춤):** <span style='color:#2563eb; font-weight:bold; font-size:18px;'>{total_qty:,}</span> 개", unsafe_allow_html=True)
        with t_col3:
            st.markdown(f"💵 **자산가치 합계 (우측 정렬단):** <span style='color:#16a34a; font-weight:bold; font-size:18px;'>{total_value:,}</span>", unsafe_allow_html=True)
            
        st.markdown("---")
        
        excel_df = filtered_df[base_cols].copy()
        csv = excel_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 필터링된 현재고 종합 보고서 다운로드 (엑셀 추출)", data=csv, file_name=f"inventory_report_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")
    else:
        st.info("검색 조건이나 필터에 맞는 재고 데이터가 없습니다.")

# --- TAB 2: 기간별 변동 통계 분석 ---
with tabs[1]:
    st.subheader("📈 시점별 재고 입출고 가감 분석 리포트")
    
    if st.session_state.history.empty:
        st.info("누적된 입출고 이력이 없어 분석할 데이터가 없습니다.")
    else:
        hist_df = st.session_state.history.copy()
        hist_df["날짜"] = pd.to_datetime(hist_df["날짜"])
        
        period_type = st.radio("📊 분석 주기 단위 선택", ["일별", "주별", "월별", "분기별", "반기별", "연도별"], horizontal=True)
        
        if period_type == "일별":
            hist_df["기간"] = hist_df["날짜"].dt.strftime("%Y-%m-%d")
        elif period_type == "주별":
            hist_df["기간"] = hist_df["날짜"].dt.strftime("%Y-%U주")
        elif period_type == "월별":
            hist_df["기간"] = hist_df["날짜"].dt.strftime("%Y-%m")
        elif period_type == "분기별":
            hist_df["기간"] = hist_df["날짜"].dt.year.astype(str) + "-Q" + hist_df["날짜"].dt.quarter.astype(str)
        elif period_type == "반기별":
            hist_df["기간"] = hist_df["날짜"].dt.year.astype(str) + "-" + hist_df["날짜"].dt.apply(lambda x: "1반기" if x.month <= 6 else "2반기")
        elif period_type == "연도별":
            hist_df["기간"] = hist_df["날짜"].dt.strftime("%Y년")
            
        hist_df["입고량"] = hist_df.apply(lambda r: r["수량"] if r["구분"] == "입고" else 0, axis=1)
        hist_df["출고량"] = hist_df.apply(lambda r: r["수량"] if r["구분"] == "출고" else 0, axis=1)
        
        summary_period = hist_df.groupby(["기간", "모델명"])[["입고량", "출고량"]].sum().reset_index()
        
        fig_period = px.bar(summary_period, x="기간", y=["입고량", "출고량"], barmode="group", title=f"📊 {period_type} 제품 모델별 가감 트렌드 분석")
        st.plotly_chart(fig_period, use_container_width=True)
        
        st.dataframe(
            summary_period,
            use_container_width=True,
            hide_index=True,
            column_config={
                "입고량": st.column_config.NumberColumn("총 입고량", format="%,d"),
                "출고량": st.column_config.NumberColumn("총 출고량", format="%,d")
            }
        )

# --- TAB 3: 입출고 및 발·수주 등록 (대문자 승환 장치 연동) ---
with tabs[2]:
    st.subheader("🔄 일일 재고 변동 가감 및 프로세스 입력 등록")
    
    with st.form("process_form", clear_on_submit=True):
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            f_date = st.date_input("업무 일자 선택", datetime.now())
            f_model = st.selectbox("대상 품목 모델명 선택", options=st.session_state.products["모델명"].unique())
        with col_b:
            f_flow = st.selectbox("업무 프로세스 선택", [
                "일반 입고 (+)", 
                "일반 출고 (-)", 
                "[발주] 공급처 물품요청", 
                "[발주] 납기일자 확인완료", 
                "[수주] 대리점 주문등록 (출고대기)",
                "[렌탈] 현장 설치 출고 (-)"
            ])
            f_qty = st.number_input("수량 입력 (개)", min_value=1, value=1, step=1)
        with col_c:
            current_base_price = int(st.session_state.products[st.session_state.products["모델명"] == f_model]["기본입고가"].values[0])
            f_price = st.number_input("이번 단가 입력 (원가 변동 시 직접 수정)", min_value=0, value=current_base_price, step=1000)
            f_loc = st.text_input("적용 창고 위치 지정", value="제1창고")
            f_note = st.text_input("거래처명 및 비고 적기")
            
        submit_btn = st.form_submit_button("💼 시스템 장부에 실시간 데이터 반영")
        
        if submit_btn:
            # 🔠 [요청사항 반영] 소문자로 들어와도 대문자로 강제 승환 가공
            f_model_upper = f_model.upper().strip()
            f_loc_upper = f_loc.upper().strip()
            
            is_reducing = f_flow in ["일반 출고 (-)", "[렌탈] 현장 설치 출고 (-)"]
            is_increasing = f_flow in ["일반 입고 (+)"]
            
            calc_qty = -f_qty if is_reducing else (f_qty if is_increasing else 0)
            mask = (st.session_state.inventory["모델명"] == f_model_upper) & (st.session_state.inventory["보관위치"] == f_loc_upper)
            
            if is_reducing and (not mask.any() or st.session_state.inventory.loc[mask, "현재재고"].values[0] < f_qty):
                st.error("❌ 창고에 실제 출고 가능한 재고 수량이 모자라 등록이 거부되었습니다.")
            else:
                if calc_qty != 0:
                    if mask.any():
                        st.session_state.inventory.loc[mask, "현재재고"] += calc_qty
                    else:
                        new_inv = pd.DataFrame([{"모델명": f_model_upper, "보관위치": f_loc_upper, "현재재고": f_qty}])
                        st.session_state.inventory = pd.concat([st.session_state.inventory, new_inv], ignore_index=True)
                
                new_hist = pd.DataFrame([{
                    "날짜": f_date.strftime("%Y-%m-%d"), "모델명": f_model_upper,
                    "구분": "출고" if is_reducing else ("입고" if is_increasing else "프로세스"),
                    "수량": f_qty, "보관위치": f_loc_upper, "상태/흐름": f_flow,
                    "변동입고가": f_price, "비고": f_note
                }])
                st.session_state.history = pd.concat([st.session_state.history, new_hist], ignore_index=True)
                
                save_database()
                st.success(f"✅ 반영 성공: {f_model_upper} {f_flow} {f_qty}개 등록완료")
                st.rerun()

# --- TAB 4: 품목 마스터(등록/삭제) 관리 (대문자 승환 안전장치 연동) ---
with tabs[3]:
    st.subheader("📑 취급 품목 마스터 무제한 등록 및 삭제")
    
    col_reg, col_del = st.columns([2, 1])
    
    with col_reg:
        st.write("➕ **신규 취급 품목 시스템 추가**")
        with st.form("product_reg_form", clear_on_submit=True):
            r_model = st.text_input("새로운 모델명 입력 (필수, 소문자 입력 시 대문자 자동 전환)")
            r_sep = st.radio("상품 분류", ["기기", "필터"], horizontal=True)
            r_cat = st.selectbox("상세 카테고리 고르기", ["공기청정기", "정수기", "비데", "소모품/기타"])
            r_brand = st.text_input("제조사 브랜드 입력", value="코웨이")
            r_type = st.selectbox("필터인 경우 타입 지정", ["해당없음", "정품", "호환"])
            r_in_price = st.number_input("기본 매입 단가 입력 (원가)", min_value=0, value=0, step=1000)
            r_out_price = st.number_input("기본 판매/렌탈 단가 입력", min_value=0, value=0, step=1000)
            
            reg_btn = st.form_submit_button("🆕 시스템 마스터에 신규 품목 등록")
            if reg_btn:
                # 🔠 [요청사항 반영] 품목 등록 즉시 대문자 승환
                r_model_upper = r_model.upper().strip()
                
                if r_model_upper in st.session_state.products["모델명"].values:
                    st.error("❌ 이미 데이터베이스에 존재하는 중복된 모델명입니다.")
                elif r_model_upper == "":
                    st.error("❌ 모델명은 공백으로 둘 수 없습니다.")
                else:
                    new_prod = pd.DataFrame([{
                        "모델명": r_model_upper, "구분": r_sep, "카테고리": r_cat,
                        "제조사": r_brand, "타입": r_type, "기본입고가": r_in_price, "판매가": r_out_price
                    }])
                    st.session_state.products = pd.concat([st.session_state.products, new_prod], ignore_index=True)
                    save_database()
                    st.success(f"🎉 신규 품목 마스터에 [{r_model_upper}] 등록 성공!")
                    st.rerun()
                    
    with col_del:
        st.write("🗑️ **품목 데이터베이스 영구 삭제**")
        del_model = st.selectbox("삭제해 버릴 품목 고르기", options=["선택안함"] + list(st.session_state.products["모델명"].unique()))
        del_btn = st.button("❌ 선택한 품목 영구 삭제 단행")
        
        if del_btn:
            if del_model == "선택안함":
                st.warning("리스트에서 삭제 대상을 먼저 마우스로 선택하세요.")
            else:
                st.session_state.products = st.session_state.products[st.session_state.products["모델명"] != del_model]
                st.session_state.inventory = st.session_state.inventory[st.session_state.inventory["모델명"] != del_model]
                save_database()
                st.success(f"🗑️ 품목 [{del_model}]이 시스템에서 제거되었습니다.")
                st.rerun()
                
    st.markdown("---")
    st.write("📑 **현재 시스템에 등록되어 가동 중인 마스터 품목 전체 데이터 정보**")
    
    st.dataframe(
        st.session_state.products,
        use_container_width=True,
        hide_index=True,
        column_config={
            "기본입고가": st.column_config.NumberColumn("기본입고가", format="%,d"),
            "판매가": st.column_config.NumberColumn("판매가", format="%,d")
        }
    )

# --- TAB 5: 누적 통합 장부 기록 ---
with tabs[4]:
    st.subheader("📜 전체 누적 비즈니스 이력 장부 기록")
    
    if len(st.session_state.history) == 0:
        st.info("장부에 누적된 변동 데이터 이력이 아직 없습니다.")
    else:
        st.dataframe(
            st.session_state.history.sort_values(by="날짜", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "수량": st.column_config.NumberColumn("수량", format="%,d"),
                "변동입고가": st.column_config.NumberColumn("변동입고가", format="%,d")
            }
        )