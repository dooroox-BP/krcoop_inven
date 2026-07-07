"""
app.py
------
인천광역시협동조합협의회 결산 입력 프로그램 (Streamlit)

실행:
    streamlit run app.py

Excel 파일을 데이터 저장소로 사용한다.
프로그램 실행 중에는 데이터 엑셀 파일을 열어두지 마세요(저장 충돌 방지).
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from src import (
    backup_service,
    config,
    excel_service,
    report_service,
    transaction_service,
    ui_components as ui,
    validation_service,
)

st.set_page_config(page_title=f"{config.ORG_NAME} 결산 입력 프로그램", layout="wide")


# ---------------------------------------------------------------------------
# 공통 헬퍼
# ---------------------------------------------------------------------------

def load_wb():
    """디스크에서 워크북을 로드(없으면 구조만 생성)."""
    return excel_service.load_or_create_workbook()


def available_years(wb) -> list[int]:
    """대장에 존재하는 회계연도 목록(내림차순). 없으면 기본연도."""
    df = excel_service.read_ledger_df(wb)
    years = sorted({int(y) for y in df["회계연도"] if int(y) > 0}, reverse=True) \
        if not df.empty else []
    if config.DEFAULT_YEAR not in years:
        years = [config.DEFAULT_YEAR] + years
    return years


def persist_and_reports(wb, year: int):
    """
    보고서/검증표를 재생성하고 백업 후 저장한다.
    반환: 백업 파일 경로(또는 None).
    """
    report_service.generate_all_reports(wb, year)
    validation_service.generate_validation_report(wb, year)
    _saved, backup = excel_service.save_with_backup(wb)
    return backup


def flash(kind: str, msg: str):
    """다음 rerun 에서 표시할 메시지를 저장한다."""
    st.session_state["_flash"] = (kind, msg)


def show_flash():
    data = st.session_state.pop("_flash", None)
    if not data:
        return
    kind, msg = data
    {"success": st.success, "warning": st.warning,
     "error": st.error, "info": st.info}.get(kind, st.info)(msg)


def year_selector(wb, key: str) -> int:
    years = available_years(wb)
    return st.selectbox("회계연도", years, index=0, key=key)


# ---------------------------------------------------------------------------
# 1. 대시보드
# ---------------------------------------------------------------------------

def page_dashboard(wb):
    st.header("📊 대시보드")
    year = year_selector(wb, "dash_year")

    summary = report_service.summary_from_workbook(wb, year)
    _rows, overall, _ = validation_service.build_validation(wb, year)

    ui.dashboard_cards(summary, overall)

    st.divider()
    df = excel_service.read_ledger_df(wb)
    inc_major, _, inc_total = report_service.income_breakdown(df, year)
    exp_major, _, exp_total = report_service.expense_breakdown(df, year)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("수입 대분류별")
        ui.show_table(inc_major)
        st.caption(f"수입합계: {ui.format_won(inc_total)}")
    with c2:
        st.subheader("지출 대분류별")
        ui.show_table(exp_major)
        st.caption(f"지출합계: {ui.format_won(exp_total)}")


# ---------------------------------------------------------------------------
# 2. 거래 입력
# ---------------------------------------------------------------------------

def page_input(wb):
    st.header("📝 거래 입력")
    st.caption("입출구분 → 대분류 → 소분류 순으로 선택하면 해당 항목만 표시됩니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        year = st.number_input("회계연도 *", min_value=2000, max_value=2100,
                               value=config.DEFAULT_YEAR, step=1)
        tx_date = st.date_input("거래일자 *", value=date.today())
        inout = st.selectbox("입출구분 *", config.INOUT_TYPES, key="in_inout")
    with c2:
        majors = config.majors_for_inout(inout)
        major = st.selectbox("대분류 *", majors, key="in_major")
        minors = config.minors_for(inout, major)
        minor = st.selectbox("소분류 *", minors, key="in_minor")
        amount = st.number_input("금액 * (원)", min_value=0, step=1000, value=0)
    with c3:
        vendor = st.text_input("거래처")
        business = st.selectbox("사업구분", [""] + config.BUSINESS_TYPES)
        method = st.selectbox("결제수단", [""] + config.PAYMENT_METHODS)

    c4, c5, c6 = st.columns(3)
    with c4:
        proof = st.selectbox("증빙여부", config.PROOF_VALUES)
    with c5:
        proof_link = st.text_input("증빙링크")
    with c6:
        writer = st.text_input("입력자", value="관리자")

    note = st.text_area("비고", height=80)

    if st.button("💾 거래 저장", type="primary"):
        data = {
            "회계연도": int(year),
            "거래일자": tx_date,
            "입출구분": inout,
            "대분류": major,
            "소분류": minor,
            "금액": int(amount),
            "거래처": vendor,
            "사업구분": business,
            "결제수단": method,
            "증빙여부": proof,
            "증빙링크": proof_link,
            "비고": note,
            "입력자": writer,
        }
        tx_id, errors = transaction_service.add_transaction(wb, data)
        if errors:
            ui.show_errors(errors)
        else:
            backup = persist_and_reports(wb, int(year))
            bname = backup.name if backup else "생성 안 함(신규 파일)"
            flash("success", f"거래 저장 완료 · 거래ID: {tx_id} · 백업: {bname}")
            st.rerun()


# ---------------------------------------------------------------------------
# 3. 거래 조회·수정
# ---------------------------------------------------------------------------

def page_manage(wb):
    st.header("🔎 거래 조회 · 수정")
    df = excel_service.read_ledger_df(wb)

    if df.empty:
        st.info("등록된 거래가 없습니다.")
        return

    # ---- 필터 ----
    with st.expander("필터", expanded=True):
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            years = ["(전체)"] + [str(y) for y in available_years(wb)]
            f_year = st.selectbox("회계연도", years)
            f_status = st.selectbox("상태", ["(전체)"] + config.STATUS_VALUES)
        with f2:
            f_inout = st.selectbox("입출구분", ["(전체)"] + config.INOUT_TYPES)
            all_majors = sorted({m for m in df["대분류"] if m})
            f_major = st.selectbox("대분류", ["(전체)"] + all_majors)
        with f3:
            f_business = st.selectbox("사업구분", ["(전체)"] + config.BUSINESS_TYPES)
            f_vendor = st.text_input("거래처 포함")
        with f4:
            f_start = st.text_input("시작일 (YYYY-MM-DD)")
            f_end = st.text_input("종료일 (YYYY-MM-DD)")

    filtered = transaction_service.filter_transactions(
        df,
        year=None if f_year == "(전체)" else int(f_year),
        inout=None if f_inout == "(전체)" else f_inout,
        major=None if f_major == "(전체)" else f_major,
        business=None if f_business == "(전체)" else f_business,
        vendor=f_vendor or None,
        start_date=f_start or None,
        end_date=f_end or None,
        status=None if f_status == "(전체)" else f_status,
    )

    st.caption(f"조회 결과: {len(filtered)}건")
    ui.show_table(filtered)

    st.divider()
    st.subheader("거래 수정 / 삭제")

    if filtered.empty:
        st.info("수정할 거래를 필터로 선택하세요.")
        return

    tx_ids = filtered["거래ID"].tolist()
    sel_id = st.selectbox("수정할 거래ID", tx_ids, key="manage_sel_id")
    row = df[df["거래ID"] == sel_id].iloc[0]

    # 위젯 key 에 sel_id 를 붙여, 다른 거래를 선택하면 편집 폼이 새로 채워지도록 한다.
    sfx = f"_{sel_id}"

    e1, e2, e3 = st.columns(3)
    with e1:
        e_inout = st.selectbox("입출구분", config.INOUT_TYPES,
                               index=_safe_index(config.INOUT_TYPES, row["입출구분"]),
                               key="ed_inout" + sfx)
        e_majors = config.majors_for_inout(e_inout)
        e_major = st.selectbox("대분류", e_majors,
                               index=_safe_index(e_majors, row["대분류"]),
                               key="ed_major" + sfx)
        e_minors = config.minors_for(e_inout, e_major)
        e_minor = st.selectbox("소분류", e_minors,
                               index=_safe_index(e_minors, row["소분류"]),
                               key="ed_minor" + sfx)
    with e2:
        e_amount = st.number_input("금액 (원)", min_value=0, step=1000,
                                   value=int(float(row["금액"])), key="ed_amount" + sfx)
        e_date = st.text_input("거래일자", value=str(row["거래일자"]),
                               key="ed_date" + sfx)
        e_vendor = st.text_input("거래처", value=str(row["거래처"]),
                                 key="ed_vendor" + sfx)
    with e3:
        e_business = st.text_input("사업구분", value=str(row["사업구분"]),
                                   key="ed_business" + sfx)
        e_method = st.text_input("결제수단", value=str(row["결제수단"]),
                                 key="ed_method" + sfx)
        e_proof = st.selectbox("증빙여부", config.PROOF_VALUES,
                               index=_safe_index(config.PROOF_VALUES, row["증빙여부"]),
                               key="ed_proof" + sfx)
    e_note = st.text_input("비고", value=str(row["비고"]), key="ed_note" + sfx)

    b1, b2, _ = st.columns([1, 1, 4])
    with b1:
        if st.button("✏️ 수정 저장", type="primary"):
            updates = {
                "입출구분": e_inout, "대분류": e_major, "소분류": e_minor,
                "금액": int(e_amount), "거래일자": e_date, "거래처": e_vendor,
                "사업구분": e_business, "결제수단": e_method,
                "증빙여부": e_proof, "비고": e_note,
            }
            errors = transaction_service.update_transaction(wb, sel_id, updates)
            if errors:
                ui.show_errors(errors)
            else:
                backup = persist_and_reports(wb, int(float(row["회계연도"])))
                flash("success", f"거래 수정 완료 · {sel_id} · 백업: "
                                 f"{backup.name if backup else '-'}")
                st.rerun()
    with b2:
        if st.button("🗑️ 삭제(상태변경)"):
            transaction_service.delete_transaction(wb, sel_id)
            backup = persist_and_reports(wb, int(float(row["회계연도"])))
            flash("warning", f"거래를 '삭제' 상태로 변경했습니다 · {sel_id} · 백업: "
                             f"{backup.name if backup else '-'}")
            st.rerun()


def _safe_index(options, value):
    try:
        return options.index(value)
    except (ValueError, AttributeError):
        return 0


# ---------------------------------------------------------------------------
# 4. 수입결산 / 5. 지출결산
# ---------------------------------------------------------------------------

def _settlement_page(wb, kind: str):
    is_income = kind == "수입"
    icon = "💰" if is_income else "💸"
    st.header(f"{icon} {kind}결산")
    year = year_selector(wb, f"settle_{kind}")

    df = excel_service.read_ledger_df(wb)
    if is_income:
        by_major, by_minor, total = report_service.income_breakdown(df, year)
    else:
        by_major, by_minor, total = report_service.expense_breakdown(df, year)

    st.subheader("대분류별 합계")
    ui.show_table(by_major)
    st.subheader("소분류별 합계")
    ui.show_table(by_minor)
    st.metric(f"{kind}합계", ui.format_won(total))

    if st.button(f"📥 {kind}결산 시트 갱신"):
        if is_income:
            report_service.generate_income_report(wb, year)
        else:
            report_service.generate_expense_report(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"{kind}결산 시트를 갱신했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()


def page_income(wb):
    _settlement_page(wb, "수입")


def page_expense(wb):
    _settlement_page(wb, "지출")


# ---------------------------------------------------------------------------
# 6. 차입금현황
# ---------------------------------------------------------------------------

def page_borrowing(wb):
    st.header("🏦 차입금현황")
    year = year_selector(wb, "borrow_year")
    b = report_service.borrowing_status(wb, year)

    ui.show_key_value([
        ("기초 차입금", b["기초차입금"]),
        ("당기 차입금 유입", b["당기유입"]),
        ("당기 차입금 상환", b["당기상환"]),
        ("기말 차입금 잔액", b["기말잔액"]),
    ])
    st.info("차입금 유입·상환은 수입·지출 합계에 포함되지 않습니다.")

    if st.button("📥 차입금현황 시트 갱신"):
        report_service.generate_borrowing_report(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"차입금현황 시트를 갱신했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()


# ---------------------------------------------------------------------------
# 7. 간이대차대조표
# ---------------------------------------------------------------------------

def page_balance(wb):
    st.header("📑 간이대차대조표")
    year = year_selector(wb, "balance_year")
    b = report_service.balance_sheet(wb, year)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("자산")
        ui.show_key_value([
            ("현금 및 예금", b["현금및예금"]),
            ("미수금", b["미수금"]),
            ("선급금", b["선급금"]),
            ("자산합계", b["자산합계"]),
        ])
    with c2:
        st.subheader("부채 · 자본")
        ui.show_key_value([
            ("단기차입금", b["단기차입금"]),
            ("순자산", b["순자산"]),
            ("부채·자본 합계", b["부채자본합계"]),
        ])

    if abs(b["자산합계"] - b["부채자본합계"]) < 1:
        st.success(f"대차평형 일치 ✅  (자산합계 = 부채·자본합계 = "
                   f"{ui.format_won(b['자산합계'])})")
    else:
        st.error("대차평형 불일치 ❌")

    if st.button("📥 간이대차대조표 시트 갱신"):
        report_service.generate_balance_sheet(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"간이대차대조표 시트를 갱신했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()


# ---------------------------------------------------------------------------
# 8. 검증표
# ---------------------------------------------------------------------------

def page_validation(wb):
    st.header("✅ 검증표")
    year = year_selector(wb, "valid_year")
    rows, overall, _ = validation_service.build_validation(wb, year)

    st.metric("종합 검증상태", ui.status_badge(overall))

    table = pd.DataFrame(rows)[["번호", "검증항목", "상태", "상세"]]
    st.dataframe(table, width="stretch", hide_index=True)

    if st.button("📥 검증표 시트 갱신"):
        validation_service.generate_validation_report(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"검증표 시트를 갱신했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()


# ---------------------------------------------------------------------------
# 9. 총회보고서
# ---------------------------------------------------------------------------

def page_report(wb):
    st.header("📜 총회보고서")
    year = year_selector(wb, "report_year")
    rep = report_service.build_assembly_report(wb, year)
    s = rep["요약"]
    b = rep["대차대조표"]

    st.title(rep["제목"])
    st.caption(f"기관명: {rep['기관명']}")

    st.subheader(f"1. {year}년 수입내역")
    ui.show_table(rep["수입_대분류"])
    st.caption(f"수입합계: {ui.format_won(s['수입합계'])}")

    st.subheader(f"2. {year}년 지출내역")
    ui.show_table(rep["지출_대분류"])
    st.caption(f"지출합계: {ui.format_won(s['지출합계'])}")

    st.subheader("3. 당기 수지 결과")
    ui.show_key_value([
        ("수입합계", s["수입합계"]), ("지출합계", s["지출합계"]),
        ("당기수지차", s["당기수지차"]),
        ("기초현금", s["기초현금"]), ("기말현금", s["기말현금"]),
    ])

    st.subheader("4. 차입금 현황")
    ui.show_key_value([
        ("기초 차입금", s["기초차입금"]), ("당기 유입", s["차입유입"]),
        ("당기 상환", s["차입상환"]), ("기말 잔액", s["차입금잔액"]),
    ])

    st.subheader("5. 간이 대차대조표")
    c1, c2 = st.columns(2)
    with c1:
        ui.show_key_value([
            ("현금 및 예금", b["현금및예금"]), ("미수금", b["미수금"]),
            ("선급금", b["선급금"]), ("자산합계", b["자산합계"]),
        ])
    with c2:
        ui.show_key_value([
            ("단기차입금", b["단기차입금"]), ("순자산", b["순자산"]),
            ("부채·자본 합계", b["부채자본합계"]),
        ])

    st.subheader("6. 주요 검토사항")
    balanced = "일치" if abs(b["자산합계"] - b["부채자본합계"]) < 1 else "불일치"
    st.write(f"- 자산합계 = 부채·자본합계 : **{balanced}**")
    st.write("- 차입금 유입·상환은 수입·지출 합계에 미포함")

    st.subheader("7. 의결주문")
    st.info(rep["의결주문"])

    st.markdown("---")
    st.caption(f"※ 주석: {rep['주석']}")

    if st.button("📥 총회보고서 시트 갱신"):
        report_service.generate_assembly_report(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"총회보고서 시트를 갱신했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()


# ---------------------------------------------------------------------------
# 10. 설정
# ---------------------------------------------------------------------------

def page_settings(wb):
    st.header("⚙️ 설정")

    st.subheader("데이터 파일")
    st.code(str(config.EXCEL_PATH))
    exists = config.EXCEL_PATH.exists()
    st.write(f"파일 존재 여부: {'있음 ✅' if exists else '없음 ❌'}")

    st.warning("프로그램 실행 중에는 이 엑셀 파일을 열어두지 마세요. "
               "열어둔 상태로 저장하면 파일 충돌이 발생할 수 있습니다.")

    st.divider()
    st.subheader("전체 보고서 재생성")
    year = year_selector(wb, "settings_year")
    if st.button("🔄 모든 보고서/검증표 재생성 후 저장"):
        summary = report_service.generate_all_reports(wb, year)
        validation_service.generate_validation_report(wb, year)
        _saved, backup = excel_service.save_with_backup(wb)
        flash("success", f"{year}년 전체 보고서를 재생성했습니다 · 백업: "
                         f"{backup.name if backup else '-'}")
        st.rerun()

    st.divider()
    st.subheader("최근 백업 파일")
    backups = backup_service.list_backups()
    if backups:
        st.dataframe(pd.DataFrame({
            "파일명": [p.name for p in backups[:20]],
            "수정시각": [datetime.fromtimestamp(p.stat().st_mtime)
                     .strftime("%Y-%m-%d %H:%M:%S") for p in backups[:20]],
        }), width="stretch", hide_index=True)
    else:
        st.info("백업 파일이 아직 없습니다.")

    st.divider()
    st.subheader("계정과목")
    acc_df = pd.DataFrame(config.ACCOUNTS, columns=["입출구분", "대분류", "소분류"])
    st.dataframe(acc_df, width="stretch", hide_index=True)


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------

PAGES = {
    "대시보드": page_dashboard,
    "거래 입력": page_input,
    "거래 조회·수정": page_manage,
    "수입결산": page_income,
    "지출결산": page_expense,
    "차입금현황": page_borrowing,
    "간이대차대조표": page_balance,
    "검증표": page_validation,
    "총회보고서": page_report,
    "설정": page_settings,
}


def main():
    st.sidebar.title(config.ORG_NAME)
    st.sidebar.caption("결산 입력 프로그램 (MVP)")
    choice = st.sidebar.radio("메뉴", list(PAGES.keys()))
    st.sidebar.divider()
    st.sidebar.caption("⚠️ 실행 중에는 데이터 엑셀 파일을 열지 마세요.")

    wb = load_wb()
    show_flash()
    PAGES[choice](wb)


if __name__ == "__main__":
    main()
