"""
report_service.py
------------------
결산 계산 및 보고서 생성.

핵심 회계 기준:
  - 수입합계  : 입출구분 == '수입' 인 거래만
  - 지출합계  : 입출구분 == '지출' 인 거래만
  - 차입유입/차입상환 : 수입·지출 합계에 포함하지 않고 차입금현황으로만 관리
  - 자산증가/자산감소 : 자산관리·간이대차대조표에만 반영 (현금흐름과 분리)
  - 자금조정 : 현금잔액 검증에만 반영 (수입·지출 합계 제외)

생성 보고서:
  - 수입결산 / 지출결산 / 차입금현황 / 간이대차대조표 / 총회보고서
"""

from datetime import datetime

import pandas as pd

from . import config, excel_service, transaction_service


# ---------------------------------------------------------------------------
# 기초잔액
# ---------------------------------------------------------------------------

def get_opening_balances(wb, year: int) -> dict:
    """
    기초잔액 시트에서 해당 연도의 기초 값을 읽는다.
    시트 형식: [회계연도, 항목, 금액]
    누락 항목은 0 으로 채운다.
    """
    df = excel_service.read_sheet_df(wb, config.SHEET_OPENING)
    result = {item: 0.0 for item in config.OPENING_ITEMS}

    if df is None or df.empty or "항목" not in df.columns:
        return result

    # 금액 숫자화
    if "금액" in df.columns:
        df = df.copy()
        df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0)

    if "회계연도" in df.columns:
        df["회계연도"] = pd.to_numeric(df["회계연도"], errors="coerce").fillna(0).astype(int)
        df = df[df["회계연도"] == int(year)]

    for _, row in df.iterrows():
        item = str(row.get("항목", "")).strip()
        if item in result:
            result[item] = float(row.get("금액", 0))
    return result


# ---------------------------------------------------------------------------
# 종합 재무 요약
# ---------------------------------------------------------------------------

def _sum_where(df: pd.DataFrame, inout: str, major: str = None) -> float:
    if df is None or df.empty:
        return 0.0
    m = df["입출구분"] == inout
    if major is not None:
        m = m & (df["대분류"] == major)
    return float(df.loc[m, "금액"].sum())


def compute_summary(df: pd.DataFrame, year: int, opening: dict) -> dict:
    """
    회계연도 기준 종합 재무 요약을 계산한다.
    df 는 전체 거래(대장) DataFrame. 내부에서 활성(삭제 제외) + 연도 필터링.
    """
    active = transaction_service.active_transactions(df, year)

    income_total = _sum_where(active, config.INOUT_INCOME)
    expense_total = _sum_where(active, config.INOUT_EXPENSE)
    borrow_in = _sum_where(active, config.INOUT_BORROW_IN)
    borrow_repay = _sum_where(active, config.INOUT_BORROW_REPAY)
    fund_adjust = _sum_where(active, config.INOUT_FUND_ADJUST)

    asset_up_recv = _sum_where(active, config.INOUT_ASSET_UP, config.ASSET_TYPE_RECEIVABLE)
    asset_dn_recv = _sum_where(active, config.INOUT_ASSET_DOWN, config.ASSET_TYPE_RECEIVABLE)
    asset_up_prep = _sum_where(active, config.INOUT_ASSET_UP, config.ASSET_TYPE_PREPAID)
    asset_dn_prep = _sum_where(active, config.INOUT_ASSET_DOWN, config.ASSET_TYPE_PREPAID)

    opening_cash = float(opening.get(config.OPENING_CASH, 0))
    opening_borrow = float(opening.get(config.OPENING_BORROWING, 0))
    opening_recv = float(opening.get(config.OPENING_RECEIVABLE, 0))
    opening_prep = float(opening.get(config.OPENING_PREPAID, 0))

    net_surplus = income_total - expense_total
    closing_cash = (opening_cash + income_total - expense_total
                    + borrow_in - borrow_repay + fund_adjust)
    closing_borrow = opening_borrow + borrow_in - borrow_repay
    receivables = opening_recv + asset_up_recv - asset_dn_recv
    prepaid = opening_prep + asset_up_prep - asset_dn_prep

    total_assets = closing_cash + receivables + prepaid
    total_liabilities = closing_borrow
    net_assets = total_assets - total_liabilities

    return {
        "회계연도": int(year),
        "수입합계": income_total,
        "지출합계": expense_total,
        "당기수지차": net_surplus,
        "차입유입": borrow_in,
        "차입상환": borrow_repay,
        "자금조정": fund_adjust,
        "기초현금": opening_cash,
        "기말현금": closing_cash,
        "기초차입금": opening_borrow,
        "차입금잔액": closing_borrow,
        "미수금": receivables,
        "선급금": prepaid,
        "자산합계": total_assets,
        "부채자본합계": total_assets,   # 대차평형 (자산합계 = 부채+순자산)
        "부채합계": total_liabilities,
        "순자산": net_assets,
    }


def summary_from_workbook(wb, year: int) -> dict:
    """워크북에서 대장·기초잔액을 읽어 요약을 계산한다."""
    df = excel_service.read_ledger_df(wb)
    opening = get_opening_balances(wb, year)
    return compute_summary(df, year, opening)


# ---------------------------------------------------------------------------
# 수입 / 지출 결산 상세
# ---------------------------------------------------------------------------

def _breakdown(df: pd.DataFrame, year: int, inout: str):
    """대분류별 / 소분류별 합계 DataFrame 과 총합을 반환한다."""
    active = transaction_service.active_transactions(df, year)
    subset = active[active["입출구분"] == inout] if not active.empty else active

    if subset is None or subset.empty:
        by_major = pd.DataFrame(columns=["대분류", "금액"])
        by_minor = pd.DataFrame(columns=["대분류", "소분류", "금액"])
        return by_major, by_minor, 0.0

    by_major = (subset.groupby("대분류", as_index=False)["금액"].sum()
                .sort_values("금액", ascending=False).reset_index(drop=True))
    by_minor = (subset.groupby(["대분류", "소분류"], as_index=False)["금액"].sum()
                .sort_values(["대분류", "금액"], ascending=[True, False])
                .reset_index(drop=True))
    total = float(subset["금액"].sum())
    return by_major, by_minor, total


def income_breakdown(df: pd.DataFrame, year: int):
    return _breakdown(df, year, config.INOUT_INCOME)


def expense_breakdown(df: pd.DataFrame, year: int):
    return _breakdown(df, year, config.INOUT_EXPENSE)


def _settlement_grid(title: str, by_major: pd.DataFrame,
                     by_minor: pd.DataFrame, total: float) -> list[list]:
    """수입/지출 결산 시트용 그리드를 만든다."""
    grid = [[title], [], ["■ 대분류별 합계"], ["대분류", "금액"]]
    for _, r in by_major.iterrows():
        grid.append([r["대분류"], float(r["금액"])])
    grid += [[], ["■ 소분류별 합계"], ["대분류", "소분류", "금액"]]
    for _, r in by_minor.iterrows():
        grid.append([r["대분류"], r["소분류"], float(r["금액"])])
    grid += [[], ["합계", total]]
    return grid


def generate_income_report(wb, year: int):
    df = excel_service.read_ledger_df(wb)
    by_major, by_minor, total = income_breakdown(df, year)
    grid = _settlement_grid(f"{year}년 수입결산", by_major, by_minor, total)
    excel_service.write_grid_to_sheet(wb, config.SHEET_INCOME, grid)
    return by_major, by_minor, total


def generate_expense_report(wb, year: int):
    df = excel_service.read_ledger_df(wb)
    by_major, by_minor, total = expense_breakdown(df, year)
    grid = _settlement_grid(f"{year}년 지출결산", by_major, by_minor, total)
    excel_service.write_grid_to_sheet(wb, config.SHEET_EXPENSE, grid)
    return by_major, by_minor, total


# ---------------------------------------------------------------------------
# 차입금현황
# ---------------------------------------------------------------------------

def borrowing_status(wb, year: int) -> dict:
    s = summary_from_workbook(wb, year)
    return {
        "기초차입금": s["기초차입금"],
        "당기유입": s["차입유입"],
        "당기상환": s["차입상환"],
        "기말잔액": s["차입금잔액"],
    }


def generate_borrowing_report(wb, year: int) -> dict:
    b = borrowing_status(wb, year)
    grid = [
        [f"{year}년 차입금현황"],
        [],
        ["항목", "금액"],
        ["기초 차입금", b["기초차입금"]],
        ["당기 차입금 유입", b["당기유입"]],
        ["당기 차입금 상환", b["당기상환"]],
        ["기말 차입금 잔액", b["기말잔액"]],
    ]
    excel_service.write_grid_to_sheet(wb, config.SHEET_BORROWING, grid)
    return b


# ---------------------------------------------------------------------------
# 간이대차대조표
# ---------------------------------------------------------------------------

def balance_sheet(wb, year: int) -> dict:
    s = summary_from_workbook(wb, year)
    return {
        "현금및예금": s["기말현금"],
        "미수금": s["미수금"],
        "선급금": s["선급금"],
        "자산합계": s["자산합계"],
        "단기차입금": s["부채합계"],
        "순자산": s["순자산"],
        "부채자본합계": s["부채자본합계"],
    }


def generate_balance_sheet(wb, year: int) -> dict:
    b = balance_sheet(wb, year)
    grid = [
        [f"{year}년 간이대차대조표"],
        [],
        ["[자산]", "금액", "", "[부채·자본]", "금액"],
        ["현금 및 예금", b["현금및예금"], "", "단기차입금", b["단기차입금"]],
        ["미수금", b["미수금"], "", "순자산", b["순자산"]],
        ["선급금", b["선급금"], "", "", ""],
        ["자산합계", b["자산합계"], "", "부채·자본 합계", b["부채자본합계"]],
    ]
    excel_service.write_grid_to_sheet(wb, config.SHEET_BALANCE, grid)
    return b


# ---------------------------------------------------------------------------
# 총회보고서
# ---------------------------------------------------------------------------

RESOLUTION_TEXT = (
    "2025년도 사업보고 및 결산 내용을 원안대로 승인하고, "
    "이를 제13차 정기총회 안건으로 상정할 것을 의결함."
)

FOOTNOTE_TEXT = (
    "본 수지결산은 협의회의 운영수입과 운영지출을 기준으로 작성하였으며, "
    "차입금 유입 및 상환은 수입·지출 합계에 포함하지 않고 별도 차입금 현황으로 관리한다."
)


def build_assembly_report(wb, year: int) -> dict:
    """총회보고서 내용을 구조화된 dict 로 반환한다(미리보기·시트 공용)."""
    df = excel_service.read_ledger_df(wb)
    summary = summary_from_workbook(wb, year)
    inc_major, _, inc_total = income_breakdown(df, year)
    exp_major, _, exp_total = expense_breakdown(df, year)
    b = balance_sheet(wb, year)

    return {
        "제목": f"{year}년 결산보고",
        "기관명": config.ORG_NAME,
        "요약": summary,
        "수입_대분류": inc_major,
        "지출_대분류": exp_major,
        "대차대조표": b,
        "의결주문": RESOLUTION_TEXT,
        "주석": FOOTNOTE_TEXT,
    }


def generate_assembly_report(wb, year: int) -> dict:
    rep = build_assembly_report(wb, year)
    s = rep["요약"]
    b = rep["대차대조표"]

    grid: list[list] = []
    grid.append([rep["제목"]])
    grid.append([f"기관명: {rep['기관명']}"])
    grid.append([])

    # 1. 수입내역
    grid.append([f"1. {year}년 수입내역"])
    grid.append(["대분류", "금액"])
    for _, r in rep["수입_대분류"].iterrows():
        grid.append([r["대분류"], float(r["금액"])])
    grid.append(["수입합계", s["수입합계"]])
    grid.append([])

    # 2. 지출내역
    grid.append([f"2. {year}년 지출내역"])
    grid.append(["대분류", "금액"])
    for _, r in rep["지출_대분류"].iterrows():
        grid.append([r["대분류"], float(r["금액"])])
    grid.append(["지출합계", s["지출합계"]])
    grid.append([])

    # 3. 당기 수지 결과
    grid.append(["3. 당기 수지 결과"])
    grid.append(["수입합계", s["수입합계"]])
    grid.append(["지출합계", s["지출합계"]])
    grid.append(["당기수지차", s["당기수지차"]])
    grid.append(["기초현금", s["기초현금"]])
    grid.append(["기말현금", s["기말현금"]])
    grid.append([])

    # 4. 차입금 현황
    grid.append(["4. 차입금 현황"])
    grid.append(["기초 차입금", s["기초차입금"]])
    grid.append(["당기 유입", s["차입유입"]])
    grid.append(["당기 상환", s["차입상환"]])
    grid.append(["기말 잔액", s["차입금잔액"]])
    grid.append([])

    # 5. 간이 대차대조표
    grid.append(["5. 간이 대차대조표"])
    grid.append(["[자산]", "금액", "[부채·자본]", "금액"])
    grid.append(["현금 및 예금", b["현금및예금"], "단기차입금", b["단기차입금"]])
    grid.append(["미수금", b["미수금"], "순자산", b["순자산"]])
    grid.append(["선급금", b["선급금"], "", ""])
    grid.append(["자산합계", b["자산합계"], "부채·자본 합계", b["부채자본합계"]])
    grid.append([])

    # 6. 주요 검토사항
    grid.append(["6. 주요 검토사항"])
    balanced = "일치" if abs(b["자산합계"] - b["부채자본합계"]) < 1 else "불일치"
    grid.append(["자산합계 = 부채·자본합계", balanced])
    grid.append(["차입금은 수입·지출에 미포함", "확인"])
    grid.append([])

    # 7. 의결주문
    grid.append(["7. 의결주문"])
    grid.append([rep["의결주문"]])
    grid.append([])

    # 주석
    grid.append(["※ 주석"])
    grid.append([rep["주석"]])

    excel_service.write_grid_to_sheet(wb, config.SHEET_REPORT, grid)
    return rep


# ---------------------------------------------------------------------------
# 자산관리 시트 (자산 거래 내역 정리)
# ---------------------------------------------------------------------------

def generate_asset_report(wb, year: int):
    df = excel_service.read_ledger_df(wb)
    active = transaction_service.active_transactions(df, year)
    asset_io = [config.INOUT_ASSET_UP, config.INOUT_ASSET_DOWN]
    subset = active[active["입출구분"].isin(asset_io)] if not active.empty else active
    cols = ["거래일자", "입출구분", "대분류", "소분류", "금액", "거래처", "비고"]
    if subset is None or subset.empty:
        out = pd.DataFrame(columns=cols)
    else:
        out = subset[cols].reset_index(drop=True)
    excel_service.write_df_to_sheet(wb, config.SHEET_ASSETS, out)
    return out


# ---------------------------------------------------------------------------
# 전체 보고서 일괄 생성
# ---------------------------------------------------------------------------

def generate_all_reports(wb, year: int) -> dict:
    """모든 결산 보고서를 재생성한다(파일 저장은 호출자 담당)."""
    generate_income_report(wb, year)
    generate_expense_report(wb, year)
    generate_borrowing_report(wb, year)
    generate_asset_report(wb, year)
    generate_balance_sheet(wb, year)
    generate_assembly_report(wb, year)
    return summary_from_workbook(wb, year)
