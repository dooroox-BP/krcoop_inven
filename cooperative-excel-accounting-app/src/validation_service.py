"""
validation_service.py
----------------------
회계 검증 및 검증표 생성.

검증 항목:
  1. 수입합계 - 지출합계 = 당기수지차
  2. 기초현금 + 수입 - 지출 + 차입유입 - 차입상환 + 자금조정 = 기말현금
  3. 자산합계 = 부채·자본합계
  4. 필수값 누락 거래
  5. 금액 오류 거래
  6. 증빙 없는 거래
  7. 계정과목 누락(미등록) 거래
  8. 상태가 '삭제'인 거래 집계 제외 여부

각 항목 상태 : 정상 / 확인필요 / 오류
"""

from . import config, excel_service, report_service, transaction_service

# 상태 상수
OK = "정상"
WARN = "확인필요"
ERROR = "오류"

# 금액 비교 허용 오차(원)
TOLERANCE = 1.0


def _worst(statuses: list[str]) -> str:
    """상태 목록 중 가장 나쁜 값을 반환한다(오류 > 확인필요 > 정상)."""
    if ERROR in statuses:
        return ERROR
    if WARN in statuses:
        return WARN
    return OK


# ---------------------------------------------------------------------------
# 데이터 품질 검사 (행 단위)
# ---------------------------------------------------------------------------

def _missing_required(active) -> list[str]:
    """필수값 누락 거래ID 목록."""
    bad = []
    for _, row in active.iterrows():
        for field in transaction_service.REQUIRED_FIELDS:
            val = row.get(field, "")
            if val is None or str(val).strip() in ("", "0") and field == "금액":
                bad.append(str(row.get("거래ID", "")))
                break
            if val is None or str(val).strip() == "":
                bad.append(str(row.get("거래ID", "")))
                break
    return bad


def _amount_errors(active) -> list[str]:
    """금액이 0 이하인 거래ID 목록."""
    bad = []
    for _, row in active.iterrows():
        try:
            if float(row.get("금액", 0)) <= 0:
                bad.append(str(row.get("거래ID", "")))
        except (ValueError, TypeError):
            bad.append(str(row.get("거래ID", "")))
    return bad


def _proof_missing(active) -> list[str]:
    """증빙이 없는 거래ID 목록."""
    bad = []
    for _, row in active.iterrows():
        proof = str(row.get("증빙여부", "")).strip()
        if proof != config.PROOF_YES:
            bad.append(str(row.get("거래ID", "")))
    return bad


def _account_missing(active) -> list[str]:
    """미등록 계정과목 거래ID 목록."""
    bad = []
    for _, row in active.iterrows():
        combo = (str(row.get("입출구분", "")),
                 str(row.get("대분류", "")),
                 str(row.get("소분류", "")))
        if combo not in config.ACCOUNTS:
            bad.append(str(row.get("거래ID", "")))
    return bad


# ---------------------------------------------------------------------------
# 검증표 구성
# ---------------------------------------------------------------------------

def build_validation(wb, year: int):
    """
    검증 항목 리스트와 종합 상태를 반환한다.
    반환: (rows: list[dict], overall: str, summary: dict)
    """
    df = excel_service.read_ledger_df(wb)
    summary = report_service.summary_from_workbook(wb, year)
    active = transaction_service.active_transactions(df, year)

    rows = []

    # 1. 수입 - 지출 = 당기수지차
    diff = summary["수입합계"] - summary["지출합계"]
    st1 = OK if abs(diff - summary["당기수지차"]) < TOLERANCE else ERROR
    rows.append({
        "번호": 1,
        "검증항목": "수입합계 - 지출합계 = 당기수지차",
        "상태": st1,
        "상세": f"{summary['수입합계']:,.0f} - {summary['지출합계']:,.0f} "
                f"= {summary['당기수지차']:,.0f}",
    })

    # 2. 현금 흐름 등식
    expected_cash = (summary["기초현금"] + summary["수입합계"] - summary["지출합계"]
                     + summary["차입유입"] - summary["차입상환"] + summary["자금조정"])
    st2 = OK if abs(expected_cash - summary["기말현금"]) < TOLERANCE else ERROR
    rows.append({
        "번호": 2,
        "검증항목": "기초현금 + 수입 - 지출 + 차입유입 - 차입상환 + 자금조정 = 기말현금",
        "상태": st2,
        "상세": f"계산값 {expected_cash:,.0f} = 기말현금 {summary['기말현금']:,.0f}",
    })

    # 3. 자산합계 = 부채·자본합계
    st3 = OK if abs(summary["자산합계"] - summary["부채자본합계"]) < TOLERANCE else ERROR
    rows.append({
        "번호": 3,
        "검증항목": "자산합계 = 부채·자본합계",
        "상태": st3,
        "상세": f"자산 {summary['자산합계']:,.0f} = 부채·자본 {summary['부채자본합계']:,.0f}",
    })

    # 4. 필수값 누락
    miss = _missing_required(active)
    st4 = ERROR if miss else OK
    rows.append({
        "번호": 4,
        "검증항목": "필수값 누락 거래",
        "상태": st4,
        "상세": f"{len(miss)}건" + (f" ({', '.join(miss[:5])}…)" if miss else " 없음"),
    })

    # 5. 금액 오류
    amt = _amount_errors(active)
    st5 = ERROR if amt else OK
    rows.append({
        "번호": 5,
        "검증항목": "금액 오류 거래 (0 이하)",
        "상태": st5,
        "상세": f"{len(amt)}건" + (f" ({', '.join(amt[:5])}…)" if amt else " 없음"),
    })

    # 6. 증빙 없음
    proof = _proof_missing(active)
    st6 = WARN if proof else OK
    rows.append({
        "번호": 6,
        "검증항목": "증빙 없는 거래",
        "상태": st6,
        "상세": f"{len(proof)}건" + (f" ({', '.join(proof[:5])}…)" if proof else " 없음"),
    })

    # 7. 계정과목 미등록
    acc = _account_missing(active)
    st7 = ERROR if acc else OK
    rows.append({
        "번호": 7,
        "검증항목": "계정과목 누락(미등록) 거래",
        "상태": st7,
        "상세": f"{len(acc)}건" + (f" ({', '.join(acc[:5])}…)" if acc else " 없음"),
    })

    # 8. 삭제 거래 집계 제외 여부
    deleted_cnt = int((df["상태"] == config.STATUS_DELETED).sum()) if not df.empty else 0
    rows.append({
        "번호": 8,
        "검증항목": "상태 '삭제' 거래 집계 제외",
        "상태": OK,
        "상세": f"삭제 {deleted_cnt}건 집계 제외 (활성 {len(active)}건 집계)",
    })

    overall = _worst([r["상태"] for r in rows])
    return rows, overall, summary


def generate_validation_report(wb, year: int):
    """검증표를 계산하여 검증표 시트에 기록한다."""
    rows, overall, summary = build_validation(wb, year)

    grid = [
        [f"{year}년 결산 검증표"],
        [f"종합 검증상태: {overall}"],
        [],
        ["번호", "검증항목", "상태", "상세"],
    ]
    for r in rows:
        grid.append([r["번호"], r["검증항목"], r["상태"], r["상세"]])

    excel_service.write_grid_to_sheet(wb, config.SHEET_VALIDATION, grid)
    return rows, overall, summary
