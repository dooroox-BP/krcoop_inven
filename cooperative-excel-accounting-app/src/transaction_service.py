"""
transaction_service.py
-----------------------
거래 데이터(거래입력대장)에 대한 업무 로직.

- 거래ID 생성 (TX-YYYYMMDD-0001)
- 입력값 검증
- 거래 추가 / 수정 / soft delete
- 거래 필터링 / 활성 거래 조회

DataFrame 기반 순수 함수 + Workbook 연동 래퍼로 구성한다.
"""

from datetime import date, datetime

import pandas as pd

from . import config, excel_service


# ---------------------------------------------------------------------------
# 날짜 / ID 유틸
# ---------------------------------------------------------------------------

def _to_date_str(value) -> str:
    """다양한 형태의 날짜 입력을 'YYYY-MM-DD' 문자열로 정규화한다."""
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    # 'YYYY-MM-DD HH:MM:SS' 형태면 앞 10자만
    return s[:10]


def _date_compact(date_str: str) -> str:
    """'YYYY-MM-DD' → 'YYYYMMDD'. 실패 시 오늘 날짜."""
    s = _to_date_str(date_str).replace("-", "")
    if len(s) >= 8 and s[:8].isdigit():
        return s[:8]
    return datetime.now().strftime("%Y%m%d")


def generate_transaction_id(df: pd.DataFrame, tx_date) -> str:
    """
    거래일자 기준으로 다음 거래ID를 생성한다.
    형식: TX-YYYYMMDD-0001
    같은 날짜의 기존 ID 중 최대 일련번호 + 1.
    """
    compact = _date_compact(tx_date)
    prefix = f"{config.TX_ID_PREFIX}-{compact}-"

    max_seq = 0
    if df is not None and not df.empty and "거래ID" in df.columns:
        for tid in df["거래ID"].astype(str):
            if tid.startswith(prefix):
                tail = tid[len(prefix):]
                if tail.isdigit():
                    max_seq = max(max_seq, int(tail))
    return f"{prefix}{max_seq + 1:04d}"


# ---------------------------------------------------------------------------
# 입력값 검증
# ---------------------------------------------------------------------------

REQUIRED_FIELDS = ["회계연도", "거래일자", "입출구분", "대분류", "소분류", "금액"]


def validate_transaction(data: dict) -> list[str]:
    """
    거래 입력값을 검증하고 오류 메시지 목록을 반환한다.
    빈 목록이면 유효한 입력.
    """
    errors: list[str] = []

    # 필수값
    for field in REQUIRED_FIELDS:
        value = data.get(field, "")
        if value is None or (isinstance(value, str) and value.strip() == ""):
            errors.append(f"필수 항목 누락: {field}")

    # 회계연도
    year = data.get("회계연도")
    try:
        if year not in (None, "") and int(year) <= 0:
            errors.append("회계연도가 올바르지 않습니다.")
    except (ValueError, TypeError):
        errors.append("회계연도는 숫자여야 합니다.")

    # 금액 > 0
    amount = data.get("금액")
    try:
        if amount in (None, ""):
            pass  # 위 필수값에서 이미 처리
        elif float(amount) <= 0:
            errors.append("금액은 0보다 커야 합니다.")
    except (ValueError, TypeError):
        errors.append("금액은 숫자여야 합니다.")

    # 입출구분 허용값
    inout = data.get("입출구분", "")
    if inout and inout not in config.INOUT_TYPES:
        errors.append(f"허용되지 않는 입출구분: {inout}")

    # 계정과목 존재 여부 (입출구분·대분류·소분류 조합)
    if inout and data.get("대분류") and data.get("소분류"):
        combo = (inout, str(data.get("대분류")), str(data.get("소분류")))
        if combo not in config.ACCOUNTS:
            errors.append(
                f"등록되지 않은 계정과목: {inout} / {data.get('대분류')} / {data.get('소분류')}"
            )

    return errors


# ---------------------------------------------------------------------------
# 행 구성
# ---------------------------------------------------------------------------

def build_row(data: dict, tx_id: str, now: datetime = None) -> dict:
    """검증된 입력값으로 거래입력대장 한 행(dict)을 만든다."""
    now = now or datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    return {
        "거래ID": tx_id,
        "회계연도": int(data.get("회계연도")),
        "거래일자": _to_date_str(data.get("거래일자")),
        "입출구분": data.get("입출구분", ""),
        "대분류": data.get("대분류", ""),
        "소분류": data.get("소분류", ""),
        "금액": float(data.get("금액", 0)),
        "거래처": data.get("거래처", ""),
        "사업구분": data.get("사업구분", ""),
        "결제수단": data.get("결제수단", ""),
        "증빙여부": data.get("증빙여부", ""),
        "증빙링크": data.get("증빙링크", ""),
        "비고": data.get("비고", ""),
        "입력자": data.get("입력자", ""),
        "입력일시": now_str,
        "수정일시": "",
        "상태": config.STATUS_NORMAL,
    }


# ---------------------------------------------------------------------------
# 순수 DataFrame 연산
# ---------------------------------------------------------------------------

def append_row(df: pd.DataFrame, row: dict) -> pd.DataFrame:
    """행 dict 를 DataFrame 에 추가한다."""
    new = pd.DataFrame([row], columns=config.LEDGER_COLUMNS)
    if df is None or df.empty:
        return new
    return pd.concat([df, new], ignore_index=True)


def apply_update(df: pd.DataFrame, tx_id: str, updates: dict,
                 now: datetime = None) -> pd.DataFrame:
    """
    거래ID 에 해당하는 행을 수정한다.
    수정일시를 기록하고 상태를 '수정'으로 바꾼다(단, 삭제 상태는 유지).
    """
    now = now or datetime.now()
    df = df.copy()
    mask = df["거래ID"].astype(str) == str(tx_id)
    if not mask.any():
        return df

    for key, value in updates.items():
        if key in df.columns:
            if key == "금액":
                value = float(value) if value not in (None, "") else 0
            elif key == "회계연도":
                value = int(value) if value not in (None, "") else 0
            df.loc[mask, key] = value

    df.loc[mask, "수정일시"] = now.strftime("%Y-%m-%d %H:%M:%S")
    # 삭제 상태가 아니면 '수정'으로 표시
    not_deleted = mask & (df["상태"] != config.STATUS_DELETED)
    df.loc[not_deleted, "상태"] = config.STATUS_MODIFIED
    return df


def apply_soft_delete(df: pd.DataFrame, tx_id: str,
                      now: datetime = None) -> pd.DataFrame:
    """상태를 '삭제'로 변경한다(실제 행 삭제 없음)."""
    now = now or datetime.now()
    df = df.copy()
    mask = df["거래ID"].astype(str) == str(tx_id)
    df.loc[mask, "상태"] = config.STATUS_DELETED
    df.loc[mask, "수정일시"] = now.strftime("%Y-%m-%d %H:%M:%S")
    return df


# ---------------------------------------------------------------------------
# 필터링 / 조회
# ---------------------------------------------------------------------------

def active_transactions(df: pd.DataFrame, year: int = None) -> pd.DataFrame:
    """
    집계에 사용할 활성 거래(상태 != 삭제)를 반환한다.
    year 를 주면 해당 회계연도만.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame(columns=config.LEDGER_COLUMNS)
    out = df[df["상태"] != config.STATUS_DELETED].copy()
    if year is not None:
        out = out[out["회계연도"] == int(year)]
    return out


def filter_transactions(df: pd.DataFrame, *, year=None, inout=None, major=None,
                        business=None, vendor=None, start_date=None,
                        end_date=None, status=None,
                        include_deleted=True) -> pd.DataFrame:
    """조회 화면용 다중 필터."""
    if df is None or df.empty:
        return pd.DataFrame(columns=config.LEDGER_COLUMNS)

    out = df.copy()

    if not include_deleted:
        out = out[out["상태"] != config.STATUS_DELETED]
    if status:
        out = out[out["상태"] == status]
    if year:
        out = out[out["회계연도"] == int(year)]
    if inout:
        out = out[out["입출구분"] == inout]
    if major:
        out = out[out["대분류"] == major]
    if business:
        out = out[out["사업구분"] == business]
    if vendor:
        out = out[out["거래처"].astype(str).str.contains(str(vendor), na=False)]
    if start_date:
        out = out[out["거래일자"].astype(str) >= _to_date_str(start_date)]
    if end_date:
        out = out[out["거래일자"].astype(str) <= _to_date_str(end_date)]

    return out


# ---------------------------------------------------------------------------
# Workbook 연동 (add / update / delete)
# ---------------------------------------------------------------------------

def add_transaction(wb, data: dict, now: datetime = None):
    """
    검증 → ID 생성 → 대장에 추가 → 시트에 반영(파일 저장은 호출자 담당).
    반환: (tx_id 또는 None, errors 리스트)
    """
    errors = validate_transaction(data)
    if errors:
        return None, errors

    df = excel_service.read_ledger_df(wb)
    tx_id = generate_transaction_id(df, data.get("거래일자"))
    row = build_row(data, tx_id, now)
    df = append_row(df, row)
    excel_service.write_df_to_sheet(wb, config.SHEET_LEDGER, df)
    return tx_id, []


def update_transaction(wb, tx_id: str, updates: dict, now: datetime = None):
    """
    거래 수정 → 시트에 반영. 반환: errors 리스트(빈 리스트면 성공).
    """
    df = excel_service.read_ledger_df(wb)

    # 수정 후 값으로 검증 (변경 대상 행 기준)
    mask = df["거래ID"].astype(str) == str(tx_id)
    if not mask.any():
        return [f"거래ID 를 찾을 수 없습니다: {tx_id}"]

    merged = df[mask].iloc[0].to_dict()
    merged.update(updates)
    errors = validate_transaction(merged)
    if errors:
        return errors

    df = apply_update(df, tx_id, updates, now)
    excel_service.write_df_to_sheet(wb, config.SHEET_LEDGER, df)
    return []


def delete_transaction(wb, tx_id: str, now: datetime = None) -> bool:
    """거래 soft delete → 시트에 반영. 반환: 성공 여부."""
    df = excel_service.read_ledger_df(wb)
    mask = df["거래ID"].astype(str) == str(tx_id)
    if not mask.any():
        return False
    df = apply_soft_delete(df, tx_id, now)
    excel_service.write_df_to_sheet(wb, config.SHEET_LEDGER, df)
    return True
