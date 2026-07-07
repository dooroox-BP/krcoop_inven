"""
ui_components.py
----------------
Streamlit 공통 UI 헬퍼.

- 금액 포맷 함수 (천 단위 콤마 + '원')
- 대시보드 카드 표시
- 공통 표 표시
- 메시지(성공/경고/오류) 표시
"""

import pandas as pd

try:
    import streamlit as st
except ModuleNotFoundError:  # 테스트 환경 등 Streamlit 미설치 시
    st = None


# ---------------------------------------------------------------------------
# 금액 포맷
# ---------------------------------------------------------------------------

def format_won(value) -> str:
    """숫자를 '1,234,567원' 형태로 포맷한다."""
    try:
        num = float(value)
    except (ValueError, TypeError):
        return str(value)
    return f"{num:,.0f}원"


def format_number(value) -> str:
    """숫자를 천 단위 콤마로 포맷한다('원' 없음)."""
    try:
        num = float(value)
    except (ValueError, TypeError):
        return str(value)
    return f"{num:,.0f}"


# ---------------------------------------------------------------------------
# 메시지
# ---------------------------------------------------------------------------

def show_success(msg: str):
    if st:
        st.success(msg)


def show_warning(msg: str):
    if st:
        st.warning(msg)


def show_error(msg: str):
    if st:
        st.error(msg)


def show_info(msg: str):
    if st:
        st.info(msg)


def show_errors(errors: list[str]):
    """검증 오류 목록을 표시한다."""
    if st and errors:
        st.error("입력값을 확인해 주세요:\n\n- " + "\n- ".join(errors))


# ---------------------------------------------------------------------------
# 대시보드 카드
# ---------------------------------------------------------------------------

def metric_card(label: str, value, is_money: bool = True):
    """단일 지표 카드."""
    if not st:
        return
    display = format_won(value) if is_money else str(value)
    st.metric(label, display)


def dashboard_cards(summary: dict, overall_status: str):
    """대시보드 상단 카드 묶음."""
    if not st:
        return

    st.subheader(f"{summary['회계연도']}년 결산 요약")

    row1 = st.columns(4)
    with row1[0]:
        st.metric("회계연도", f"{summary['회계연도']}년")
    with row1[1]:
        st.metric("총수입", format_won(summary["수입합계"]))
    with row1[2]:
        st.metric("총지출", format_won(summary["지출합계"]))
    with row1[3]:
        st.metric("당기수지차", format_won(summary["당기수지차"]))

    row2 = st.columns(4)
    with row2[0]:
        st.metric("기말현금", format_won(summary["기말현금"]))
    with row2[1]:
        st.metric("차입금잔액", format_won(summary["차입금잔액"]))
    with row2[2]:
        st.metric("미수금", format_won(summary["미수금"]))
    with row2[3]:
        st.metric("선급금", format_won(summary["선급금"]))

    row3 = st.columns(4)
    with row3[0]:
        st.metric("자산합계", format_won(summary["자산합계"]))
    with row3[1]:
        st.metric("부채·자본합계", format_won(summary["부채자본합계"]))
    with row3[2]:
        st.metric("순자산", format_won(summary["순자산"]))
    with row3[3]:
        emoji = {"정상": "✅", "확인필요": "⚠️", "오류": "❌"}.get(overall_status, "")
        st.metric("검증상태", f"{emoji} {overall_status}")


# ---------------------------------------------------------------------------
# 표 표시
# ---------------------------------------------------------------------------

def money_dataframe(df: pd.DataFrame, money_cols=("금액",)):
    """금액 컬럼을 포맷한 사본 DataFrame 을 반환한다(표시용)."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in money_cols:
        if col in out.columns:
            out[col] = out[col].apply(format_won)
    return out


def show_table(df: pd.DataFrame, money_cols=("금액",)):
    """금액 포맷을 적용하여 표를 표시한다."""
    if not st:
        return
    if df is None or df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    st.dataframe(money_dataframe(df, money_cols), width="stretch", hide_index=True)


def show_key_value(pairs: list[tuple], money: bool = True):
    """(항목, 금액) 목록을 2열 표로 표시한다."""
    if not st:
        return
    data = []
    for label, value in pairs:
        data.append({"항목": label,
                     "금액": format_won(value) if money else value})
    st.dataframe(pd.DataFrame(data), width="stretch", hide_index=True)


def status_badge(status: str) -> str:
    """검증 상태를 이모지 배지 문자열로."""
    return {"정상": "✅ 정상", "확인필요": "⚠️ 확인필요",
            "오류": "❌ 오류"}.get(status, status)
