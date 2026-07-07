"""
test_accounting.py
------------------
결산 로직 및 2025년 검증 기준 테스트.

테스트 기준(요구사항 15번)에 대응:
  - 거래 입력/수정/삭제
  - soft delete 후 집계 제외
  - 수입/지출 합계, 차입금 분리
  - 당기수지차 / 기말현금 / 대차평형
  - 검증표 정상
  - 저장 전 백업 생성
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest  # noqa: E402

from scripts import sample_data  # noqa: E402
from scripts.seed_sample_data import build_seed_workbook  # noqa: E402
from src import (  # noqa: E402
    backup_service,
    config,
    excel_service,
    report_service,
    transaction_service,
    validation_service,
)

YEAR = 2025

# 2025년 검증 기준 값
EXPECTED = {
    "수입합계": 54_942_020,
    "지출합계": 43_656_222,
    "당기수지차": 11_285_798,
    "기말현금": 13_909_238,
    "차입금잔액": 9_000_000,
    "자산합계": 17_009_238,
    "부채자본합계": 17_009_238,
    "미수금": 3_100_000,
    "순자산": 8_009_238,
    "기초현금": 8_623_440,
    "차입유입": 9_000_000,
    "차입상환": 15_000_000,
}


@pytest.fixture
def wb():
    """샘플 데이터가 채워진 워크북(메모리)."""
    workbook, _ = build_seed_workbook()
    return workbook


def _summary(wb):
    return report_service.summary_from_workbook(wb, YEAR)


# ---------------------------------------------------------------------------
# 합계 / 결산 값
# ---------------------------------------------------------------------------

def test_income_total(wb):
    assert _summary(wb)["수입합계"] == EXPECTED["수입합계"]


def test_expense_total(wb):
    assert _summary(wb)["지출합계"] == EXPECTED["지출합계"]


def test_net_surplus(wb):
    s = _summary(wb)
    assert s["당기수지차"] == EXPECTED["당기수지차"]
    assert s["수입합계"] - s["지출합계"] == s["당기수지차"]


def test_closing_cash(wb):
    assert _summary(wb)["기말현금"] == EXPECTED["기말현금"]


def test_borrowing_balance(wb):
    s = _summary(wb)
    assert s["차입금잔액"] == EXPECTED["차입금잔액"]
    assert s["차입유입"] == EXPECTED["차입유입"]
    assert s["차입상환"] == EXPECTED["차입상환"]


def test_assets_equal_liabilities_and_equity(wb):
    s = _summary(wb)
    assert s["자산합계"] == EXPECTED["자산합계"]
    assert s["부채자본합계"] == EXPECTED["부채자본합계"]
    assert s["자산합계"] == s["부채자본합계"]
    # 부채 + 순자산 = 자산
    assert s["부채합계"] + s["순자산"] == s["자산합계"]


def test_receivables(wb):
    assert _summary(wb)["미수금"] == EXPECTED["미수금"]


# ---------------------------------------------------------------------------
# 차입금은 수입/지출에 포함되지 않는다
# ---------------------------------------------------------------------------

def test_borrowing_inflow_not_in_income(wb):
    """차입유입 거래가 있어도 수입합계에는 포함되지 않는다."""
    s = _summary(wb)
    assert s["차입유입"] > 0                       # 차입유입 거래 존재
    assert s["수입합계"] == EXPECTED["수입합계"]    # 수입합계는 순수 수입만


def test_repayment_not_in_expense(wb):
    """차입상환 거래가 있어도 지출합계에는 포함되지 않는다."""
    s = _summary(wb)
    assert s["차입상환"] > 0
    assert s["지출합계"] == EXPECTED["지출합계"]


# ---------------------------------------------------------------------------
# 거래 입력 / 수정 / 삭제
# ---------------------------------------------------------------------------

def test_transaction_id_format(wb):
    df = excel_service.read_ledger_df(wb)
    tid = transaction_service.generate_transaction_id(df, "2025-03-10")
    assert tid.startswith("TX-20250310-")
    assert len(tid.split("-")[-1]) == 4


def test_add_transaction_increases_income(wb):
    before = _summary(wb)["수입합계"]
    tx_id, errors = transaction_service.add_transaction(wb, {
        "회계연도": 2025, "거래일자": "2025-06-01", "입출구분": "수입",
        "대분류": "회비수입", "소분류": "연회비", "금액": 500_000,
        "증빙여부": "있음", "입력자": "tester",
    })
    assert errors == []
    assert tx_id is not None
    assert _summary(wb)["수입합계"] == before + 500_000


def test_update_transaction(wb):
    df = excel_service.read_ledger_df(wb)
    # 연회비 12,000,000 거래 찾기
    row = df[(df["소분류"] == "연회비")].iloc[0]
    tx_id = row["거래ID"]
    errors = transaction_service.update_transaction(wb, tx_id, {"금액": 13_000_000})
    assert errors == []
    df2 = excel_service.read_ledger_df(wb)
    updated = df2[df2["거래ID"] == tx_id].iloc[0]
    assert float(updated["금액"]) == 13_000_000
    assert updated["상태"] == config.STATUS_MODIFIED
    assert updated["수정일시"] != ""


def test_soft_delete_changes_status_not_row(wb):
    df = excel_service.read_ledger_df(wb)
    before_rows = len(df)
    tx_id = df[df["소분류"] == "연회비"].iloc[0]["거래ID"]

    ok = transaction_service.delete_transaction(wb, tx_id)
    assert ok

    df2 = excel_service.read_ledger_df(wb)
    assert len(df2) == before_rows                     # 행은 그대로 (soft delete)
    deleted = df2[df2["거래ID"] == tx_id].iloc[0]
    assert deleted["상태"] == config.STATUS_DELETED


def test_deleted_excluded_from_aggregation(wb):
    before = _summary(wb)["수입합계"]
    df = excel_service.read_ledger_df(wb)
    row = df[df["소분류"] == "연회비"].iloc[0]
    amount = float(row["금액"])
    transaction_service.delete_transaction(wb, row["거래ID"])
    after = _summary(wb)["수입합계"]
    assert after == before - amount


# ---------------------------------------------------------------------------
# 입력값 검증
# ---------------------------------------------------------------------------

def test_validate_missing_required():
    errors = transaction_service.validate_transaction({
        "회계연도": 2025, "거래일자": "", "입출구분": "수입",
        "대분류": "회비수입", "소분류": "연회비", "금액": 1000,
    })
    assert any("거래일자" in e for e in errors)


def test_validate_amount_must_be_positive():
    errors = transaction_service.validate_transaction({
        "회계연도": 2025, "거래일자": "2025-01-01", "입출구분": "수입",
        "대분류": "회비수입", "소분류": "연회비", "금액": 0,
    })
    assert any("금액" in e for e in errors)


def test_validate_unknown_account():
    errors = transaction_service.validate_transaction({
        "회계연도": 2025, "거래일자": "2025-01-01", "입출구분": "수입",
        "대분류": "없는대분류", "소분류": "없는소분류", "금액": 1000,
    })
    assert any("계정과목" in e for e in errors)


# ---------------------------------------------------------------------------
# 검증표
# ---------------------------------------------------------------------------

def test_validation_status_normal(wb):
    rows, overall, summary = validation_service.build_validation(wb, YEAR)
    assert overall == "정상"
    assert all(r["상태"] == "정상" for r in rows)


def test_validation_detects_amount_error(wb):
    # 잘못된 금액을 강제로 주입 후 검증
    df = excel_service.read_ledger_df(wb)
    df.loc[0, "금액"] = -100
    excel_service.write_df_to_sheet(wb, config.SHEET_LEDGER, df)
    rows, overall, _ = validation_service.build_validation(wb, YEAR)
    assert overall in ("오류", "확인필요")
    amount_row = next(r for r in rows if "금액 오류" in r["검증항목"])
    assert amount_row["상태"] == "오류"


# ---------------------------------------------------------------------------
# 백업
# ---------------------------------------------------------------------------

def test_backup_created_before_save(tmp_path, monkeypatch, wb):
    data_file = tmp_path / "data.xlsx"
    backup_dir = tmp_path / "backups"
    monkeypatch.setattr(config, "BACKUP_DIR", backup_dir)

    # 최초 저장: 기존 파일이 없으므로 백업 없음
    _saved, first_backup = excel_service.save_with_backup(wb, data_file)
    assert data_file.exists()
    assert first_backup is None

    # 두 번째 저장: 기존 파일이 있으므로 백업 생성
    _saved2, second_backup = excel_service.save_with_backup(wb, data_file)
    assert second_backup is not None
    assert second_backup.exists()
    assert second_backup.parent == backup_dir


def test_list_backups(tmp_path):
    backup_dir = tmp_path / "backups"
    backup_dir.mkdir()
    (backup_dir / "a.xlsx").write_text("x")
    (backup_dir / "b.xlsx").write_text("y")
    found = backup_service.list_backups(backup_dir)
    assert len(found) == 2


# ---------------------------------------------------------------------------
# 시트 생성 / 보고서
# ---------------------------------------------------------------------------

def test_all_report_sheets_exist(wb):
    report_service.generate_all_reports(wb, YEAR)
    for name in (config.SHEET_INCOME, config.SHEET_EXPENSE, config.SHEET_BORROWING,
                 config.SHEET_BALANCE, config.SHEET_VALIDATION, config.SHEET_REPORT):
        assert name in wb.sheetnames


def test_sample_transaction_count(wb):
    df = excel_service.read_ledger_df(wb)
    assert len(df) == len(sample_data.all_transactions())
