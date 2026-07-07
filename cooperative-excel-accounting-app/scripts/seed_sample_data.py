"""
seed_sample_data.py
-------------------
2025년 샘플 데이터를 담은 엑셀 데이터 파일을 생성한다.

실행:
    (프로젝트 루트에서)
    python scripts/seed_sample_data.py

동작:
  1. 기존 데이터 파일 삭제(있으면)
  2. 필수 시트/헤더 생성
  3. 계정과목 / 기초잔액 / 목록 / 사용안내 시트 채우기
  4. 샘플 거래 입력
  5. 전체 보고서 + 검증표 생성
  6. 저장 후 검증 요약 출력
"""

import sys
from pathlib import Path

# 프로젝트 루트를 import 경로에 추가
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd  # noqa: E402

from scripts import sample_data  # noqa: E402
from src import (  # noqa: E402
    config,
    excel_service,
    report_service,
    transaction_service,
    validation_service,
)


def _write_accounts(wb):
    rows = [{"입출구분": io, "대분류": mj, "소분류": mn}
            for io, mj, mn in config.ACCOUNTS]
    df = pd.DataFrame(rows, columns=["입출구분", "대분류", "소분류"])
    excel_service.write_df_to_sheet(wb, config.SHEET_ACCOUNTS, df)


def _write_opening(wb):
    rows = [{"회계연도": sample_data.YEAR, "항목": k, "금액": v}
            for k, v in sample_data.OPENING_BALANCES.items()]
    df = pd.DataFrame(rows, columns=["회계연도", "항목", "금액"])
    excel_service.write_df_to_sheet(wb, config.SHEET_OPENING, df)


def _write_lists(wb):
    cols = {
        "입출구분": config.INOUT_TYPES,
        "사업구분": config.BUSINESS_TYPES,
        "결제수단": config.PAYMENT_METHODS,
        "증빙여부": config.PROOF_VALUES,
        "상태": config.STATUS_VALUES,
    }
    maxlen = max(len(v) for v in cols.values())
    data = {k: (v + [""] * (maxlen - len(v))) for k, v in cols.items()}
    df = pd.DataFrame(data)
    excel_service.write_df_to_sheet(wb, config.SHEET_LIST, df)


def _write_guide(wb):
    lines = [
        [f"{config.ORG_NAME} 결산 입력 프로그램"],
        [],
        ["■ 사용 순서"],
        ["1. 프로그램 실행 전 이 엑셀 파일을 반드시 닫아 주세요."],
        ["2. 웹 화면(거래 입력)에서 거래를 입력합니다."],
        ["3. 저장하면 거래입력대장에 기록되고 백업이 자동 생성됩니다."],
        ["4. 각 결산/보고서 화면에서 '갱신' 시 시트가 자동으로 다시 계산됩니다."],
        [],
        ["■ 회계 기준"],
        ["- 차입금 유입/상환은 수입·지출 합계에 포함하지 않습니다(차입금현황으로 관리)."],
        ["- 자산증가/감소는 간이대차대조표에만 반영합니다."],
        ["- 자금조정은 현금잔액 검증에만 반영합니다."],
        ["- 삭제 거래(상태=삭제)는 집계에서 제외합니다."],
        [],
        ["■ 주의"],
        ["프로그램 실행 중에는 이 엑셀 파일을 열어두지 마세요(저장 충돌 방지)."],
    ]
    excel_service.write_grid_to_sheet(wb, config.SHEET_GUIDE, lines)


def build_seed_workbook():
    """
    샘플 데이터가 채워진 워크북을 메모리 상에서 만들어 반환한다.
    (파일 저장/삭제 없음 — 테스트에서 재사용 가능)
    반환: (wb, added_count)
    """
    wb = excel_service.new_workbook()

    # 참조 시트 채우기
    _write_accounts(wb)
    _write_opening(wb)
    _write_lists(wb)
    _write_guide(wb)

    # 샘플 거래 입력
    added = 0
    for tx in sample_data.all_transactions():
        tx_id, errors = transaction_service.add_transaction(wb, tx)
        if errors:
            print(f"[경고] 거래 입력 실패: {errors} / {tx}")
        else:
            added += 1

    # 보고서 + 검증표 생성
    report_service.generate_all_reports(wb, sample_data.YEAR)
    validation_service.generate_validation_report(wb, sample_data.YEAR)
    return wb, added


def seed():
    # 1. 기존 파일 삭제
    if config.EXCEL_PATH.exists():
        config.EXCEL_PATH.unlink()

    # 2. 워크북 구성 (구조 + 참조시트 + 거래 + 보고서)
    wb, added = build_seed_workbook()

    # 3. 요약/검증 재계산 (출력용)
    rows, overall, summary = validation_service.build_validation(
        wb, sample_data.YEAR)

    # 4. 저장 (초기 시드는 백업 없이 저장)
    excel_service.save_workbook(wb)

    # 결과 출력
    print("=" * 60)
    print(f"샘플 데이터 생성 완료: {config.EXCEL_PATH}")
    print(f"입력 거래 수: {added}건")
    print("-" * 60)
    print(f"수입합계      : {summary['수입합계']:>15,.0f} 원")
    print(f"지출합계      : {summary['지출합계']:>15,.0f} 원")
    print(f"당기수지차    : {summary['당기수지차']:>15,.0f} 원")
    print(f"기초현금      : {summary['기초현금']:>15,.0f} 원")
    print(f"차입유입      : {summary['차입유입']:>15,.0f} 원")
    print(f"차입상환      : {summary['차입상환']:>15,.0f} 원")
    print(f"기말현금      : {summary['기말현금']:>15,.0f} 원")
    print(f"차입금잔액    : {summary['차입금잔액']:>15,.0f} 원")
    print(f"미수금        : {summary['미수금']:>15,.0f} 원")
    print(f"자산합계      : {summary['자산합계']:>15,.0f} 원")
    print(f"부채·자본합계 : {summary['부채자본합계']:>15,.0f} 원")
    print(f"순자산        : {summary['순자산']:>15,.0f} 원")
    print("-" * 60)
    print(f"종합 검증상태 : {overall}")
    for r in rows:
        print(f"  [{r['상태']}] {r['검증항목']}")
    print("=" * 60)
    return summary, overall


if __name__ == "__main__":
    seed()
