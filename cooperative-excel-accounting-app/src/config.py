"""
config.py
---------
프로젝트 전역 설정 값을 정의한다.

- 파일 경로 (엑셀 데이터 파일, 백업 폴더)
- 시트 이름 목록
- 거래입력대장 컬럼 정의
- 입출구분 목록 / 사업구분 목록 / 결제수단 목록
- 기본 계정과목 (입출구분 → 대분류 → 소분류)
- 기관 정보 및 기본 회계연도
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# 1. 파일 경로
# ---------------------------------------------------------------------------
# 프로젝트 루트 = 이 파일(src/config.py) 기준 상위 폴더
BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backups"

# 데이터 저장소로 사용하는 엑셀 파일
EXCEL_FILENAME = "협의회_결산_입력프로그램_엑셀양식.xlsx"
EXCEL_PATH = DATA_DIR / EXCEL_FILENAME

# ---------------------------------------------------------------------------
# 2. 기관 정보 / 기본 값
# ---------------------------------------------------------------------------
ORG_NAME = "인천광역시협동조합협의회"
DEFAULT_YEAR = 2025
MEETING_TITLE = "제13차 정기총회"

# ---------------------------------------------------------------------------
# 3. 시트 이름
# ---------------------------------------------------------------------------
SHEET_GUIDE = "사용안내"
SHEET_LEDGER = "거래입력대장"
SHEET_ACCOUNTS = "계정과목"
SHEET_OPENING = "기초잔액"
SHEET_ASSETS = "자산관리"
SHEET_INCOME = "수입결산"
SHEET_EXPENSE = "지출결산"
SHEET_BORROWING = "차입금현황"
SHEET_BALANCE = "간이대차대조표"
SHEET_VALIDATION = "검증표"
SHEET_REPORT = "총회보고서"
SHEET_LIST = "목록"

# 프로그램이 관리하는 전체 시트 (없으면 자동 생성)
ALL_SHEETS = [
    SHEET_GUIDE,
    SHEET_LEDGER,
    SHEET_ACCOUNTS,
    SHEET_OPENING,
    SHEET_ASSETS,
    SHEET_INCOME,
    SHEET_EXPENSE,
    SHEET_BORROWING,
    SHEET_BALANCE,
    SHEET_VALIDATION,
    SHEET_REPORT,
    SHEET_LIST,
]

# ---------------------------------------------------------------------------
# 4. 거래입력대장 컬럼
# ---------------------------------------------------------------------------
LEDGER_COLUMNS = [
    "거래ID",
    "회계연도",
    "거래일자",
    "입출구분",
    "대분류",
    "소분류",
    "금액",
    "거래처",
    "사업구분",
    "결제수단",
    "증빙여부",
    "증빙링크",
    "비고",
    "입력자",
    "입력일시",
    "수정일시",
    "상태",
]

# 상태값
STATUS_NORMAL = "정상"
STATUS_MODIFIED = "수정"
STATUS_DELETED = "삭제"
STATUS_VALUES = [STATUS_NORMAL, STATUS_MODIFIED, STATUS_DELETED]

# ---------------------------------------------------------------------------
# 5. 입출구분
# ---------------------------------------------------------------------------
INOUT_INCOME = "수입"
INOUT_EXPENSE = "지출"
INOUT_BORROW_IN = "차입유입"
INOUT_BORROW_REPAY = "차입상환"
INOUT_ASSET_UP = "자산증가"
INOUT_ASSET_DOWN = "자산감소"
INOUT_FUND_ADJUST = "자금조정"

INOUT_TYPES = [
    INOUT_INCOME,
    INOUT_EXPENSE,
    INOUT_BORROW_IN,
    INOUT_BORROW_REPAY,
    INOUT_ASSET_UP,
    INOUT_ASSET_DOWN,
    INOUT_FUND_ADJUST,
]

# ---------------------------------------------------------------------------
# 6. 사업구분 / 결제수단 / 증빙여부
# ---------------------------------------------------------------------------
BUSINESS_TYPES = [
    "일반운영",
    "활성화사업",
    "교육청사업",
    "행사사업",
    "기타사업",
]

PAYMENT_METHODS = [
    "계좌이체",
    "현금",
    "카드",
    "자동이체",
    "기타",
]

PROOF_YES = "있음"
PROOF_NO = "없음"
PROOF_VALUES = [PROOF_YES, PROOF_NO]

# ---------------------------------------------------------------------------
# 7. 기본 계정과목
# ---------------------------------------------------------------------------
# 각 항목: (입출구분, 대분류, 소분류)
#
# 회계 기준 참고:
#   - 차입유입 / 차입상환 : 수입/지출 합계에 포함하지 않고 차입금현황으로만 관리
#   - 자산증가 / 자산감소 : 자산관리 및 간이대차대조표에만 반영 (현금흐름과 분리)
#     · 대분류를 자산 유형(미수금 / 선급금)으로 통일하여 대차대조표 집계를 단순화
#   - 자금조정 : 현금잔액 검증에만 반영 (수입/지출 합계 제외)
ACCOUNTS = [
    # ---- 수입 ----
    (INOUT_INCOME, "회비수입", "연회비"),
    (INOUT_INCOME, "회비수입", "이사회비"),
    (INOUT_INCOME, "후원금수입", "사업후원금"),
    (INOUT_INCOME, "기부금수입", "기부금"),
    (INOUT_INCOME, "행사수입", "행사수입"),
    (INOUT_INCOME, "사업수입", "사업수입"),
    (INOUT_INCOME, "기타수입", "이자수입"),
    (INOUT_INCOME, "기타수입", "기타수입"),
    (INOUT_INCOME, "기타수입", "환급금"),
    (INOUT_INCOME, "기타수입", "사업비정산환급"),
    # ---- 지출 ----
    (INOUT_EXPENSE, "회의운영비", "회의진행비"),
    (INOUT_EXPENSE, "회의운영비", "회의비"),
    (INOUT_EXPENSE, "회의운영비", "화환"),
    (INOUT_EXPENSE, "활동비", "교통비"),
    (INOUT_EXPENSE, "활동비", "주유비"),
    (INOUT_EXPENSE, "활동비", "주차비"),
    (INOUT_EXPENSE, "사무국운영비", "임대료"),
    (INOUT_EXPENSE, "사무국운영비", "통신인터넷"),
    (INOUT_EXPENSE, "사무국운영비", "업무폰전화비"),
    (INOUT_EXPENSE, "사무국운영비", "일반수용비"),
    (INOUT_EXPENSE, "사무국운영비", "수수료"),
    (INOUT_EXPENSE, "인건비", "급여"),
    (INOUT_EXPENSE, "세금공과", "국세지방세"),
    (INOUT_EXPENSE, "사업비", "활성화사업자부담"),
    (INOUT_EXPENSE, "사업비", "교육청사업"),
    (INOUT_EXPENSE, "사업비", "세금납부"),
    (INOUT_EXPENSE, "사업비", "운영지원비"),
    (INOUT_EXPENSE, "기타지출", "오입금송금"),
    (INOUT_EXPENSE, "기타지출", "반환금"),
    (INOUT_EXPENSE, "기타지출", "사업비선급금"),
    # ---- 차입 ----
    (INOUT_BORROW_IN, "단기차입금", "단기차입금"),
    (INOUT_BORROW_REPAY, "단기차입금상환", "단기차입금상환"),
    # ---- 자산 (대분류 = 자산유형) ----
    (INOUT_ASSET_UP, "미수금", "미수금"),
    (INOUT_ASSET_UP, "선급금", "선급금"),
    (INOUT_ASSET_DOWN, "미수금", "미수금회수"),
    (INOUT_ASSET_DOWN, "선급금", "선급금정산"),
    # ---- 자금조정 ----
    (INOUT_FUND_ADJUST, "오입금", "오입금"),
    (INOUT_FUND_ADJUST, "환급금", "환급금"),
    (INOUT_FUND_ADJUST, "반환금", "반환금"),
]

# 자산 유형 (간이대차대조표 집계용)
ASSET_TYPE_RECEIVABLE = "미수금"   # 미수금
ASSET_TYPE_PREPAID = "선급금"      # 선급금

# ---------------------------------------------------------------------------
# 8. 계정과목 조회 헬퍼
# ---------------------------------------------------------------------------

def majors_for_inout(inout: str):
    """주어진 입출구분에 해당하는 대분류 목록(중복 제거, 입력 순서 유지)."""
    seen = []
    for io, major, _minor in ACCOUNTS:
        if io == inout and major not in seen:
            seen.append(major)
    return seen


def minors_for(inout: str, major: str):
    """주어진 입출구분·대분류에 해당하는 소분류 목록."""
    seen = []
    for io, mj, minor in ACCOUNTS:
        if io == inout and mj == major and minor not in seen:
            seen.append(minor)
    return seen


# ---------------------------------------------------------------------------
# 9. 기타 설정
# ---------------------------------------------------------------------------
# 거래ID 형식 : TX-YYYYMMDD-0001
TX_ID_PREFIX = "TX"

# 백업 파일명 형식 : 협의회_결산_입력프로그램_엑셀양식_YYYYMMDD_HHMMSS.xlsx
BACKUP_TIME_FORMAT = "%Y%m%d_%H%M%S"

# 기초잔액 항목 키
OPENING_CASH = "기초현금및예금"
OPENING_BORROWING = "기초단기차입금"
OPENING_RECEIVABLE = "기초미수금"
OPENING_PREPAID = "기초선급금"

OPENING_ITEMS = [
    OPENING_CASH,
    OPENING_BORROWING,
    OPENING_RECEIVABLE,
    OPENING_PREPAID,
]
