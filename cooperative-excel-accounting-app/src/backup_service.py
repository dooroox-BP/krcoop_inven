"""
backup_service.py
-----------------
엑셀 파일을 수정하기 전에 backups 폴더에 백업본을 생성한다.

- 백업 파일명 형식 : 협의회_결산_입력프로그램_엑셀양식_YYYYMMDD_HHMMSS.xlsx
- 최근 백업 목록 조회
- 오래된 백업 정리(선택)
"""

import shutil
from datetime import datetime
from pathlib import Path

from . import config


def create_backup(excel_path: Path = None, backup_dir: Path = None) -> Path | None:
    """
    현재 엑셀 파일을 백업 폴더로 복사한다.

    반환값:
        생성된 백업 파일 경로. 원본 파일이 없으면 None.
    """
    excel_path = Path(excel_path) if excel_path else config.EXCEL_PATH
    backup_dir = Path(backup_dir) if backup_dir else config.BACKUP_DIR

    if not excel_path.exists():
        # 아직 원본이 없으면 백업할 대상이 없음
        return None

    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime(config.BACKUP_TIME_FORMAT)
    stem = excel_path.stem  # 확장자 제외 파일명
    suffix = excel_path.suffix  # .xlsx
    backup_name = f"{stem}_{timestamp}{suffix}"
    backup_path = backup_dir / backup_name

    shutil.copy2(excel_path, backup_path)
    return backup_path


def list_backups(backup_dir: Path = None) -> list[Path]:
    """
    최근 백업 목록을 최신순으로 반환한다.
    """
    backup_dir = Path(backup_dir) if backup_dir else config.BACKUP_DIR
    if not backup_dir.exists():
        return []

    files = [p for p in backup_dir.glob("*.xlsx") if p.is_file()]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files


def prune_backups(keep: int = 30, backup_dir: Path = None) -> list[Path]:
    """
    최근 `keep`개의 백업만 남기고 나머지를 삭제한다.

    반환값:
        삭제된 파일 경로 목록.
    """
    files = list_backups(backup_dir)
    to_delete = files[keep:]
    for p in to_delete:
        try:
            p.unlink()
        except OSError:
            pass
    return to_delete
