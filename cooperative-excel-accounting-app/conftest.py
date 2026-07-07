"""
conftest.py (프로젝트 루트)
---------------------------
pytest 실행 시 프로젝트 루트를 import 경로에 추가하여
`from src import ...`, `from scripts import ...` 가 동작하도록 한다.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
