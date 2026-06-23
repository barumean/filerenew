# -*- coding: utf-8 -*-
"""
filerenew
=========
엑셀과 PPT를 한 프로그램에서 정리하는 파일 정리기.

- excelrenew : 깨진 정의된 이름 · 외부 링크 · 숨겨진 시트 정리
- pptrenew   : 비표준/임베드 폰트 정리

두 도구의 "ZIP 안에서 문제되는 부분만 외과적으로 수정한다"는 메커니즘은
그대로 유지하면서, 결과 보고 형식·저장 규칙·입력 옵션을 하나로 통일했습니다.
"""

from filerenew.core import CleanResult
from filerenew.dispatcher import (
    DEFAULT_FONT,
    EXCEL_EXTS,
    PPT_EXTS,
    SUPPORTED_EXTS,
    clean_file,
    clean_many,
    kind_of,
)

__all__ = [
    "CleanResult",
    "clean_file",
    "clean_many",
    "kind_of",
    "DEFAULT_FONT",
    "EXCEL_EXTS",
    "PPT_EXTS",
    "SUPPORTED_EXTS",
]

__version__ = "1.0.0"
