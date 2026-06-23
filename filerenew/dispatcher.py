# -*- coding: utf-8 -*-
"""
dispatcher.py
=============
파일 확장자를 보고 적절한 정리기(엑셀/PPT)로 보내 주는 공통 진입점.

GUI 와 CLI 는 모두 이 모듈의 ``clean_file`` 한 함수를 호출하면 되고, 결과는
포맷과 무관하게 동일한 ``CleanResult`` 로 돌아옵니다(통일된 보고 형식).
"""

from __future__ import annotations

import os

from filerenew.core import CleanResult
from filerenew import excel_cleaner, ppt_cleaner

# 포맷별 지원 확장자
EXCEL_EXTS = excel_cleaner.SUPPORTED_EXTS
PPT_EXTS = ppt_cleaner.SUPPORTED_EXTS
SUPPORTED_EXTS = EXCEL_EXTS | PPT_EXTS

DEFAULT_FONT = ppt_cleaner.DEFAULT_FONT


def kind_of(path: str) -> str | None:
    """경로의 확장자로 처리 종류를 판별. ('excel'/'ppt'/None)"""
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXTS:
        return "excel"
    if ext in PPT_EXTS:
        return "ppt"
    return None


def clean_file(
    src_path: str,
    dst_path: str | None = None,
    *,
    overwrite: bool = False,
    backup: bool = True,
    # 엑셀 옵션
    delete_names: bool = True,
    remove_external_links: bool = True,
    unhide_sheets: bool = True,
    keep_print_areas: bool = True,
    drop_calc_chain: bool = True,
    # PPT 옵션
    default_font: str = DEFAULT_FONT,
    replace_fonts: bool = True,
    remove_embedded: bool = True,
) -> CleanResult:
    """
    파일 한 개를 확장자에 맞는 정리기로 처리한다.

    엑셀 옵션은 PPT 처리 시 무시되고, PPT 옵션은 엑셀 처리 시 무시된다.
    지원하지 않는 확장자면 실패한 CleanResult 를 돌려준다.
    """
    kind = kind_of(src_path)
    if kind == "excel":
        return excel_cleaner.clean_workbook(
            src_path, dst_path,
            delete_names=delete_names,
            remove_external_links=remove_external_links,
            unhide_sheets=unhide_sheets,
            keep_print_areas=keep_print_areas,
            drop_calc_chain=drop_calc_chain,
            overwrite=overwrite,
            backup=backup,
        )
    if kind == "ppt":
        return ppt_cleaner.clean_presentation(
            src_path, dst_path,
            default_font=default_font,
            replace_fonts=replace_fonts,
            remove_embedded=remove_embedded,
            overwrite=overwrite,
            backup=backup,
        )

    ext = os.path.splitext(src_path)[1].lower()
    result = CleanResult(src_path=src_path)
    result.message = (
        f"지원하지 않는 형식({ext}). "
        "엑셀(.xlsx/.xlsm/.xltx/.xltm) 또는 PPT(.pptx/.pptm/.potx)만 가능합니다."
    )
    return result


def clean_many(paths, **kwargs) -> list[CleanResult]:
    """여러 파일을 순서대로 정리하고 결과 리스트를 반환."""
    return [clean_file(p, **kwargs) for p in paths]
