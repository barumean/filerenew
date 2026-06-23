# -*- coding: utf-8 -*-
"""
excel_cleaner.py
================
엑셀(.xlsx / .xlsm / .xltx / .xltm) 파일에서 오류를 유발하는 요소를
한 번에 정리해 "정상 파일"로 되돌려 주는 핵심 로직 모듈입니다.

정리 대상
---------
1. 정의된 이름(Defined Names) 전부 삭제
   - 숨겨진 이름, #REF! 로 깨진 이름, 전역/시트범위 이름 모두 포함.
   - (VBA 의 Names.Delete 와 동일한 효과를 파일 레벨에서 수행)
2. 외부 링크/연결(External Links) 제거
   - 다른 통합문서를 참조하는 외부 링크 정의와 관계(rels)를 제거.
3. 숨겨진 시트 다시 표시(Unhide)
   - hidden / veryHidden 상태의 시트를 모두 visible 로 복구.

설계상의 핵심
-------------
.xlsx 계열 파일은 사실 여러 XML 파일을 담은 ZIP 압축 파일입니다.
openpyxl 같은 라이브러리로 다시 저장하면 차트·피벗테이블·매크로 등이
손실될 수 있으므로, 여기서는 ZIP 안의 "문제되는 부분"만 외과적으로
수정하고 나머지는 원본 그대로 복사합니다(core.rewrite_zip 사용). 따라서
외부 의존성이 전혀 없고(표준 라이브러리만 사용), 원본 콘텐츠 손실 위험이
최소화됩니다.
"""

from __future__ import annotations

import os
import re
import zipfile

from filerenew.core import (
    CleanResult,
    decide_output_path,
    finalize_output,
    rewrite_zip,
)

# 처리 종류 라벨(결과 보고용)
KIND = "엑셀"

# ZIP 기반(=XML 묶음)으로 처리 가능한 확장자. .xls(구형 이진 포맷)는 제외.
SUPPORTED_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
EXTERNAL_LINKS_PREFIX = "xl/externalLinks/"
CALC_CHAIN = "xl/calcChain.xml"

# 엑셀이 내부적으로 "정의된 이름"으로 저장하는 인쇄 영역/제목.
# 이 이름들은 삭제 대상에서 제외할 수 있다(인쇄 영역 유지).
PRINT_BUILTIN_NAMES = ("_xlnm.Print_Area", "_xlnm.Print_Titles")


# ---------------------------------------------------------------------------
# workbook.xml 수정 헬퍼
# ---------------------------------------------------------------------------
def _strip_defined_names(xml: str, keep_print_areas: bool = True) -> tuple[str, int]:
    """
    <definedNames> 안의 개별 <definedName> 항목을 제거한다.

    keep_print_areas=True 이면 인쇄 영역/제목(_xlnm.Print_Area,
    _xlnm.Print_Titles)은 남겨 두고 나머지 이름만 삭제한다.
    삭제한 이름 개수를 함께 반환한다.
    """
    count = 0

    def _process_block(block_match: re.Match) -> str:
        nonlocal count
        block = block_match.group(0)

        kept_entries: list[str] = []

        def _each_name(entry_match: re.Match) -> str:
            nonlocal count
            entry = entry_match.group(0)
            name_attr = re.search(r'\bname="([^"]*)"', entry)
            name = name_attr.group(1) if name_attr else ""
            if keep_print_areas and name in PRINT_BUILTIN_NAMES:
                kept_entries.append(entry)  # 인쇄 영역은 보존
            else:
                count += 1  # 삭제
            return ""

        # 블록 내부의 개별 <definedName> 항목을 순회
        re.sub(
            r"<definedName\b[^>]*?(?:/>|>.*?</definedName>)",
            _each_name,
            block,
            flags=re.DOTALL,
        )

        if kept_entries:
            return "<definedNames>" + "".join(kept_entries) + "</definedNames>"
        return ""  # 남길 이름이 없으면 블록 자체 제거

    # <definedNames> ... </definedNames>
    xml = re.sub(
        r"<definedNames\b.*?</definedNames>",
        _process_block,
        xml,
        flags=re.DOTALL,
    )
    # 비어있는 self-closing 형태도 제거 (<definedNames/>)
    xml = re.sub(r"<definedNames\b[^>]*/>", "", xml)
    return xml, count


def _strip_external_references(xml: str) -> str:
    """workbook.xml 내 <externalReferences> 블록 제거."""
    xml = re.sub(
        r"<externalReferences\b.*?</externalReferences>",
        "",
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(r"<externalReferences\b[^>]*/>", "", xml)
    return xml


def _unhide_sheets(xml: str) -> tuple[str, int]:
    """<sheet> 요소의 state="hidden"/"veryHidden" 속성을 제거(=visible)."""
    matches = re.findall(r'\sstate="(?:hidden|veryHidden)"', xml)
    xml = re.sub(r'\sstate="(?:hidden|veryHidden)"', "", xml)
    return xml, len(matches)


# ---------------------------------------------------------------------------
# 관계(rels) / 콘텐츠 형식 수정 헬퍼
# ---------------------------------------------------------------------------
def _strip_external_link_rels(rels_xml: str) -> tuple[str, int]:
    """workbook.xml.rels 에서 externalLink 관계 항목을 제거."""
    pattern = r"<Relationship\b[^>]*externalLink[^>]*/>"
    count = len(re.findall(pattern, rels_xml))
    rels_xml = re.sub(pattern, "", rels_xml)
    return rels_xml, count


def _strip_external_link_content_types(ct_xml: str) -> str:
    """[Content_Types].xml 에서 externalLink Override 항목 제거."""
    return re.sub(r"<Override\b[^>]*externalLink[^>]*/>", "", ct_xml)


def _strip_calc_chain_rels(rels_xml: str) -> str:
    """workbook.xml.rels 에서 calcChain 관계 항목을 제거."""
    return re.sub(r"<Relationship\b[^>]*calcChain[^>]*/>", "", rels_xml)


def _strip_calc_chain_content_types(ct_xml: str) -> str:
    """[Content_Types].xml 에서 calcChain Override 항목 제거."""
    return re.sub(r"<Override\b[^>]*calcChain[^>]*/>", "", ct_xml)


# ---------------------------------------------------------------------------
# 메인 처리 함수
# ---------------------------------------------------------------------------
def clean_workbook(
    src_path: str,
    dst_path: str | None = None,
    *,
    delete_names: bool = True,
    remove_external_links: bool = True,
    unhide_sheets: bool = True,
    keep_print_areas: bool = True,
    drop_calc_chain: bool = True,
    overwrite: bool = False,
    backup: bool = True,
) -> CleanResult:
    """
    엑셀 파일 한 개를 정리한다.

    Parameters
    ----------
    src_path : 원본 파일 경로
    dst_path : 저장 경로. None 이면 자동 결정.
               - overwrite=False: 같은 폴더에 "<이름>_정리됨.<확장자>"
               - overwrite=True : 원본 경로 (backup=True 면 .bak 백업 생성)
    delete_names, remove_external_links, unhide_sheets : 수행할 작업 선택
    keep_print_areas : 인쇄 영역/제목(_xlnm.Print_Area/_xlnm.Print_Titles)은
                       삭제하지 않고 보존(기본값 True)
    drop_calc_chain : 계산 순서 캐시(xl/calcChain.xml)를 제거. 구조가 바뀌면
                      엑셀이 "제거된 레코드(계산 속성)" 경고를 띄우는데, 이
                      파일을 미리 지워 두면 경고가 사라지고 엑셀이 자동
                      재생성한다(기본값 True).
    overwrite : 원본을 덮어쓸지 여부
    backup : overwrite=True 일 때 원본 백업(.bak) 생성 여부

    반환값: CleanResult (성공 여부·통계·경고)
    """
    result = CleanResult(src_path=src_path, kind=KIND)

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        result.message = (
            f"지원하지 않는 형식({ext}). .xlsx/.xlsm/.xltx/.xltm 만 가능합니다. "
            "구형 .xls 파일은 먼저 .xlsx 로 저장해 주세요."
        )
        return result

    if not zipfile.is_zipfile(src_path):
        result.message = "올바른 엑셀 파일이 아니거나 손상되어 ZIP 으로 열 수 없습니다."
        return result

    dst_path = decide_output_path(src_path, dst_path, overwrite=overwrite)
    result.dst_path = dst_path

    names_removed = 0
    external_links_removed = 0
    sheets_unhidden = 0

    def _transform(name: str, data: bytes):
        nonlocal names_removed, external_links_removed, sheets_unhidden

        # 1) 외부 링크 관련 파일은 통째로 제외(링크 파일 + 해당 _rels)
        if remove_external_links and name.startswith(EXTERNAL_LINKS_PREFIX):
            return None

        # 1-2) 계산 순서 캐시는 제거(엑셀이 재생성 → 계산속성 경고 방지)
        if drop_calc_chain and name == CALC_CHAIN:
            return None

        # 2) workbook.xml: 이름/외부참조 제거, 시트 숨김 해제
        if name == WORKBOOK_XML:
            xml = data.decode("utf-8")
            if delete_names:
                xml, n = _strip_defined_names(xml, keep_print_areas)
                names_removed = n
            if remove_external_links:
                xml = _strip_external_references(xml)
            if unhide_sheets:
                xml, n = _unhide_sheets(xml)
                sheets_unhidden = n
            return xml.encode("utf-8")

        # 3) workbook.xml.rels: 외부 링크 + calcChain 관계 제거
        #    (외부 링크는 연결 통합문서 1개당 관계 1개 = 정확한 링크 수)
        if name == WORKBOOK_RELS:
            rels = data.decode("utf-8")
            if remove_external_links:
                rels, n = _strip_external_link_rels(rels)
                external_links_removed = n
            if drop_calc_chain:
                rels = _strip_calc_chain_rels(rels)
            return rels.encode("utf-8")

        # 4) [Content_Types].xml: 외부 링크 + calcChain Override 제거
        if name == CONTENT_TYPES:
            ct = data.decode("utf-8")
            if remove_external_links:
                ct = _strip_external_link_content_types(ct)
            if drop_calc_chain:
                ct = _strip_calc_chain_content_types(ct)
            return ct.encode("utf-8")

        # 그 외는 원본 그대로 복사
        return data

    tmp_path = dst_path + ".tmp_renew"
    try:
        rewrite_zip(src_path, tmp_path, _transform)
    except Exception as exc:  # noqa: BLE001
        result.message = f"처리 중 오류: {exc}"
        return result

    try:
        finalize_output(
            tmp_path, dst_path, src_path,
            overwrite=overwrite, backup=backup, warnings=result.warnings,
        )
    except Exception as exc:  # noqa: BLE001
        result.message = f"파일 저장 오류: {exc}"
        return result

    result.stats = {
        "정의된 이름 삭제": names_removed,
        "외부 링크 제거": external_links_removed,
        "숨겨진 시트 표시": sheets_unhidden,
    }
    result.ok = True
    result.message = "정상 처리되었습니다."
    return result
