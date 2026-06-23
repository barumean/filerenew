# -*- coding: utf-8 -*-
"""
ppt_cleaner.py
==============
PPTX(.pptx / .pptm / .potx) 파일의 폰트를 정리하는 핵심 로직 모듈입니다.

PPTX 파일은 내부적으로 ZIP 압축 안에 여러 XML 문서가 들어있는 구조입니다.
폰트는 다음 위치에서 참조됩니다.

  - ppt/theme/themeN.xml        : 테마의 majorFont / minorFont 정의
  - ppt/slides/slideN.xml 등    : <a:latin>, <a:ea>, <a:cs>, <a:sym>, <a:buFont> 의 typeface 속성
  - ppt/slideMasters, slideLayouts, notesSlides, handoutMasters
  - ppt/presentation.xml        : <p:embeddedFontLst> (임베드된 폰트 목록)
  - ppt/fonts/*.fntdata         : 실제 임베드된 폰트 바이너리

이 모듈은 위 참조들을 모두 찾아서
  1) 표준(안전) 폰트가 아닌 폰트를 기본 폰트(기본값: 맑은 고딕)로 치환하고
  2) 파일에 끼워넣어진(임베드) 폰트를 제거합니다.

OTF 임베드 폰트가 "TrueType 폰트가 아니므로 대치할 수 없습니다" 류의 경고를 내며
교체가 안 되는 문제를, 참조 자체를 기본 폰트로 바꿔버려서 한 번에 해결합니다.

엑셀 정리기와 동일하게 core.rewrite_zip 으로 "문제되는 부분만" 외과적으로
고치므로 슬라이드 내용·도형·차트 등 나머지는 원본 그대로 보존됩니다.
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
KIND = "PPT"

# ZIP 기반으로 처리 가능한 확장자.
SUPPORTED_EXTS = {".pptx", ".pptm", ".potx"}

# 기본 대치 폰트
DEFAULT_FONT = "맑은 고딕"

# 치환하지 않고 그대로 둘 "안전한" 폰트 목록(소문자 비교).
# - 대부분의 Windows / Office 환경에 기본 설치되어 있는 폰트
# - 기호/불릿용 폰트(이걸 바꾸면 기호가 깨지므로 반드시 보존)
SAFE_FONTS = {
    # 한글 기본
    "맑은 고딕", "맑은 고딕 semilight", "malgun gothic", "malgun gothic semilight",
    "굴림", "굴림체", "돋움", "돋움체", "바탕", "바탕체", "궁서", "궁서체",
    "gulim", "gulimche", "dotum", "dotumche", "batang", "batangche",
    "gungsuh", "gungsuhche",
    # 영문 기본 (Office 기본 테마 폰트 포함)
    "arial", "arial black", "arial narrow", "calibri", "calibri light",
    "cambria", "cambria math", "times new roman", "verdana", "tahoma",
    "segoe ui", "segoe ui light", "segoe ui semibold", "courier new",
    "georgia", "comic sans ms", "consolas", "trebuchet ms", "impact",
    "lucida sans unicode", "lucida console", "palatino linotype", "garamond",
    "century gothic", "book antiqua", "candara", "constantia", "corbel",
    "franklin gothic medium", "gabriola", "sylfaen",
    # 기호/불릿 폰트 (절대 치환 금지)
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "marlett", "mt extra",
}

# typeface="..." 속성을 찾는 정규식 (네임스페이스 접두사와 무관하게 동작)
_TYPEFACE_RE = re.compile(r'typeface="([^"]*)"')

# presentation.xml 의 임베드 폰트 목록 블록
_EMBEDDED_FONT_LST_RE = re.compile(
    r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", re.DOTALL
)
# 폰트 임베드 관련 속성 (자기닫힘/공백 포함 다양한 형태 제거)
_EMBED_ATTR_RE = re.compile(r'\s+(?:embedTrueTypeFonts|saveSubsetFonts)="[^"]*"')

# 폰트 참조가 들어있을 수 있는 XML 파트 경로 패턴
_FONT_BEARING_PATHS = (
    "ppt/theme/",
    "ppt/slides/",
    "ppt/slideMasters/",
    "ppt/slideLayouts/",
    "ppt/notesSlides/",
    "ppt/notesMasters/",
    "ppt/handoutMasters/",
    "ppt/charts/",
    "ppt/diagrams/",
)

PRESENTATION_XML = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
EMBEDDED_FONT_PREFIX = "ppt/fonts/"


def _should_replace(font_name, safe_lower):
    """이 폰트 이름을 치환해야 하는가?"""
    if font_name == "":
        return False
    # '+mn-lt', '+mj-ea' 같은 테마 참조는 그대로 둔다(테마 정의를 이미 고치므로).
    if font_name.startswith("+"):
        return False
    return font_name.strip().lower() not in safe_lower


def _replace_typefaces(xml_text, default_font, safe_lower, stats):
    def repl(m):
        name = m.group(1)
        if _should_replace(name, safe_lower):
            stats[name] = stats.get(name, 0) + 1
            return 'typeface="%s"' % default_font
        return m.group(0)

    return _TYPEFACE_RE.sub(repl, xml_text)


def _is_font_bearing_xml(name):
    if not name.endswith(".xml"):
        return False
    return any(name.startswith(p) for p in _FONT_BEARING_PATHS)


def _clean_presentation_xml(xml_text):
    """presentation.xml 에서 임베드 폰트 목록과 임베드 옵션을 제거."""
    xml_text = _EMBEDDED_FONT_LST_RE.sub("", xml_text)
    xml_text = _EMBED_ATTR_RE.sub("", xml_text)
    return xml_text


def _clean_presentation_rels(xml_text):
    """presentation.xml.rels 에서 폰트 파트를 가리키는 관계(Relationship) 제거."""
    # <Relationship ... Type=".../font" ... Target="fonts/..."/> 형태 제거
    rel_re = re.compile(r"<Relationship\b[^>]*/>")

    def repl(m):
        tag = m.group(0)
        if "/font" in tag or "fonts/" in tag.replace("\\", "/"):
            return ""
        return tag

    return rel_re.sub(repl, xml_text)


# ---------------------------------------------------------------------------
# 메인 처리 함수
# ---------------------------------------------------------------------------
def clean_presentation(
    src_path: str,
    dst_path: str | None = None,
    *,
    default_font: str = DEFAULT_FONT,
    replace_fonts: bool = True,
    remove_embedded: bool = True,
    overwrite: bool = False,
    backup: bool = True,
) -> CleanResult:
    """
    PPTX 파일 한 개의 폰트를 정리한다.

    Parameters
    ----------
    src_path : 원본 .pptx 경로
    dst_path : 저장 경로. None 이면 자동 결정(엑셀 정리기와 동일 규칙).
               - overwrite=False: 같은 폴더에 "<이름>_정리됨.<확장자>"
               - overwrite=True : 원본 경로 (backup=True 면 .bak 백업 생성)
    default_font   : 비표준 폰트를 치환할 기본 폰트.
    replace_fonts  : 비표준 폰트 → 기본 폰트 치환 수행 여부.
    remove_embedded: 임베드(끼워넣은) 폰트 제거 수행 여부.
    overwrite : 원본을 덮어쓸지 여부
    backup : overwrite=True 일 때 원본 백업(.bak) 생성 여부

    반환값: CleanResult (성공 여부·통계·경고)
    """
    result = CleanResult(src_path=src_path, kind=KIND)

    ext = os.path.splitext(src_path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        result.message = (
            f"지원하지 않는 형식({ext}). .pptx/.pptm/.potx 만 가능합니다."
        )
        return result

    if not zipfile.is_zipfile(src_path):
        result.message = "올바른 PPTX 파일이 아니거나 손상되어 ZIP 으로 열 수 없습니다."
        return result

    dst_path = decide_output_path(src_path, dst_path, overwrite=overwrite)
    result.dst_path = dst_path

    safe_lower = {f.lower() for f in SAFE_FONTS}
    # 기본 폰트 자신도 안전 목록에 포함(자기 자신으로 무한 치환 방지/일관성)
    safe_lower.add(default_font.strip().lower())

    replaced_stats: dict = {}
    removed_fonts: list = []

    def _transform(name: str, data: bytes):
        norm = name.replace("\\", "/")

        # 1) 임베드 폰트 바이너리 제거
        if remove_embedded and norm.startswith(EMBEDDED_FONT_PREFIX):
            removed_fonts.append(norm)
            return None

        # 2) presentation.xml : 임베드 폰트 목록/옵션 제거 + typeface 치환
        if norm == PRESENTATION_XML:
            text = data.decode("utf-8")
            if remove_embedded:
                text = _clean_presentation_xml(text)
            if replace_fonts:
                text = _replace_typefaces(text, default_font, safe_lower, replaced_stats)
            return text.encode("utf-8")

        # 3) presentation.xml.rels : 폰트 관계 제거
        if norm == PRESENTATION_RELS:
            if remove_embedded:
                text = _clean_presentation_rels(data.decode("utf-8"))
                return text.encode("utf-8")
            return data

        # 4) 폰트 참조가 있는 XML : typeface 치환
        if replace_fonts and _is_font_bearing_xml(norm):
            text = _replace_typefaces(data.decode("utf-8"), default_font, safe_lower, replaced_stats)
            return text.encode("utf-8")

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
        "폰트 치환": sum(replaced_stats.values()),
        "임베드 폰트 제거": len(removed_fonts),
    }
    result.ok = True
    result.message = "정상 처리되었습니다."
    return result
