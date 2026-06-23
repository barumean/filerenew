# -*- coding: utf-8 -*-
"""
PPT 폰트 정리 로직(filerenew.ppt_cleaner)의 동작을 검증하는 단위 테스트.

엑셀 테스트와 동일하게 unittest 로 통일했습니다.
    python -m unittest discover -s tests -v
"""

import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from filerenew import clean_file
from filerenew.ppt_cleaner import clean_presentation

PRESENTATION_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:p="ns" embedTrueTypeFonts="1" saveSubsetFonts="1">'
    '<p:embeddedFontLst>'
    '<p:embeddedFont><p:font typeface="SomeWeirdOTF"/>'
    '<p:regular r:id="rId99"/></p:embeddedFont>'
    '</p:embeddedFontLst>'
    '<p:sldIdLst/></p:presentation>'
)

PRESENTATION_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="r">'
    '<Relationship Id="rId1" Type="http://x/slide" Target="slides/slide1.xml"/>'
    '<Relationship Id="rId99" Type="http://x/font" Target="fonts/font1.fntdata"/>'
    '</Relationships>'
)

THEME_XML = (
    '<?xml version="1.0"?>'
    '<a:theme xmlns:a="ns"><a:themeElements><a:fontScheme>'
    '<a:majorFont><a:latin typeface="Weird OTF Font"/>'
    '<a:ea typeface=""/><a:cs typeface=""/>'
    '<a:font script="Hang" typeface="이상한폰트"/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme></a:themeElements></a:theme>'
)

SLIDE_XML = (
    '<?xml version="1.0"?>'
    '<p:sld xmlns:p="ns" xmlns:a="ns2"><p:cSld><p:spTree>'
    '<a:p><a:r><a:rPr><a:latin typeface="HY견고딕"/>'
    '<a:ea typeface="HY견고딕"/><a:cs typeface="+mn-cs"/></a:rPr>'
    '<a:t>안녕</a:t></a:r></a:p>'
    '<a:p><a:pPr><a:buFont typeface="Wingdings"/></a:pPr>'
    '<a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>hi</a:t></a:r></a:p>'
    '</p:spTree></p:cSld></p:sld>'
)


def _make_sample_pptx(path: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("ppt/presentation.xml", PRESENTATION_XML)
        z.writestr("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS)
        z.writestr("ppt/theme/theme1.xml", THEME_XML)
        z.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        z.writestr("ppt/fonts/font1.fntdata", b"\x00\x01\x02FAKEFONT")


class TestCleanPresentation(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "sample.pptx")
        _make_sample_pptx(self.src)

    def _read(self, path, member):
        with zipfile.ZipFile(path) as z:
            return z.read(member).decode("utf-8")

    def _names(self, path):
        with zipfile.ZipFile(path) as z:
            return z.namelist()

    def test_result_shape(self):
        r = clean_presentation(self.src)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(r.kind, "PPT")
        self.assertEqual(r.stats["임베드 폰트 제거"], 1)
        self.assertGreater(r.stats["폰트 치환"], 0)

    def test_embedded_font_removed(self):
        r = clean_presentation(self.src)
        self.assertNotIn("ppt/fonts/font1.fntdata", self._names(r.dst_path))

    def test_presentation_embed_cleaned(self):
        r = clean_presentation(self.src)
        pres = self._read(r.dst_path, "ppt/presentation.xml")
        self.assertNotIn("embeddedFontLst", pres)
        self.assertNotIn("embedTrueTypeFonts", pres)

    def test_font_rels_removed_slide_rels_kept(self):
        r = clean_presentation(self.src)
        rels = self._read(r.dst_path, "ppt/_rels/presentation.xml.rels")
        self.assertNotIn("fonts/font1.fntdata", rels)
        self.assertNotIn("rId99", rels)
        self.assertIn("slides/slide1.xml", rels)  # 슬라이드 관계는 유지

    def test_theme_replaced_and_preserved(self):
        r = clean_presentation(self.src)
        theme = self._read(r.dst_path, "ppt/theme/theme1.xml")
        self.assertNotIn("Weird OTF Font", theme)
        self.assertNotIn("이상한폰트", theme)
        self.assertIn("맑은 고딕", theme)        # 비표준 latin 치환
        self.assertIn('typeface="Calibri"', theme)  # 표준 폰트 유지
        self.assertIn('typeface=""', theme)         # 빈 typeface 유지

    def test_slide_replaced_and_preserved(self):
        r = clean_presentation(self.src)
        slide = self._read(r.dst_path, "ppt/slides/slide1.xml")
        self.assertNotIn("HY견고딕", slide)             # 비표준 치환됨
        self.assertIn('typeface="Arial"', slide)        # 표준 폰트 유지
        self.assertIn('typeface="Wingdings"', slide)    # 기호/불릿 폰트 유지
        self.assertIn('typeface="+mn-cs"', slide)       # 테마 참조 유지

    def test_custom_default_font(self):
        r = clean_presentation(self.src, default_font="굴림")
        theme = self._read(r.dst_path, "ppt/theme/theme1.xml")
        self.assertIn('typeface="굴림"', theme)

    def test_default_output_is_new_file(self):
        r = clean_presentation(self.src)
        self.assertTrue(r.dst_path.endswith("_정리됨.pptx"))
        self.assertTrue(os.path.exists(self.src))  # 원본 보존

    def test_overwrite_creates_backup(self):
        r = clean_presentation(self.src, overwrite=True, backup=True)
        self.assertEqual(os.path.abspath(r.dst_path), os.path.abspath(self.src))
        self.assertTrue(os.path.exists(self.src + ".bak"))

    def test_unsupported_extension(self):
        bad = os.path.join(self.tmp, "doc.odp")
        with open(bad, "wb") as f:
            f.write(b"not a zip")
        r = clean_presentation(bad)
        self.assertFalse(r.ok)
        self.assertIn("지원하지 않는", r.message)

    def test_dispatch_routes_ppt(self):
        # 통합 진입점(clean_file)이 확장자로 PPT 정리기에 라우팅하는지 확인
        r = clean_file(self.src)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(r.kind, "PPT")
        self.assertEqual(r.stats["임베드 폰트 제거"], 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
