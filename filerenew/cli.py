# -*- coding: utf-8 -*-
"""
cli.py
======
파일 정리기(엑셀 + PPT)의 명령줄 인터페이스.

엑셀과 PPT를 한 명령으로 함께 처리하며, 확장자를 보고 자동으로 알맞은
정리를 적용합니다.

사용 예
-------
    # 새 파일로 저장 (원본 보존, "<이름>_정리됨" 으로 저장)
    python -m filerenew "보고서.xlsx" "발표자료.pptx"

    # 원본 덮어쓰기(.bak 백업 생성)
    python -m filerenew --overwrite "파일.xlsx"

    # 엑셀: 특정 작업만 끄기 / 인쇄 영역도 삭제 / 계산 캐시 유지
    python -m filerenew --no-links --no-unhide "파일.xlsx"
    python -m filerenew --delete-print-areas --keep-calc-chain "파일.xlsx"

    # PPT: 대치 기본 폰트 지정
    python -m filerenew --font "맑은 고딕" "발표.pptx"
"""

from __future__ import annotations

import argparse
import glob

from filerenew.dispatcher import DEFAULT_FONT, clean_file


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="filerenew",
        description="엑셀/PPT 파일을 정리합니다. "
        "엑셀은 정의된 이름·외부 링크·숨겨진 시트를, "
        "PPT는 비표준/임베드 폰트를 정리합니다.",
    )
    parser.add_argument("files", nargs="+", help="정리할 파일 경로(여러 개·와일드카드 가능)")

    # 공통 저장 옵션
    parser.add_argument("--overwrite", action="store_true",
                        help="원본 덮어쓰기(.bak 백업 생성)")

    # 엑셀 전용 옵션
    excel = parser.add_argument_group("엑셀 옵션")
    excel.add_argument("--no-names", action="store_true", help="정의된 이름 삭제 안 함")
    excel.add_argument("--no-links", action="store_true", help="외부 링크 제거 안 함")
    excel.add_argument("--no-unhide", action="store_true", help="숨겨진 시트 표시 안 함")
    excel.add_argument("--delete-print-areas", action="store_true",
                       help="인쇄 영역(Print Area)도 함께 삭제 (기본은 보존)")
    excel.add_argument("--keep-calc-chain", action="store_true",
                       help="계산 순서 캐시(calcChain.xml)를 남김 (기본은 제거)")

    # PPT 전용 옵션
    ppt = parser.add_argument_group("PPT 옵션")
    ppt.add_argument("--font", default=DEFAULT_FONT,
                     help="대치할 기본 폰트 (기본값: %s)" % DEFAULT_FONT)
    ppt.add_argument("--no-replace-fonts", action="store_true",
                     help="비표준 폰트 치환 안 함")
    ppt.add_argument("--no-remove-embedded", action="store_true",
                     help="임베드 폰트 제거 안 함")
    return parser


def main(argv=None) -> int:
    args = _build_parser().parse_args(argv)

    # 와일드카드 확장 (셸이 확장하지 않는 환경 대비)
    paths: list[str] = []
    for pattern in args.files:
        matched = glob.glob(pattern)
        paths.extend(matched if matched else [pattern])

    failed = 0
    for path in paths:
        r = clean_file(
            path,
            overwrite=args.overwrite,
            # 엑셀
            delete_names=not args.no_names,
            remove_external_links=not args.no_links,
            unhide_sheets=not args.no_unhide,
            keep_print_areas=not args.delete_print_areas,
            drop_calc_chain=not args.keep_calc_chain,
            # PPT
            default_font=args.font,
            replace_fonts=not args.no_replace_fonts,
            remove_embedded=not args.no_remove_embedded,
        )
        print(r.summary())
        if r.ok and r.dst_path != r.src_path:
            print(f"        저장 위치: {r.dst_path}")
        for w in r.warnings:
            print(f"        ⚠ {w}")
        if not r.ok:
            failed += 1

    print(f"\n완료: 성공 {len(paths) - failed}개, 실패 {failed}개")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
