# -*- coding: utf-8 -*-
"""
core.py
=======
엑셀/PPT 정리기가 공통으로 쓰는 자료구조와 헬퍼 모음.

excelrenew 와 pptrenew 는 둘 다 "ZIP(=XML 묶음) 안에서 문제되는 부분만
외과적으로 수정한다"는 동일한 메커니즘을 씁니다. 다만 결과 보고 형식,
저장 경로 규칙, 백업 방식이 서로 달랐는데, 이 모듈에서 그 부분을 하나로
통일합니다. 각 포맷별 "무엇을 고치는가"는 그대로 두고, "어떻게 보고하고
저장하는가"만 공통화합니다.

핵심 구성요소
-------------
- ``CleanResult`` : 파일 한 개 처리 결과(성공 여부·통계·경고)를 담는 공통 보고서.
- ``rewrite_zip``  : ZIP 을 멤버 단위로 순회하며 변환/삭제하는 공통 엔진.
- ``decide_output_path`` / ``finalize_output`` : 저장 경로 결정과 백업/원자적 교체.

표준 라이브러리만 사용하므로 별도 설치가 필요 없습니다.
"""

from __future__ import annotations

import os
import shutil
import zipfile
from dataclasses import dataclass, field

# 새 파일로 저장할 때 붙이는 공통 접미사. (이전: 엑셀 "_정리됨", PPT "_폰트정리")
DEFAULT_SUFFIX = "_정리됨"


@dataclass
class CleanResult:
    """
    파일 한 개를 정리한 결과 보고서(엑셀·PPT 공통).

    Attributes
    ----------
    src_path : 원본 파일 경로
    dst_path : 저장된 파일 경로
    kind     : 처리한 종류 표시용 라벨("엑셀" / "PPT")
    ok       : 성공 여부
    message  : 사람이 읽을 결과/오류 메시지
    stats    : 표시용 통계. {라벨: 개수} 형태이며 순서가 유지된다.
               예) {"정의된 이름 삭제": 3, "외부 링크 제거": 1}
    warnings : 부가 경고 메시지 목록
    """

    src_path: str
    dst_path: str = ""
    kind: str = ""
    ok: bool = False
    message: str = ""
    stats: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)

    def summary(self) -> str:
        """한 줄짜리 결과 요약 문자열."""
        name = os.path.basename(self.src_path)
        if not self.ok:
            return f"[실패] {name} → {self.message}"
        head = f"[완료] {name}"
        if self.kind:
            head += f" ({self.kind})"
        if self.stats:
            detail = ", ".join(f"{label} {count}개" for label, count in self.stats.items())
            return f"{head} → {detail}"
        return f"{head} → {self.message or '정상 처리되었습니다.'}"


# ---------------------------------------------------------------------------
# ZIP 처리 공통 엔진
# ---------------------------------------------------------------------------
def rewrite_zip(src_path: str, tmp_path: str, transform) -> None:
    """
    ZIP 파일을 멤버 단위로 순회하며 새 ZIP(tmp_path)을 만든다.

    transform(name, data) 콜백이 각 멤버를 어떻게 처리할지 결정한다.
      - ``bytes`` 반환  : 그 내용으로 기록(원본 압축 메타데이터 유지)
      - ``None`` 반환   : 해당 멤버를 결과에서 제외(=삭제)

    openpyxl/python-pptx 같은 라이브러리로 다시 저장하지 않고 ZIP 안의
    "문제되는 부분만" 고치므로 차트·매크로·서식 등 나머지는 원본 그대로
    보존된다. 오류가 나면 임시 파일을 지우고 예외를 다시 던진다.
    """
    try:
        with zipfile.ZipFile(src_path, "r") as zin, \
                zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                new_data = transform(item.filename, data)
                if new_data is None:
                    continue  # 이 멤버는 제거
                zout.writestr(item, new_data)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


# ---------------------------------------------------------------------------
# 저장 경로 / 백업 / 원자적 교체
# ---------------------------------------------------------------------------
def decide_output_path(
    src_path: str,
    dst_path: str | None,
    *,
    overwrite: bool,
    suffix: str = DEFAULT_SUFFIX,
) -> str:
    """
    저장 경로를 결정한다.

    - dst_path 가 주어지면 그대로 사용.
    - overwrite=True  : 원본 경로(덮어쓰기).
    - overwrite=False : 같은 폴더에 "<이름><접미사><확장자>".
    """
    if dst_path is not None:
        return dst_path
    if overwrite:
        return src_path
    root, ext = os.path.splitext(src_path)
    return f"{root}{suffix}{ext}"


def finalize_output(
    tmp_path: str,
    dst_path: str,
    src_path: str,
    *,
    overwrite: bool,
    backup: bool,
    warnings: list,
) -> None:
    """
    임시 파일을 최종 위치로 옮긴다. 원본을 덮어쓰는 경우 백업(.bak)을 만든다.

    같은 이름의 .bak 이 이미 있으면 .bak1, .bak2 … 로 번호를 붙여 보존하고
    그 사실을 warnings 에 남긴다.
    """
    try:
        same_as_src = os.path.abspath(dst_path) == os.path.abspath(src_path)
        if overwrite and backup and same_as_src:
            bak = src_path + ".bak"
            if not os.path.exists(bak):
                shutil.copy2(src_path, bak)
            else:
                i = 1
                while os.path.exists(f"{src_path}.bak{i}"):
                    i += 1
                shutil.copy2(src_path, f"{src_path}.bak{i}")
                warnings.append(f"백업: {os.path.basename(src_path)}.bak{i}")
        os.replace(tmp_path, dst_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise
