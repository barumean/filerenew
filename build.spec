# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 빌드 스펙.

빌드:
    pyinstaller build.spec

또는 build.bat 실행. 결과물: dist/파일정리기.exe (단일 실행 파일)

※ 스펙 파일 이름은 ASCII(build.spec)로 둡니다. Windows 배치 파일에서
   한글 파일명을 인자로 넘기면 인코딩이 깨져 "Spec file not found" 오류가
   나기 때문입니다. 최종 exe 이름은 아래 EXE(name=...) 에서 한글로 지정합니다.
"""

from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# 드래그앤드롭 라이브러리(tkdnd 네이티브 바이너리 포함)를 통째로 수집.
# 설치되어 있지 않으면 건너뛰고, exe는 파일선택 방식으로 동작한다.
for _pkg in ("tkinterdnd2",):
    try:
        _d, _b, _h = collect_all(_pkg)
        datas += _d
        binaries += _b
        hiddenimports += _h
    except Exception as _e:  # noqa: BLE001
        print("[spec] %s 수집 건너뜀: %s" % (_pkg, _e))


a = Analysis(
    ["run_gui.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="파일정리기",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI 앱: 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="app.ico",       # 아이콘이 있으면 주석 해제
)
