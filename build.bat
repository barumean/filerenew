@echo off
REM ============================================================
REM  파일 정리기(엑셀+PPT) - Windows 단일 EXE 빌드 스크립트
REM  사용법: 이 파일을 더블클릭하거나 명령창에서 build.bat 실행
REM  결과물: dist\파일정리기.exe
REM ============================================================
chcp 65001 >nul
setlocal

REM  이 배치 파일이 있는 폴더로 이동(어디서 실행하든 run_gui.py 를 찾도록)
cd /d "%~dp0"

echo.
echo [1/4] 가상환경(.venv) 준비...
if not exist ".venv" (
    python -m venv .venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패. Python이 설치되어 있는지 확인하세요.
        pause & exit /b 1
    )
)

echo [2/4] 의존성 설치...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 (
    echo [오류] 의존성 설치 실패.
    pause & exit /b 1
)

echo [3/4] 이전 빌드 정리...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/4] EXE 빌드 (PyInstaller)...
pyinstaller build.spec --noconfirm
if errorlevel 1 (
    echo [오류] 빌드 실패.
    pause & exit /b 1
)

echo.
echo ============================================================
echo  빌드 완료!  결과물: dist\파일정리기.exe
echo ============================================================
echo.
pause
