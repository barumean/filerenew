@echo off
REM ============================================================
REM  파일 정리기(엑셀+PPT) - Windows 단일 EXE 빌드 스크립트
REM  사용법: 이 파일을 더블클릭하거나 명령창에서 build.bat 실행
REM  필요:   Python 3.8+ 설치 (설치 시 "Add Python to PATH" 체크 권장)
REM  결과물: dist\파일정리기.exe  (단일 실행 파일)
REM ============================================================
chcp 65001 >nul
setlocal

REM  이 배치 파일이 있는 폴더로 이동(어디서 실행하든 run_gui.py 를 찾도록)
cd /d "%~dp0"

REM  파이썬 실행기 자동 감지: 'py' 런처 우선, 없으면 'python'
set "PY=python"
where py >nul 2>&1 && set "PY=py"
%PY% --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python을 찾을 수 없습니다.
    echo        https://www.python.org 에서 설치 후 다시 실행하세요.
    echo        설치 시 "Add Python to PATH" 를 체크하면 편합니다.
    pause & exit /b 1
)

echo.
echo [1/4] 가상환경(.venv) 준비...
if not exist ".venv" (
    %PY% -m venv .venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패.
        pause & exit /b 1
    )
)

echo [2/4] 의존성 설치...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 의존성 설치 실패. 인터넷 연결을 확인하세요.
    pause & exit /b 1
)
python -m pip install pyinstaller
if errorlevel 1 (
    echo [오류] PyInstaller 설치 실패.
    pause & exit /b 1
)

echo [3/4] 이전 빌드 정리...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/4] EXE 빌드 (PyInstaller)...
python -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo [오류] 빌드 실패. 위 메시지를 확인하세요.
    pause & exit /b 1
)

if not exist "dist\파일정리기.exe" (
    echo [오류] 결과물(dist\파일정리기.exe)을 찾을 수 없습니다.
    pause & exit /b 1
)

echo.
echo ============================================================
echo  빌드 완료!  결과물: dist\파일정리기.exe
echo ============================================================
echo.
echo  - 파이썬 없는 PC에서도 더블클릭으로 실행됩니다.
echo  - 드래그 앤 드롭까지 포함되어 빌드됩니다(tkinterdnd2).
echo.
pause
