@echo off

chcp 65001 > nul

REM Python 설치 확인
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python이 이미 설치되어 있습니다.
    python --version
) else (
    echo Python이 설치되어 있지 않습니다. Microsoft Store에서 설치를 시작합니다.
    REM Microsoft Store 열기
    start "" "ms-windows-store://pdp/?productid=9PJPW5LDXLZ5"
    echo Microsoft Store에서 Python 3.12 설치를 진행해 주세요.
    REM 설치 완료 대기
    echo Python 설치 완료까지 대기 중...
    :check_python
    timeout /t 5 /nobreak >nul
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        goto check_python
    )
    echo Python 설치가 완료되었습니다.
)

REM 현재 디렉토리 설정
set CURRENT_DIR=%cd%

REM 가상 환경 디렉토리 설정
set VENV_DIR=%CURRENT_DIR%\auto

REM 저장소 디렉토리 설정
set REPO_DIR=%CURRENT_DIR%\auto\srcs

REM 가상 환경이 없으면 생성
if not exist %VENV_DIR% (
    echo 가상 환경 생성 중...
    python -m venv %VENV_DIR%
)

REM 저장소 디렉토리 생성
if not exist %REPO_DIR% (
    echo 저장소 디렉토리 생성 중...
    mkdir %REPO_DIR%
)

set SRC_FILE=ver3.py
set SRC_FILE2=compare.py
set SRC_FUNC=func.py
set SRC_IMPORT=imports.py
set SRC_CUSTOM=custom.py
set SETTING_DIR=setting

REM ver3.py 파일 확인 및 이동
if exist %CURRENT_DIR%\%SRC_FILE% (
    echo %SRC_FILE% 파일을 srcs 폴더로 이동 중...
    echo %SRC_FILE2% 파일을 srcs 폴더로 이동 중...
    echo %SRC_IMPORT% 파일을 srcs 폴더로 이동 중...
    echo %SRC_FUNC% 파일을 srcs 폴더로 이동 중...
    move %CURRENT_DIR%\%SRC_FILE% %REPO_DIR%
    move %CURRENT_DIR%\%SRC_FILE2% %REPO_DIR%
    move %CURRENT_DIR%\%SRC_IMPORT% %REPO_DIR%
    move %CURRENT_DIR%\%SRC_FUNC% %REPO_DIR%
) else (
    echo %SRC_FILE% 파일이 없습니다. 다운로드해 주세요.
    echo %SRC_FILE% 파일을 %REPO_DIR% 경로에 다운로드한 후 스크립트를 다시 실행해 주세요.
    pause
    exit
)

REM setting 디렉토리 이동
if exist %CURRENT_DIR%\%SETTING_DIR% (
    echo %SETTING_DIR% 디렉토리를 srcs 폴더로 이동 중...
    move %CURRENT_DIR%\%SETTING_DIR% %REPO_DIR%
) else (
    echo %SETTING_DIR% 디렉토리가 없습니다. 생성해 주세요.
    mkdir %CURRENT_DIR%\%SETTING_DIR%
    echo %SETTING_DIR% 디렉토리를 %REPO_DIR% 경로로 이동시킨 후 스크립트를 다시 실행해 주세요.
    pause
    exit
)

REM requirements.txt 파일 경로 설정
set REQUIREMENTS_FILE=%VENV_DIR%\requirements.txt

REM requirements.txt 파일이 없으면 생성
if not exist %REQUIREMENTS_FILE% (
    echo requirements.txt 파일 생성 중...
    echo openpyxl > %REQUIREMENTS_FILE%
    echo xlwings >> %REQUIREMENTS_FILE%
    echo PyQt5 >> %REQUIREMENTS_FILE%
    echo pillow >> %REQUIREMENTS_FILE%
    echo pandas >> %REQUIREMENTS_FILE%
)

REM 가상 환경 활성화
call %VENV_DIR%\Scripts\activate.bat

REM 필요한 라이브러리 설치
echo 라이브러리 설치 중...
pip install -r %REQUIREMENTS_FILE%

REM custom.py 파일을 openpyxl/packaging/ 디렉토리로 이동
echo custom.py 파일을 openpyxl/packaging/ 디렉토리로 이동 중...
set OPENPYXL_DIR=%VENV_DIR%\Lib\site-packages\openpyxl\packaging
if not exist %OPENPYXL_DIR% (
    echo openpyxl 디렉토리가 존재하지 않습니다. 스크립트를 종료합니다.
    pause
    exit
)
move %CURRENT_DIR%\%SRC_CUSTOM% %OPENPYXL_DIR%

REM test.bat 파일 경로 설정
set TEST_BAT_FILE=%VENV_DIR%\activate.bat

REM actiave.bat 파일 생성
echo @echo off > %TEST_BAT_FILE%
echo call %VENV_DIR%\Scripts\activate.bat >> %TEST_BAT_FILE%
echo python %REPO_DIR%\%SRC_FILE% >> %TEST_BAT_FILE%
echo pause >> %TEST_BAT_FILE%
echo call %VENV_DIR%\Scripts\deactivate.bat >> %TEST_BAT_FILE%


REM 가상 환경 비활성화
call %VENV_DIR%\Scripts\deactivate.bat

echo 작업 완료!
pause