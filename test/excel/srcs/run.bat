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
set VENV_DIR=%CURRENT_DIR%\myenv

REM 저장소 디렉토리 설정
set REPO_DIR=%CURRENT_DIR%\myenv\srcs

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

REM test.py 파일 확인 및 이동
if exist %CURRENT_DIR%\test.py (
    echo test.py 파일을 srcs 폴더로 이동 중...
    move %CURRENT_DIR%\test.py %REPO_DIR%
) else (
    echo test.py 파일이 없습니다. 다운로드해 주세요.
    echo test.py 파일을 %REPO_DIR% 경로에 다운로드한 후 스크립트를 다시 실행해 주세요.
    pause
    exit
)


REM requirements.txt 파일 경로 설정
set REQUIREMENTS_FILE=%VENV_DIR%\requirements.txt

REM requirements.txt 파일이 없으면 생성
if not exist %REQUIREMENTS_FILE% (
    echo requirements.txt 파일 생성 중...
    echo openpyxl > %REQUIREMENTS_FILE%
    echo pandas >> %REQUIREMENTS_FILE%
    echo pillow >> %REQUIREMENTS_FILE%
    echo tabulate >> %REQUIREMENTS_FILE%
	echo tkintertable >> %REQUIREMENTS_FILE%
	echo wcwidth >> %REQUIREMENTS_FILE%
)

REM 가상 환경 활성화
call %VENV_DIR%\Scripts\activate.bat

REM 필요한 라이브러리 설치
echo 라이브러리 설치 중...
pip install -r %REQUIREMENTS_FILE%

REM test.bat 파일 경로 설정
set TEST_BAT_FILE=%VENV_DIR%\test.bat

REM test.bat 파일 생성
echo @echo off > %TEST_BAT_FILE%
echo call %VENV_DIR%\Scripts\activate.bat >> %TEST_BAT_FILE%
echo python %REPO_DIR%\test.py >> %TEST_BAT_FILE%
echo pause >> %TEST_BAT_FILE%
echo call %VENV_DIR%\Scripts\deactivate.bat >> %TEST_BAT_FILE%

REM 가상 환경 비활성화
call %VENV_DIR%\Scripts\deactivate.bat

echo 작업 완료!
pause