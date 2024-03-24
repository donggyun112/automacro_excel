@echo off
chcp 65001 > nul

python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python이 이미 설치되어 있습니다.
    python --version
) else (
    echo Python이 설치되어 있지 않습니다. Microsoft Store에서 설치를 시작합니다.
    
    REM Microsoft Store 열기
    start "" "ms-windows-store://pdp/?productid=9P7QFQMJRFP7"
    
    echo Microsoft Store에서 Python 3 설치를 진행해 주세요.
    pause
)

REM 현재 디렉토리 설정
set CURRENT_DIR=%cd%

REM 가상 환경 디렉토리 설정
set VENV_DIR=%CURRENT_DIR%\myenv

REM 저장소 디렉토리 설정
set REPO_DIR=%VENV_DIR%\srcs

REM 가상 환경이 없으면 생성
if not exist %VENV_DIR% (
    echo 가상 환경 생성 중...
    python -m venv %VENV_DIR%
)

REM 저장소가 없으면 클론
if not exist %REPO_DIR% (
    echo 저장소 클론 중...
    cd %VENV_DIR%
    git clone https://github.com/donggyun112/work.git srcs
    cd %CURRENT_DIR%
) else (
    echo 저장소 업데이트 중...
    cd %REPO_DIR%
    git pull
    cd %CURRENT_DIR%
)

REM requirements.txt 파일 경로 설정
set REQUIREMENTS_FILE=%VENV_DIR%\requirements.txt

REM requirements.txt 파일이 없으면 생성
if not exist %REQUIREMENTS_FILE% (
    echo requirements.txt 파일 생성 중...
    echo openpyxl > %REQUIREMENTS_FILE%
    echo pandas >> %REQUIREMENTS_FILE%
    echo pillow >> %REQUIREMENTS_FILE%
)

REM 가상 환경 활성화
call %VENV_DIR%\Scripts\activate.bat

REM 필요한 라이브러리 설치
echo 라이브러리 설치 중...
pip install -r %REQUIREMENTS_FILE%

set REQUIREMENTS_FILE=%VENV_DIR%\test.bat

REM test.bat 파일이 없으면 생성
if not exist %REQUIREMENTS_FILE% (
    echo test.txt 파일 생성 중...
    echo @echo off > %REQUIREMENTS_FILE%
    echo call %VENV_DIR%\Scripts\deactivate.bat >> %REQUIREMENTS_FILE%
    echo python3 srcs\test.py >> %REQUIREMENTS_FILE%
    echo call %VENV_DIR%\Scripts\deactivate.bat >> %REQUIREMENTS_FILE%
)




REM 가상 환경 비활성화
call %VENV_DIR%\Scripts\deactivate.bat

echo 작업 완료!
pause