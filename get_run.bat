REM Git 설치 확인
git --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Git이 이미 설치되어 있습니다.
    git --version
) else (
    echo Git이 설치되어 있지 않습니다. Git 설치를 시작합니다.
    REM Git 설치 파일 다운로드
    curl -L -o git-installer.exe https://github.com/git-for-windows/git/releases/download/v2.30.0.windows.1/Git-2.30.0-64-bit.exe
    REM Git 설치 파일 실행
    start "" git-installer.exe /VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /COMPONENTS="icons,ext\reg\shellhere,assoc,assoc_sh"
    echo Git 설치가 완료될 때까지 잠시 기다려 주세요...
    timeout /t 30 /nobreak >nul
)

echo Git에 로그인한 후 아무 키나 누르세요...
pause >nul

set CURRENT_DIR=%cd%
set PEPO_DIR=%CURRENT_DIR%\auto

if not exist %PEPO_DIR% (
    echo 저장소 클론 중...
    cd %CURRENT_DIR%
    git clone https://github.com/donggyun112/automacro_excel.git auto
) else (
    echo 저장소 업데이트 중...
    cd %PEPO_DIR%
    git pull
    cd %CURRENT_DIR%
)

cd %PEPO_DIR%

REM run.bat 파일 실행
if exist run.bat (
    echo run.bat 파일 실행 중...
    call run.bat
) else (
    echo run.bat 파일이 존재하지 않습니다.
)

cd %CURRENT_DIR%