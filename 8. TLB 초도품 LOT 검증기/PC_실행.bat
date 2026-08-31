@echo off
chcp 65001 > nul
cd /d "%~dp0"

where python > nul 2>&1
if errorlevel 1 (
    echo [실행 오류] Python을 찾을 수 없습니다.
    echo index.html 파일을 직접 더블 클릭해서 실행하세요.
    pause
    exit /b 1
)

echo TLB 초도품 LOT 검증기를 실행합니다.
echo.
echo 이 검은 창을 닫으면 앱 실행도 종료됩니다.
echo 브라우저가 열리지 않으면 아래 주소를 직접 입력하세요.
echo http://127.0.0.1:8766/?v=1
echo.

start "" "http://127.0.0.1:8766/?v=1"
python -m http.server 8766 --bind 127.0.0.1
