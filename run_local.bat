@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 미소로 일일 보고 앱을 시작합니다... 잠시 후 브라우저가 열립니다.
echo (창을 닫으면 서버가 종료됩니다.)
start "" cmd /c "timeout /t 4 >nul && start http://localhost:8501/?branch=%%EB%%B6%%84%%EB%%8B%%B9"
python -m streamlit run app.py --server.port 8501 --server.headless true
