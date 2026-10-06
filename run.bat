@echo off
cd /d "%~dp0"
echo ============================================
echo   Google Maps Lead Crawler
echo ============================================
echo.
if not exist ".venv" goto SETUP
goto RUN
:SETUP
echo Creating venv...
python -m venv .venv
call .venv\Scripts\activate.bat
echo Installing dependencies...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
echo Installing browser...
playwright install chromium
echo.
echo Setup done! Run again to start.
pause
exit /b
:RUN
call .venv\Scripts\activate.bat
python maps_crawler.py
pause
