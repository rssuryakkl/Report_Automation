@echo off
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
echo ========================================================
echo Testing Daily Report Email Send (Latest Row)...
echo ========================================================
python send_today_report.py --latest --force
echo.
pause

