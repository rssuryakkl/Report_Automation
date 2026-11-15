@echo off
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
echo ========================================================
echo Registering Daily Report 7:00 PM Windows Task...
echo ========================================================
python scheduler_manager.py
echo.
pause

