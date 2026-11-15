@echo off
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
echo ========================================================
echo Removing Daily Report 7:00 PM Windows Task...
echo ========================================================
python scheduler_manager.py unregister
echo.
pause

