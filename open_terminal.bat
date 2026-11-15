@echo off
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
cd /d "%~dp0"
title Daily Report Automation - Command Terminal
echo =======================================================
echo Daily Report Automation Terminal Ready
echo Python Version:
python --version
echo =======================================================
echo.
cmd.exe /k
