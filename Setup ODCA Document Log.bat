@echo off
title ODCA Document Log - Setup
echo.
echo   =============================================
echo     ODCA Document Log - Setting up...
echo     Please wait, this may take a few minutes.
echo   =============================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   Something went wrong. Please ask Mai for help.
    echo.
    pause
)
