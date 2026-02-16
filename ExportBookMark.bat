@echo off
REM Clears the screen for a fresh start
cls

REM Sets the window title
title Bookmark Export Utility v2.7

REM Display a friendly banner
echo.
echo =========================================
echo  Starting Bookmark Export Process...
echo =========================================
echo.

REM Runs the PowerShell script
REM "%~dp0" ensures the script path is correct regardless of where the BAT file is run from.
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0ExportBookMarks.ps1"

REM Line break for aesthetics
echo.
echo =========================================
echo  Export Process Complete.
echo =========================================
echo.

REM Pauses the window so the user can read the summary output from the PS script
pause