@echo off
REM ============================================================
REM CloudFuze Compensation Tool — double-click to push to GitHub
REM ============================================================
REM Just double-click this file. A PowerShell window opens, does
REM all the work, then prompts you to sign in to GitHub once via
REM your browser. After that, you're done.
REM ============================================================

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0push_to_github.ps1"

REM Keep window open if PowerShell exits with an error
if errorlevel 1 (
    echo.
    echo Something went wrong. Read the message above.
    pause
)
