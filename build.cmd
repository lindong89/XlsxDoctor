@echo off
rem ============================================================
rem  XlsxDoctor build wrapper
rem  Double-click this file to build without dealing with
rem  PowerShell execution policy.
rem  Passes any arguments through to build.ps1, e.g.:
rem      build.cmd -DebugBuild
rem      build.cmd -OutDir D:\out
rem ============================================================

setlocal
cd /d "%~dp0"

rem Only pause when launched by double-click (cmd /c "...build.cmd"),
rem not when invoked from an already-open terminal.
set INTERACTIVE=
echo %cmdcmdline% | findstr /i /c:"%~nx0" >nul && set INTERACTIVE=1

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1" %*
set RC=%ERRORLEVEL%

echo.
if not "%RC%"=="0" (
    echo [BUILD FAILED] exit code %RC%
) else (
    echo [BUILD OK]
)

if defined INTERACTIVE pause
exit /b %RC%
