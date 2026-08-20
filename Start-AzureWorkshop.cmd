@echo off
setlocal

title Azure WAF/CAF Workshop
cd /d "%~dp0"

where pwsh.exe >nul 2>&1
if errorlevel 1 (
    echo PowerShell 7 is required but pwsh.exe was not found.
    echo Install it from https://aka.ms/powershell and run this launcher again.
    echo.
    pause
    exit /b 1
)

echo Starting the Azure WAF/CAF Workshop...
echo.

pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0powershell\Launch-AzureWorkshop.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo The workshop ended with exit code %EXIT_CODE%.
) else (
    echo The workshop completed successfully.
    if exist "%~dp0AzureWorkshop\07_Dashboard\WAF_Dashboard.html" (
        echo Opening the consolidated dashboard...
        start "" "%~dp0AzureWorkshop\07_Dashboard\WAF_Dashboard.html"
    )
)

echo.
echo Press any key to close this window.
pause >nul
exit /b %EXIT_CODE%