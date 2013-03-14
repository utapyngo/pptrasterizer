@echo off

call %~dp0vars.cmd

set "version_url=http://utapyngo.github.com/pptrasterizer/version.txt"
set "installer_url=http://utapyngo.github.com/pptrasterizer/%installer_name%"

echo Checking for the latest version online...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%version_url%\", \"%TEMP%\%short_name%_latest_version.txt\")"
if errorlevel 1 (
    echo Unable to check version online. Check your Internet connection and firewall rules.
    pause
    goto :eof
)
set /p latest_version=<"%TEMP%\%short_name%_latest_version.txt"
del "%TEMP%\%short_name%_latest_version.txt"

if %version% GEQ %latest_version% (
    echo You are already using the latest version of %product_name%.
    echo Press any key to exit . . .
    pause>nul
    goto :eof
)

echo Downloading the installer...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%installer_url%\", \"%TEMP%\%installer_name%\")"

if errorlevel 1 (
    echo Unable to download the installer. Check your Internet connection and firewall rules.
    pause
    goto :eof
)

call %TEMP%\%installer_name%

del %TEMP%\%installer_name%
