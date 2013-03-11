@echo off

call %~dp0vars.cmd

set "installer_url=https://dl.dropbox.com/u/62722148/%short_name%-releases/%installer_name%"
set "install_dir=%~dp0"

echo Downloading the installer...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%installer_url%\", \"%TEMP%\%installer_name%\")"

if errorlevel 1 (
    echo Unable to download the installer. Check your Internet connection and firewall rules.
    pause
    goto :eof
)

call %TEMP%\%installer_name%

rm %TEMP%\%installer_name%
