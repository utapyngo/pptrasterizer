@echo off

set "installer_url=https://dl.dropbox.com/u/62722148/ppt_rasterize-releases/setup_ppt_rasterize.bat"
set "install_dir=%~dp0"
set "installer_name=setup_ppt_rasterize.bat"

echo Downloading the installer...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%installer_url%\", \"%TEMP%\%installer_name%\")"

if errorlevel 1 (
    echo Unable to download the installer.
    goto :eof
)

call %TEMP%\%installer_name%

rm %TEMP%\%installer_name%
