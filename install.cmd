@echo off
:: copy the distribution files to %USERPROFILE%\.ppt_rasterize
:: register the "Rasterize" shell command for PowerPoint presentations and slide shows

setlocal enabledelayedexpansion
for /f %%p in ('where powershell') do set "powershell=%%p"

if not exist !powershell! (
    echo Windows Powershell not found
    echo Please install Windows Powershell first.
    goto :eof
)

set "install_dir=%USERPROFILE%\.ppt_rasterize"
if not exist !install_dir! (
    mkdir !install_dir!
)
set "master_url=https://raw.github.com/utapyngo/ppt_rasterize/master"

for %%f in (ppt_rasterize.ps1 unreg.ps1 uninstall.cmd install.cmd) do (
    if exist "%~dp0\%%f" (
        copy "%~dp0\%%f" "!install_dir!"
    ) else (
        echo File "%%f" not found. Trying to download it from "!master_url!/%%f"
        powershell -command "(New-Object Net.WebClient).DownloadFile(\"!master_url!/%%f\", \"!install_dir!\%%f\")"
        if errorlevel 1 (
            echo Unable to download.
        )
    )
)

for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Installing for PowerPoint.%%v
    reg add HKCU\Software\Classes\PowerPoint.%%v\shell\Rasterize\command /f /ve /d "\"!powershell!\" -executionpolicy bypass -file \"!install_dir!\ppt_rasterize.ps1\" \"^%%1\""
)
