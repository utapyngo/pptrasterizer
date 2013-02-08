@echo off
setlocal enabledelayedexpansion
for /f %%p in ('where powershell') do set "powershell=%%p"

if not exist !powershell! (
    echo Windows Powershell not found
    echo Please install Windows Powershell first.
    goto :eof
)

if not exist %cd%\ppt_rasterize.ps1 (
    echo %cd%\ppt_rasterize.ps1 not found.
    echo Trying to download...
    powershell -command "(New-Object Net.WebClient).DownloadFile('https://raw.github.com/utapyngo/ppt_rasterize/master/ppt_rasterize.ps1', 'ppt_rasterize.ps1')"
    if errorlevel 1 (
        echo Could not download.
        echo Please unpack everything, not just install.cmd.
        goto :eof
    )
)

for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Installing for PowerPoint.%%v
    reg add HKCU\Software\Classes\PowerPoint.%%v\shell\Rasterize\command /f /ve /d "\"!powershell!\" -executionpolicy bypass -file \"%cd%\ppt_rasterize.ps1\" \"^%%1\""
)
