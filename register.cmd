@echo off
for /f %%p in ('where powershell') do set "powershell=%%p"

for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Installing for PowerPoint.%%v
    reg add HKCU\Software\Classes\PowerPoint.%%v\shell\Rasterize\command /f /ve /d "\"%powershell%\" -executionpolicy bypass -file \"%~dp0ppt_rasterize.ps1\" \"^%%1\""
)
