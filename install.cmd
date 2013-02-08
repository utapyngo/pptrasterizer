@echo off
setlocal enabledelayedexpansion
for /f %%p in ('where powershell') do set "powershell=%%p"
if not exist !powershell! (
  echo Windows Powershell not found. Please install Windows Powershell first.
  goto :eof
)


for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Installing for PowerPoint.%%v
    reg add HKCR\PowerPoint.%%v\shell\Rasterize\command /f /ve /d "\"!powershell!\" -executionpolicy bypass -file \"%cd%\ppt_rasterize.ps1\" \"^%%1\""
)
