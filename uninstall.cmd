@echo off
setlocal enabledelayedexpansion

for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Uninstalling for PowerPoint.%%v
    reg delete HKCU\Software\Classes\PowerPoint.%%v\shell\Rasterize /f
)
