@echo off
if [%1]==[] (
    echo Usage: ppt_rasterize "Presentation.pptx"
    goto :eof
)

powershell -ExecutionPolicy Bypass "%~dp0ppt_rasterize.ps1" '%1'
