@echo off
setlocal enabledelayedexpansion

:: unregister the "Rasterize" command
powershell -executionpolicy bypass -file %~dp0unreg.ps1

:: remove the files
for %%f in (ppt_rasterize.ps1 unreg.ps1 install.cmd stories.txt README.md ppt_rasterize.py) do (
    if exist "%~dp0\%%f" (
        rm "%~dp0\%%f"
    )
)

:: count files left
set cnt=0
for %%f in (%~dp0) do set /a cnt+=1

:: remove the directory if it is empty
:: remove uninstall.cmd otherwise
if !cnt!==1 (
    echo Removing %~dp0
    cd ..
    rmdir %~dp0 /s /q
) else (
    rm %0
)
