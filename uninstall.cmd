@echo off
setlocal EnableDelayedExpansion

call %~dp0vars.cmd

echo %product_name% is about to be removed from your computer.
pause

:: unregister the "Rasterize" command
powershell -executionpolicy bypass -file %~dp0unreg.ps1

:: remove from Programs and Features
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\%short_name%" /f >nul

:: remove the files
set "files="
for /f %%f in (files.txt) do (
    set "files=!files! %%f"
)
for %%f in (!files!) do (
    :: skip this file
    if %%f neq %~nx0 (
        if exist "%~dp0\%%f" (
            echo Removing %%f
            rm "%~dp0\%%f"
        )
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
    echo %product_name% has been uninstalled from your computer.
    rmdir %~dp0 /s /q
) else (
    echo %product_name% has been uninstalled from your computer.
    rm %0
)
