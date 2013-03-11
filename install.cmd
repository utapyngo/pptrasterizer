setlocal DisableDelayedExpansion

echo This will install %product_name% version %version% to you computer.
pause

for /f %%p in ('where powershell') do set "powershell=%%p"

if not exist %powershell% (
    echo Windows Powershell not found
    echo Please install Windows Powershell first.
    pause
    goto :eof
)

set "install_dir=%USERPROFILE%\.%short_name%\"

if exist "%install_dir%version.txt" (
    :: already installed
    set /p current_version=<"%install_dir%version.txt"
    setlocal EnableDelayedExpansion
    echo Your version: !current_version!
    echo Latest version: %version%
    if !current_version! GEQ %version% (
        echo You are already using the latest version of %product_name%.
        echo Press any key to exit the installer . . .
        pause>nul
        goto :eof
    ) else (
        echo Uninstalling previous version...
        call %install_dir%uninstall.cmd 2>nul
        pause
    )
    endlocal
)

if not exist "%install_dir%" (
    mkdir "%install_dir%"
)

:: only extract if we are not already there
for %%A in ("%~dp0") do for %%B in ("%install_dir%") do if %%~fA NEQ %%~fB (
    echo Copying files...
    type nul >%install_dir%files.txt
    for %%f in (%files%) do (
        call :get %%f > "%install_dir%\%%f"
        echo   %%f
        echo %%f>>%install_dir%files.txt
    )
    echo files.txt>>%install_dir%files.txt
)

:: add uninstall information to Programs and Features
set "uninstall_key=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\%short_name%"
reg add %uninstall_key% /f /v DisplayName /d "%product_name%" >nul
reg add %uninstall_key% /f /v InstallLocation /d %install_dir% >nul
reg add %uninstall_key% /f /t REG_DWORD /v NoModify /d 1 >nul
reg add %uninstall_key% /f /t REG_DWORD /v NoRepair /d 1 >nul
reg add %uninstall_key% /f /v UninstallString /d %install_dir%uninstall.cmd >nul
reg add %uninstall_key% /f /v DisplayVersion /d %version% >nul
:: register shell command
call "%install_dir%register.cmd"

:: thank for installing if only not updating
if not "%installer_url%" (
    start http://j.mp/pptrasterizer-installed
)

exit /b

:get
set "skip="
for /f "delims=:" %%N in ('findstr /x /n "::begin.%~1" "%~f0"') do if not defined skip set skip=%%N
set "end="
for /f "delims=:" %%N in ('findstr /x /n "::end.%~1" "%~f0"') do if %%N gtr %skip% if not defined end set end=%%N
for /f "skip=%skip% tokens=*" %%A in ('findstr /n "^" "%~f0"') do (
    for /f "delims=:" %%N in ("%%A") do if %%N geq %end% exit /b
    set "line=%%A"
    setlocal EnableDelayedExpansion
    echo(!line:*:=!
    endlocal
)
exit /b
