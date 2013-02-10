setlocal DisableDelayedExpansion
for /f %%p in ('where powershell') do set "powershell=%%p"

if not exist %powershell% (
    echo Windows Powershell not found
    echo Please install Windows Powershell first.
    goto :eof
)

set "install_dir=%USERPROFILE%\.ppt_rasterize\"
if not exist %install_dir% (
    mkdir %install_dir%
)

:: only copy if we are not already copied
for %%A in ("%~dp0") do for %%B in ("%install_dir%") do if %%~fA NEQ %%~fB (
    for %%f in (%files%) do (
        if exist "%~dp0\%%f" (
            echo Copying %%f
            copy "%~dp0\%%f" "%install_dir%"
        ) else (
            call :get %%f > "%install_dir%\%%f"
        )
        echo %%f>>%install_dir%files.txt
    )
    echo files.txt>>%install_dir%files.txt
)

:: register shell command
call "%install_dir%register.cmd"
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

